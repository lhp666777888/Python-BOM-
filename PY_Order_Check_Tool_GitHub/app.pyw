# -*- coding: utf-8 -*-
"""
Python 3.8 兼容版：按“父项订单数量”核算 BOM 子项需求与欠料

核心规则（本版最重要的修改）：
1. BOM中的“额定消耗”是【该父项/成品的单机用量】，绝对不能先把不同BOM的额定消耗相加。
2. 先分别计算每个父项的需求：
       子项需求 = 单机用量 × 该父项订单数量
3. 只有在第2步计算完成后，才允许按“子项物料编码”汇总总需求：
       用量合计 = Σ(各父项的 单机用量 × 订单数量)
4. 欠料数量：
       欠料合计 = MAX(用量合计 - 库存数量, 0)
5. 采购缺口：
       采购缺口 = MAX(欠料合计 - EAS在途, 0)
6. EAS申请去重：
       已经转成采购订单的申请数量，不再计入EAS申请；避免和EAS在途重复。
       优先按采购订单中的“源单/申请单号”精确关联；若导出表没有源单号，
       则按“同物料编码 + 申请日期不晚于采购日期 + FIFO”进行保守匹配。
7. 申请缺口：
       申请缺口 = MAX(采购缺口 - EAS未转单申请, 0)

示例：
父项 14-2302-0412-00 订单数量=500
其子项 14-2301-0055 单机用量=380
则需求数量 = 380 × 500 = 190000
如果即时库存=5000，则欠料 = 190000 - 5000 = 185000

使用方式：
1. 把本脚本、BOM文件.zip、线束资料.zip、PY订单核料表模板放在同一文件夹；
2. 双击或在 PyCharm / VS Code / IDLE 中直接运行；
3. 程序会弹出“订单数量输入”窗口；
4. 在每个父项后输入订单数量，例如 14-2302-0412-00 输入 500；
5. 点击“开始核料”，程序自动生成结果文件。

依赖：openpyxl
如果本机未安装：pip install openpyxl
"""

import os
import re
import sys
import threading
import traceback
import tempfile
import zipfile
from collections import OrderedDict, defaultdict
from copy import copy
from datetime import date, datetime
from pathlib import Path

try:
    from openpyxl import load_workbook
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.worksheet.views import Selection
except ImportError:
    raise SystemExit(
        "未安装 openpyxl。请先安装 openpyxl 后再运行。\n"
        "Python 3.8 可使用：pip install openpyxl"
    )


# ============================================================
# 一、运行目录 / 资源目录
# ============================================================
def get_app_dir():
    """源码运行时=脚本目录；打包EXE后=EXE所在目录。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def get_resource_dir():
    """PyInstaller资源目录；源码运行时仍为脚本目录。"""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return get_app_dir()


BASE_DIR = get_app_dir()
RESOURCE_DIR = get_resource_dir()

# 下面四个变量仅作为默认值；桌面界面中都可以点击“选择”重新指定。
BOM_ZIP_PATH = BASE_DIR / "BOM文件.zip"
SOURCE_ZIP_PATH = BASE_DIR / "线束资料.zip"
_template_in_resource = RESOURCE_DIR / "py订单核料表OPMC-深圳-2026-9-04.xlsx"
TEMPLATE_XLSX_PATH = _template_in_resource if _template_in_resource.exists() else (BASE_DIR / "py订单核料表OPMC-深圳-2026-9-04.xlsx")
OUTPUT_XLSX_PATH = BASE_DIR / "PY订单核料表_核料结果.xlsx"


# ============================================================
# 二、订单数量配置
# ============================================================
# True：运行时弹窗输入订单数量（推荐）
# False：不弹窗，直接使用 DEFAULT_ORDER_QTY 中的数量
USE_GUI_ORDER_INPUT = True

# 弹窗里的默认值，也可以直接在这里修改。
# 例如本次做 14-2302-0412-00 共 500 条，另一个BOM暂时不做：
DEFAULT_ORDER_QTY = {
    "14-2302-0412-00": 500,
    "14-2302-0471-00": 0,
}

# 自动测试/批处理时可设置环境变量 PY_ORDER_NO_GUI=1 跳过弹窗。
# 普通用户无需理会。


# ============================================================
# 三、模板设置
# ============================================================
OUTPUT_SHEET_NAME = "总表"
START_ROW = 6
MAX_CLEAR_ROW = 5000
ORDER_SHEET_NAME = "订单输入"
DETAIL_SHEET_NAME = "BOM需求明细"
EAS_RELATION_SHEET_NAME = "EAS关联明细"

# 当采购订单导出表没有“源申请单号/源单编号”等字段时，
# 是否按“同物料 + 日期先后 + FIFO”推定申请与采购订单的关联。
# True 推荐：可以避免同一申请既出现在EAS申请又出现在EAS在途。
FALLBACK_MATCH_APPLICATION_TO_PO = True


# ============================================================
# 四、通用工具
# ============================================================
def text(value):
    """转为去除首尾空格的字符串。"""
    if value is None:
        return ""
    return str(value).strip()


def number(value):
    """尽量转为浮点数，无法转换则返回0。"""
    if value is None or value == "":
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace(",", "").strip())
    except (TypeError, ValueError):
        return 0.0


def clean_num(value):
    """整数型浮点数转整数，便于显示。"""
    if value is None:
        return None
    try:
        v = float(value)
    except (TypeError, ValueError):
        return value
    if v.is_integer():
        return int(v)
    return v


def date_value(value):
    """兼容Excel日期对象和常见文本日期。"""
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day)

    s = str(value).strip()
    for fmt in (
        "%Y/%m/%d %H:%M:%S",
        "%Y/%m/%d",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d",
        "%Y.%m.%d %H:%M:%S",
        "%Y.%m.%d",
    ):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def is_closed_status(value):
    """判断EAS单据是否已关闭/作废。"""
    s = text(value)
    if not s:
        return False

    # 这些状态不是关闭状态
    if s in ("未关闭", "正常", "审核中", "已审核", "已提交", "已审批", "生效"):
        return False

    return any(
        keyword in s
        for keyword in ("已关闭", "作废", "取消", "驳回", "终止", "结算", "完结", "关闭")
    )


def check_file(path, description):
    if not path.exists():
        raise FileNotFoundError(
            "%s不存在：%s\n请检查脚本顶部的路径配置。" % (description, path)
        )


def get_header_index(header_row):
    """返回 {表头名称: 列下标}，下标从0开始。"""
    result = {}
    for i, value in enumerate(header_row):
        name = text(value)
        if name and name not in result:
            result[name] = i
    return result


def read_xlsx_rows(path):
    """读取Excel第一个工作表的数据。"""
    wb = load_workbook(str(path), read_only=True, data_only=True)
    try:
        ws = wb.worksheets[0]
        return list(ws.iter_rows(values_only=True))
    finally:
        wb.close()


def xlsx_files(input_path, temp_dirs):
    """支持zip、文件夹或单个xlsx。"""
    input_path = Path(input_path)

    if input_path.is_dir():
        return sorted(input_path.rglob("*.xlsx"))

    if input_path.suffix.lower() == ".zip":
        td = tempfile.TemporaryDirectory(prefix="py_order_check_")
        temp_dirs.append(td)
        with zipfile.ZipFile(str(input_path), "r") as zf:
            zf.extractall(td.name)
        return sorted(Path(td.name).rglob("*.xlsx"))

    if input_path.suffix.lower() == ".xlsx":
        return [input_path]

    raise ValueError("不支持的输入文件：%s" % input_path)


# ============================================================
# 五、BOM处理
# 关键：只在【同一个父项内部】合并重复子项；绝不跨父项合并单机用量
# ============================================================
def _decode_hash_u(value):
    """兼容部分ZIP工具把中文文件名显示成 #U4e2d 形式。"""
    s = text(value)
    if not s:
        return ""

    def repl(match):
        try:
            return chr(int(match.group(1), 16))
        except Exception:
            return match.group(0)

    return re.sub(r"#U([0-9A-Fa-f]{4,6})", repl, s)


def _infer_parent_name_from_filename(path, parent_code):
    """当BOM缺少“物料名称”列时，尝试从文件名补父项名称。"""
    try:
        stem = _decode_hash_u(Path(path).stem)
        pos = stem.find(parent_code)
        if pos >= 0:
            tail = stem[pos + len(parent_code):]
            tail = re.sub(r"^[\s_\-—–]+", "", tail).strip()
            if tail:
                return tail
    except Exception:
        pass
    return ""


def collect_bom_details(files, diagnostics=None):
    """
    返回：
      parents: OrderedDict
          父项编码 -> {code, name}
      details: list
          每一条表示“一个父项中的一个子项”的单机用量。

    兼容规则：
    - “父项物料编码、物料编码、额定消耗、计量单位”是核料必需列；
    - “物料名称、规格型号”允许缺失，不再因此整张BOM被跳过；
    - 若缺“物料名称”，父项名称优先从文件名推断，子项名称再从其它BOM同编码回填；
    - 若同一子项在同一个父项BOM里出现多次，会在该父项内部相加；
    - 不同父项之间的单机用量绝不相加。
    """
    if diagnostics is None:
        diagnostics = []

    parents = OrderedDict()
    detail_map = OrderedDict()   # key=(parent_code, child_code)
    name_by_code = {}

    for path in files:
        rows = read_xlsx_rows(path)
        if not rows:
            diagnostics.append("跳过空BOM：%s" % path.name)
            continue

        idx = get_header_index(rows[0])

        # 真正参与核料的硬性表头。物料名称/规格型号属于展示字段，不应导致整张BOM失效。
        hard_required = [
            "父项物料编码",
            "物料编码",
            "额定消耗",
            "计量单位",
        ]
        missing_hard = [name for name in hard_required if name not in idx]
        if missing_hard:
            msg = "跳过BOM（缺少核料必要表头：%s）：%s" % ("、".join(missing_hard), path.name)
            print(msg)
            diagnostics.append(msg)
            continue

        has_name = "物料名称" in idx
        has_spec = "规格型号" in idx
        file_parents = []

        for row in rows[1:]:
            parent = text(row[idx["父项物料编码"]]) if idx["父项物料编码"] < len(row) else ""
            code = text(row[idx["物料编码"]]) if idx["物料编码"] < len(row) else ""

            name = ""
            if has_name and idx["物料名称"] < len(row):
                name = text(row[idx["物料名称"]])

            spec = ""
            if has_spec and idx["规格型号"] < len(row):
                spec = text(row[idx["规格型号"]])

            unit = text(row[idx["计量单位"]]) if idx["计量单位"] < len(row) else ""

            if not parent:
                continue

            if parent not in file_parents:
                file_parents.append(parent)

            if name and code:
                name_by_code[code] = name

            # 记录父项。即使没有“物料名称”列，也必须识别父项。
            if parent not in parents:
                parents[parent] = {
                    "code": parent,
                    "name": _infer_parent_name_from_filename(path, parent),
                }

            if code == parent:
                if name:
                    parents[parent]["name"] = name
                    name_by_code[parent] = name
                elif not parents[parent].get("name"):
                    parents[parent]["name"] = _infer_parent_name_from_filename(path, parent)
                continue

            # 排除空行，只保留子项
            if not code:
                continue

            unit_qty = number(row[idx["额定消耗"]]) if idx["额定消耗"] < len(row) else 0.0
            if unit_qty == 0:
                continue

            key = (parent, code)
            if key not in detail_map:
                detail_map[key] = {
                    "parent": parent,
                    "parent_name": parents[parent].get("name", ""),
                    "code": code,
                    "name": name,
                    "spec": spec,
                    "unit": unit,
                    "unit_qty": 0.0,
                    "source_files": [],
                }
            else:
                # 同一父项同一子项出现多次时，优先保留非空展示字段。
                if name and not detail_map[key].get("name"):
                    detail_map[key]["name"] = name
                if spec and not detail_map[key].get("spec"):
                    detail_map[key]["spec"] = spec
                if unit and not detail_map[key].get("unit"):
                    detail_map[key]["unit"] = unit

            # 只在同一个父项内部累加，绝不跨父项累加
            detail_map[key]["unit_qty"] += unit_qty
            if path.name not in detail_map[key]["source_files"]:
                detail_map[key]["source_files"].append(path.name)

        if not has_name:
            msg = "兼容导入BOM（缺少“物料名称”列，未跳过）：%s" % path.name
            if file_parents:
                msg += "；识别父项：%s" % "、".join(file_parents)
            diagnostics.append(msg)

    # 全部BOM读完以后，用其它BOM里相同物料编码的名称回填缺失名称。
    for key, item in detail_map.items():
        p = item["parent"]
        if p in parents:
            if not item.get("parent_name"):
                item["parent_name"] = parents[p].get("name", "")
        if not item.get("name"):
            item["name"] = name_by_code.get(item["code"], "")
        if not item.get("name"):
            item["name"] = "（BOM未提供物料名称）"

    return parents, list(detail_map.values())


# ============================================================
# 六、订单数量输入
# ============================================================
def default_order_qty(parents):
    """按父项生成默认订单数量。"""
    result = OrderedDict()
    for code in parents.keys():
        result[code] = max(number(DEFAULT_ORDER_QTY.get(code, 0)), 0.0)
    return result


def ask_order_quantities(parents):
    """
    使用 tkinter 输入所有父项订单数量。

    支持：
    - BOM很多时滚动浏览；
    - 按父项编码/名称搜索；
    - 本次不生产的保持0；
    - 若系统没有tkinter，则使用DEFAULT_ORDER_QTY。
    """
    defaults = default_order_qty(parents)

    if os.environ.get("PY_ORDER_NO_GUI") == "1" or not USE_GUI_ORDER_INPUT:
        return defaults

    try:
        import tkinter as tk
        from tkinter import messagebox
    except ImportError:
        print("未检测到 tkinter，改用代码顶部 DEFAULT_ORDER_QTY。")
        return defaults

    result_holder = {"value": None}
    root = tk.Tk()
    root.title("PY订单核料 - 输入订单数量")
    root.geometry("900x650")
    root.minsize(760, 480)

    try:
        root.attributes("-topmost", True)
    except Exception:
        pass

    top = tk.Frame(root)
    top.pack(fill="x", padx=12, pady=(12, 6))

    tk.Label(
        top,
        text="请输入各父项/线束订单数量；本次不生产的保持0。BOM很多时可搜索并滚动。",
        font=("Microsoft YaHei", 11, "bold"),
        anchor="w",
    ).pack(fill="x", pady=(0, 8))

    search_row = tk.Frame(top)
    search_row.pack(fill="x")
    tk.Label(search_row, text="搜索BOM：").pack(side="left")
    search_var = tk.StringVar()
    search_entry = tk.Entry(search_row, textvariable=search_var)
    search_entry.pack(side="left", fill="x", expand=True, padx=(4, 8))
    count_var = tk.StringVar(value="共 %d 个BOM" % len(parents))
    tk.Label(search_row, textvariable=count_var).pack(side="right")

    table_header = tk.Frame(root)
    table_header.pack(fill="x", padx=12)
    tk.Label(table_header, text="父项物料编码", width=24, anchor="w", font=("Microsoft YaHei", 9, "bold")).pack(side="left")
    tk.Label(table_header, text="物料名称", width=46, anchor="w", font=("Microsoft YaHei", 9, "bold")).pack(side="left")
    tk.Label(table_header, text="订单数量", width=14, anchor="e", font=("Microsoft YaHei", 9, "bold")).pack(side="left")

    body = tk.Frame(root)
    body.pack(fill="both", expand=True, padx=12, pady=(2, 6))

    canvas = tk.Canvas(body, highlightthickness=0)
    scrollbar = tk.Scrollbar(body, orient="vertical", command=canvas.yview)
    inner = tk.Frame(canvas)
    inner_id = canvas.create_window((0, 0), window=inner, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    def on_inner_configure(event=None):
        canvas.configure(scrollregion=canvas.bbox("all"))

    def on_canvas_configure(event):
        try:
            canvas.itemconfigure(inner_id, width=event.width)
        except Exception:
            pass

    inner.bind("<Configure>", on_inner_configure)
    canvas.bind("<Configure>", on_canvas_configure)

    variables = OrderedDict()
    row_widgets = OrderedDict()

    for code, info in parents.items():
        row_frame = tk.Frame(inner)
        row_frame.pack(fill="x", pady=2)

        tk.Label(row_frame, text=code, width=24, anchor="w").pack(side="left")
        tk.Label(row_frame, text=info.get("name", ""), width=46, anchor="w").pack(side="left")

        default_value = clean_num(defaults.get(code, 0))
        var = tk.StringVar(value=str(default_value))
        entry = tk.Entry(row_frame, textvariable=var, width=14, justify="right")
        entry.pack(side="left", padx=(6, 8))

        variables[code] = var
        row_widgets[code] = (row_frame, info.get("name", ""))

    def apply_filter(*args):
        keyword = search_var.get().strip().lower()
        shown = 0
        for code, pair in row_widgets.items():
            frame, name = pair
            matched = (not keyword) or (keyword in code.lower()) or (keyword in text(name).lower())
            if matched:
                if not frame.winfo_ismapped():
                    frame.pack(fill="x", pady=2)
                shown += 1
            else:
                frame.pack_forget()
        count_var.set("显示 %d / 共 %d 个BOM" % (shown, len(parents)))
        root.after_idle(on_inner_configure)

    try:
        search_var.trace_add("write", apply_filter)
    except AttributeError:
        search_var.trace("w", apply_filter)

    def on_mousewheel(event):
        try:
            if event.delta:
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except Exception:
            pass

    canvas.bind_all("<MouseWheel>", on_mousewheel)

    button_bar = tk.Frame(root)
    button_bar.pack(fill="x", padx=12, pady=(4, 12))

    def on_submit():
        values = OrderedDict()
        for code, var in variables.items():
            raw = var.get().strip()
            if raw == "":
                qty = 0.0
            else:
                try:
                    qty = float(raw.replace(",", ""))
                except ValueError:
                    messagebox.showerror("输入错误", "%s 的订单数量不是有效数字：%s" % (code, raw))
                    return
            if qty < 0:
                messagebox.showerror("输入错误", "%s 的订单数量不能小于0。" % code)
                return
            values[code] = qty

        if not any(qty > 0 for qty in values.values()):
            messagebox.showerror("输入错误", "至少需要一个父项的订单数量大于0。")
            return

        result_holder["value"] = values
        try:
            canvas.unbind_all("<MouseWheel>")
        except Exception:
            pass
        root.destroy()

    def on_cancel():
        result_holder["value"] = None
        try:
            canvas.unbind_all("<MouseWheel>")
        except Exception:
            pass
        root.destroy()

    tk.Button(button_bar, text="开始核料", width=14, command=on_submit).pack(side="right", padx=(8, 0))
    tk.Button(button_bar, text="取消", width=10, command=on_cancel).pack(side="right")

    root.protocol("WM_DELETE_WINDOW", on_cancel)
    search_entry.focus_set()
    root.mainloop()

    if result_holder["value"] is None:
        raise RuntimeError("用户取消了订单数量输入。")

    return result_holder["value"]


# ============================================================
# 七、根据订单数量计算物料需求
# 关键：先 父项×单机用量，再按子项编码汇总
# ============================================================
def build_material_requirements(details, order_qty_map):
    """
    计算最终材料需求。

    返回：
      materials: OrderedDict，按子项物料编码汇总后的结果
      active_details: 只保留订单数量>0的BOM需求明细
    """
    materials = OrderedDict()
    active_details = []

    for item in details:
        parent = item["parent"]
        order_qty = max(number(order_qty_map.get(parent, 0)), 0.0)
        if order_qty <= 0:
            continue

        unit_qty = number(item["unit_qty"])
        required_qty = unit_qty * order_qty

        detail_row = dict(item)
        detail_row["order_qty"] = order_qty
        detail_row["required_qty"] = required_qty
        active_details.append(detail_row)

        code = item["code"]
        if code not in materials:
            materials[code] = {
                "code": code,
                "name": item["name"],
                "spec": item["spec"],
                "unit": item["unit"],
                "required_qty": 0.0,
                "contributions": [],
            }

        # 注意：这里累加的是“已乘订单数量后的需求量”，不是累加单机用量
        materials[code]["required_qty"] += required_qty
        materials[code]["contributions"].append({
            "parent": parent,
            "parent_name": item.get("parent_name", ""),
            "order_qty": order_qty,
            "unit_qty": unit_qty,
            "required_qty": required_qty,
        })

    return materials, active_details


def get_quota_display(contributions):
    """
    返回 (用于V列显示的单机用量, 是否可作为单一数值用于“套数”换算)。

    - 只有一个有效父项：直接显示该父项单机用量。
    - 多个有效父项且单机用量完全相同：显示共同单机用量。
    - 多个有效父项且单机用量不同：显示“父项:单机用量”说明，套数列留空。
    """
    if not contributions:
        return "", None

    quotas = [number(x["unit_qty"]) for x in contributions]
    unique = []
    for q in quotas:
        if not any(abs(q - old) < 1e-9 for old in unique):
            unique.append(q)

    if len(unique) == 1:
        return clean_num(unique[0]), unique[0]

    parts = []
    for x in contributions:
        parts.append("%s:%s" % (x["parent"], clean_num(x["unit_qty"])))
    return "；".join(parts), None


# ============================================================
# 八、EAS资料处理
# ============================================================
SOURCE_DOC_HEADERS = (
    "源单编号", "源单单号", "来源单据编号", "来源单号",
    "采购申请单号", "申请单号", "关联申请单号", "关联申请单",
    "源申请单号", "源申请单", "来源申请单号", "来源申请单",
)


def is_invalid_document_status(value):
    """作废/取消/驳回/终止的单据不参与申请-订单关联。"""
    s = text(value)
    if not s:
        return False
    return any(k in s for k in ("作废", "取消", "驳回", "终止"))


def first_value(row, idx, names):
    """从多个可能的表头中取第一个非空值。"""
    for name in names:
        if name in idx and idx[name] < len(row):
            value = row[idx[name]]
            if value is not None and value != "":
                return value
    return None


def source_reference_text(row, idx):
    """收集采购订单里可能存在的源申请/源单字段，用于精确关联。"""
    values = []
    for name in SOURCE_DOC_HEADERS:
        if name in idx and idx[name] < len(row):
            v = text(row[idx[name]])
            if v:
                values.append(v)
    # 有些EAS导出会把源单号放到备注中，备注只作为补充，不单独证明关联。
    if "备注" in idx and idx["备注"] < len(row):
        remark = text(row[idx["备注"]])
        if remark:
            values.append(remark)
    return " | ".join(values)


def po_has_explicit_source_columns(idx):
    return any(name in idx for name in SOURCE_DOC_HEADERS)


def collect_eas(files):
    """
    汇总EAS资料，并处理【采购申请 -> 采购订单 -> EAS在途】关联。

    核心去重：
    - EAS在途 = 有效且未关闭采购订单的剩余入库数量。
    - EAS申请 = 尚未转成采购订单的有效申请剩余量 + 有效委外申请。
    - 已经转成采购订单的采购申请数量，从EAS申请中扣掉，不再重复计算。

    关联优先级：
    1) 若采购订单导出含“源单编号/采购申请单号”等字段，按申请单号精确匹配；
    2) 若源单字段不存在/为空，可按同物料、日期先后、FIFO进行匹配。
    """
    inventory = defaultdict(float)
    receipts = defaultdict(list)
    po_docs = defaultdict(list)

    purchase_orders = []
    purchase_apps = []
    outsource_apps = []

    for path in files:
        rows = read_xlsx_rows(path)
        if not rows:
            continue

        idx = get_header_index(rows[0])
        heads = set(idx.keys())

        # 1. 即时库存
        if "库存量(主单位)" in heads and "仓库名称" in heads:
            for row in rows[1:]:
                code = text(row[idx["物料编码"]]) if "物料编码" in idx and idx["物料编码"] < len(row) else ""
                if not code or code == "合计":
                    continue
                inventory[code] += number(row[idx["库存量(主单位)"]])

        # 2. 采购入库单：最新入库价 + 入库信息汇总
        elif "实收数量" in heads and "入库日期" in heads and "采购数量" in heads:
            for row in rows[1:]:
                code = text(row[idx["物料编码"]]) if "物料编码" in idx and idx["物料编码"] < len(row) else ""
                if not code or code == "合计":
                    continue

                qty = number(row[idx["实收数量"]])
                if qty <= 0:
                    continue

                price = None
                for key in ("成本价", "单价", "含税单价"):
                    if key in idx and idx[key] < len(row):
                        value = row[idx[key]]
                        if value is not None and value != "":
                            price = number(value)
                            break

                receipts[code].append({
                    "doc": text(row[idx["单据编号"]]) if "单据编号" in idx and idx["单据编号"] < len(row) else "",
                    "date": date_value(row[idx["入库日期"]]),
                    "qty": qty,
                    "price": price,
                })

        # 3. 采购订单
        elif "剩余入库数量" in heads and "采购数量" in heads and "交货日期" in heads:
            explicit_source_columns = po_has_explicit_source_columns(idx)

            for row in rows[1:]:
                code = text(row[idx["物料编码"]]) if "物料编码" in idx and idx["物料编码"] < len(row) else ""
                if not code or code == "合计":
                    continue

                status = text(row[idx["单据状态"]]) if "单据状态" in idx and idx["单据状态"] < len(row) else ""
                closed = text(row[idx["关闭状态"]]) if "关闭状态" in idx and idx["关闭状态"] < len(row) else ""
                biz = text(row[idx["业务关闭"]]) if "业务关闭" in idx and idx["业务关闭"] < len(row) else ""

                # 作废/取消/驳回/终止的订单，既不作为转单依据，也不算在途。
                if is_invalid_document_status(status) or is_invalid_document_status(closed) or is_invalid_document_status(biz):
                    continue

                ordered_qty = max(number(row[idx["采购数量"]]), 0.0)
                remain_cell = row[idx["剩余入库数量"]] if idx["剩余入库数量"] < len(row) else None

                if remain_cell is None or remain_cell == "":
                    received = number(row[idx["累计入库数量"]]) if "累计入库数量" in idx and idx["累计入库数量"] < len(row) else 0.0
                    returned = number(row[idx["累计退料数量"]]) if "累计退料数量" in idx and idx["累计退料数量"] < len(row) else 0.0
                    remain_qty = max(ordered_qty - received + returned, 0.0)
                else:
                    remain_qty = max(number(remain_cell), 0.0)

                doc = text(row[idx["单据编号"]]) if "单据编号" in idx and idx["单据编号"] < len(row) else ""
                po_date = date_value(row[idx["采购日期"]]) if "采购日期" in idx and idx["采购日期"] < len(row) else None
                name = text(row[idx["物料名称"]]) if "物料名称" in idx and idx["物料名称"] < len(row) else ""

                # 是否还属于EAS在途：关闭/结算订单不算在途，但仍可证明申请已转订单。
                is_open_for_transit = not (
                    is_closed_status(status) or
                    is_closed_status(closed) or
                    is_closed_status(biz)
                )

                purchase_orders.append({
                    "doc": doc,
                    "date": po_date,
                    "code": code,
                    "name": name,
                    "ordered_qty": ordered_qty,
                    "remain_qty": remain_qty,
                    "open_for_transit": is_open_for_transit,
                    "source_text": source_reference_text(row, idx),
                    "has_source_columns": explicit_source_columns,
                })

        # 4. 采购申请单
        elif "批准数量" in heads and "申请数量" in heads and "申请日期" in heads:
            for row in rows[1:]:
                code = text(row[idx["物料编码"]]) if "物料编码" in idx and idx["物料编码"] < len(row) else ""
                if not code or code == "合计":
                    continue

                status = ""
                if "数据状态" in idx and idx["数据状态"] < len(row):
                    status = text(row[idx["数据状态"]])
                elif "单据状态" in idx and idx["单据状态"] < len(row):
                    status = text(row[idx["单据状态"]])

                closed = text(row[idx["关闭状态"]]) if "关闭状态" in idx and idx["关闭状态"] < len(row) else ""
                if is_closed_status(status) or is_closed_status(closed):
                    continue

                approved = number(row[idx["批准数量"]])
                requested = number(row[idx["申请数量"]])
                qty = approved if approved != 0 else requested
                qty = max(qty, 0.0)
                if qty <= 0:
                    continue

                purchase_apps.append({
                    "type": "采购申请",
                    "doc": text(row[idx["单据编号"]]) if "单据编号" in idx and idx["单据编号"] < len(row) else "",
                    "date": date_value(row[idx["申请日期"]]),
                    "code": code,
                    "name": text(row[idx["物料名称"]]) if "物料名称" in idx and idx["物料名称"] < len(row) else "",
                    "qty": qty,
                    "remaining_qty": qty,
                    "allocations": [],
                })

        # 5. 委外申请单：没有采购订单关系时直接作为EAS申请；已结算/关闭的不计。
        elif "领料状态" in heads and "业务状态" in heads and "数量" in heads:
            for row in rows[1:]:
                code = text(row[idx["物料编码"]]) if "物料编码" in idx and idx["物料编码"] < len(row) else ""
                if not code or code == "合计":
                    continue

                status = text(row[idx["单据状态"]]) if "单据状态" in idx and idx["单据状态"] < len(row) else ""
                biz = text(row[idx["业务状态"]]) if "业务状态" in idx and idx["业务状态"] < len(row) else ""
                if is_closed_status(status) or is_closed_status(biz):
                    continue

                qty = max(number(row[idx["数量"]]), 0.0)
                if qty <= 0:
                    continue

                outsource_apps.append({
                    "type": "委外申请",
                    "doc": text(row[idx["单据编号"]]) if "单据编号" in idx and idx["单据编号"] < len(row) else "",
                    "date": date_value(row[idx["单据日期"]]) if "单据日期" in idx and idx["单据日期"] < len(row) else None,
                    "code": code,
                    "name": text(row[idx["物料名称"]]) if "物料名称" in idx and idx["物料名称"] < len(row) else "",
                    "qty": qty,
                })

    # --------------------------------------------------------
    # A. 采购申请 与 采购订单 关联去重
    # --------------------------------------------------------
    apps_by_code = defaultdict(list)
    for app in purchase_apps:
        apps_by_code[app["code"]].append(app)

    for code in apps_by_code:
        apps_by_code[code].sort(key=lambda x: (x["date"] or datetime.min, x["doc"]))

    purchase_orders.sort(key=lambda x: (x["date"] or datetime.min, x["doc"], x["code"]))

    relation_rows = []

    for po in purchase_orders:
        code = po["code"]
        qty_to_allocate = max(number(po["ordered_qty"]), 0.0)
        candidates = []
        match_method = ""

        # 优先：源单号/申请单号精确匹配。
        source_text_lower = text(po.get("source_text", "")).lower()
        if source_text_lower:
            for app in apps_by_code.get(code, []):
                if app["remaining_qty"] <= 0:
                    continue
                doc = text(app.get("doc", ""))
                if doc and doc.lower() in source_text_lower:
                    candidates.append(app)
            if candidates:
                match_method = "源申请单号精确关联"

        # 其次：当前导出没有可用源单号时，按同物料+日期+FIFO匹配。
        # 如果源单字段明确有值但没有匹配到已加载申请，则不强行FIFO，避免误关联。
        explicit_relation_declared = bool(po.get("has_source_columns") and source_text_lower)
        if (
            not candidates
            and FALLBACK_MATCH_APPLICATION_TO_PO
            and not explicit_relation_declared
        ):
            for app in apps_by_code.get(code, []):
                if app["remaining_qty"] <= 0:
                    continue
                app_date = app.get("date")
                po_date = po.get("date")
                if app_date is not None and po_date is not None and app_date > po_date:
                    continue
                candidates.append(app)
            if candidates:
                match_method = "同物料+日期FIFO推定"

        matched_apps = []
        for app in candidates:
            if qty_to_allocate <= 0:
                break
            available = max(number(app["remaining_qty"]), 0.0)
            if available <= 0:
                continue
            alloc = min(available, qty_to_allocate)
            if alloc <= 0:
                continue

            app["remaining_qty"] -= alloc
            qty_to_allocate -= alloc
            app["allocations"].append({
                "po_doc": po["doc"],
                "po_date": po["date"],
                "qty": alloc,
                "method": match_method,
            })
            matched_apps.append((app["doc"], alloc))

        po["matched_apps"] = matched_apps
        po["match_method"] = match_method if matched_apps else "未匹配到采购申请"

    # --------------------------------------------------------
    # B. 汇总 EAS在途、EAS申请（未转订单）
    # --------------------------------------------------------
    transit = defaultdict(float)
    applications = defaultdict(float)
    raw_applications = defaultdict(float)
    converted_applications = defaultdict(float)

    # 采购订单在途
    for po in purchase_orders:
        if po["open_for_transit"] and po["remain_qty"] > 0:
            transit[po["code"]] += po["remain_qty"]
            if po["doc"] and po["doc"] not in po_docs[po["code"]]:
                po_docs[po["code"]].append(po["doc"])

    # 采购申请：只保留尚未转采购订单的部分
    for app in purchase_apps:
        raw_applications[app["code"]] += app["qty"]
        converted_qty = max(app["qty"] - app["remaining_qty"], 0.0)
        converted_applications[app["code"]] += converted_qty
        applications[app["code"]] += max(app["remaining_qty"], 0.0)

        po_text = "；".join(
            "%s(%s)" % (x["po_doc"], clean_num(x["qty"]))
            for x in app["allocations"]
            if x.get("po_doc")
        )
        methods = []
        for x in app["allocations"]:
            m = x.get("method", "")
            if m and m not in methods:
                methods.append(m)

        relation_rows.append({
            "record_type": "采购申请",
            "code": app["code"],
            "name": app["name"],
            "app_doc": app["doc"],
            "app_date": app["date"],
            "app_qty": app["qty"],
            "converted_qty": converted_qty,
            "remaining_app_qty": max(app["remaining_qty"], 0.0),
            "po_doc": po_text,
            "po_date": None,
            "po_qty": None,
            "transit_qty": None,
            "relation": "；".join(methods) if methods else "尚未关联采购订单",
        })

    # 委外申请：保持为EAS申请，但不与标准采购订单做FIFO抵扣
    for app in outsource_apps:
        raw_applications[app["code"]] += app["qty"]
        applications[app["code"]] += app["qty"]
        relation_rows.append({
            "record_type": "委外申请",
            "code": app["code"],
            "name": app["name"],
            "app_doc": app["doc"],
            "app_date": app["date"],
            "app_qty": app["qty"],
            "converted_qty": 0.0,
            "remaining_app_qty": app["qty"],
            "po_doc": "",
            "po_date": None,
            "po_qty": None,
            "transit_qty": None,
            "relation": "委外申请未参与标准采购订单抵扣",
        })

    # 采购订单审计行
    for po in purchase_orders:
        matched_text = "；".join(
            "%s(%s)" % (doc, clean_num(qty))
            for doc, qty in po.get("matched_apps", [])
            if doc
        )
        relation_rows.append({
            "record_type": "采购订单",
            "code": po["code"],
            "name": po["name"],
            "app_doc": matched_text,
            "app_date": None,
            "app_qty": None,
            "converted_qty": None,
            "remaining_app_qty": None,
            "po_doc": po["doc"],
            "po_date": po["date"],
            "po_qty": po["ordered_qty"],
            "transit_qty": po["remain_qty"] if po["open_for_transit"] else 0.0,
            "relation": po.get("match_method", ""),
        })

    # --------------------------------------------------------
    # C. 最新入库价、入库信息汇总
    # --------------------------------------------------------
    latest_price = {}
    inbound_info = {}

    for code, items in receipts.items():
        items = sorted(
            items,
            key=lambda item: (item["date"] or datetime.min, item["doc"]),
            reverse=True,
        )

        latest_price[code] = next(
            (item["price"] for item in items if item["price"] is not None),
            None,
        )

        pieces = []
        for item in items:
            date_text = item["date"].strftime("%Y-%m-%d") if item["date"] else ""
            qty = clean_num(item["qty"])
            pieces.append("%s %s 入库%s" % (date_text, item["doc"], qty))

        inbound_info[code] = "；".join(pieces)

    return (
        inventory,
        latest_price,
        inbound_info,
        transit,
        applications,
        po_docs,
        relation_rows,
        raw_applications,
        converted_applications,
    )


# ============================================================
# 九、Excel输出辅助
# ============================================================
def copy_row_style(ws, source_row, target_row, start_col=1, end_col=22):
    """复制模板数据行样式。"""
    if target_row == source_row:
        return

    for col in range(start_col, end_col + 1):
        src = ws.cell(source_row, col)
        dst = ws.cell(target_row, col)

        if src.has_style:
            dst._style = copy(src._style)
        if src.number_format:
            dst.number_format = src.number_format
        if src.font:
            dst.font = copy(src.font)
        if src.fill:
            dst.fill = copy(src.fill)
        if src.border:
            dst.border = copy(src.border)
        if src.alignment:
            dst.alignment = copy(src.alignment)
        if src.protection:
            dst.protection = copy(src.protection)

    if source_row in ws.row_dimensions:
        ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height


def clear_old_data(ws, start_row, end_row):
    """清空A:V旧结果，不破坏模板样式，并取消旧隐藏状态。"""
    for row in ws.iter_rows(min_row=start_row, max_row=end_row, min_col=1, max_col=22):
        for cell in row:
            cell.value = None

    for r in range(start_row, end_row + 1):
        ws.row_dimensions[r].hidden = False

    try:
        ws.auto_filter.filterColumn = []
        ws.auto_filter.sortState = None
    except Exception:
        pass


def reset_sheet_view(ws, end_row):
    """重置Excel/WPS视图，避免旧冻结窗格导致数据行看不见。"""
    for r in range(1, max(end_row, START_ROW) + 1):
        ws.row_dimensions[r].hidden = False

    for col_letter in [chr(64 + i) for i in range(1, 23)]:
        ws.column_dimensions[col_letter].hidden = False

    ws.freeze_panes = None

    try:
        ws.sheet_view.topLeftCell = "A1"
        ws.sheet_view.zoomScale = 100
        ws.sheet_view.zoomScaleNormal = 100
        ws.sheet_view.selection = [Selection(activeCell="A6", sqref="A6")]
    except Exception:
        pass

    ws.freeze_panes = "A6"

    try:
        ws.sheet_view.selection = [
            Selection(pane="bottomLeft", activeCell="A6", sqref="A6")
        ]
    except Exception:
        pass


def set_excel_recalculation(workbook):
    """设置Excel打开时自动重算。虽然本版主要写入数值，也保留该设置。"""
    try:
        calc = workbook.calculation
        calc.calcMode = "auto"
        calc.fullCalcOnLoad = True
        calc.forceFullCalc = True
    except Exception:
        try:
            calc = workbook.calculation_properties
            calc.calcMode = "auto"
            calc.fullCalcOnLoad = True
            calc.forceFullCalc = True
        except Exception:
            pass


def style_simple_sheet(ws, widths):
    """给新增的订单/明细表做简单黑白灰样式。"""
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for cell in ws[1]:
        cell.font = Font(bold=True, color="000000")
        cell.fill = PatternFill(fill_type="solid", fgColor="D9D9D9")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)

    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


# ============================================================
# 十、写“订单输入”与“BOM需求明细”工作表
# ============================================================
def write_order_sheet(wb, parents, order_qty_map):
    if ORDER_SHEET_NAME in wb.sheetnames:
        del wb[ORDER_SHEET_NAME]
    ws = wb.create_sheet(ORDER_SHEET_NAME)

    ws.append(["父项物料编码", "父项物料名称", "订单数量"])
    for code, info in parents.items():
        ws.append([
            code,
            info.get("name", ""),
            clean_num(order_qty_map.get(code, 0)),
        ])

    style_simple_sheet(ws, {"A": 22, "B": 30, "C": 14})
    for r in range(2, ws.max_row + 1):
        ws.cell(r, 3).number_format = "0.###"


def write_detail_sheet(wb, active_details):
    if DETAIL_SHEET_NAME in wb.sheetnames:
        del wb[DETAIL_SHEET_NAME]
    ws = wb.create_sheet(DETAIL_SHEET_NAME)

    ws.append([
        "父项物料编码",
        "父项物料名称",
        "订单数量",
        "子项物料编码",
        "子项物料名称",
        "规格型号",
        "单机用量",
        "计量单位",
        "需求数量",
        "计算说明",
    ])

    for item in active_details:
        ws.append([
            item["parent"],
            item.get("parent_name", ""),
            clean_num(item["order_qty"]),
            item["code"],
            item["name"],
            item["spec"],
            clean_num(item["unit_qty"]),
            item["unit"],
            clean_num(item["required_qty"]),
            "%s × %s = %s" % (
                clean_num(item["unit_qty"]),
                clean_num(item["order_qty"]),
                clean_num(item["required_qty"]),
            ),
        ])

    style_simple_sheet(ws, {
        "A": 22,
        "B": 28,
        "C": 12,
        "D": 20,
        "E": 22,
        "F": 48,
        "G": 12,
        "H": 10,
        "I": 14,
        "J": 24,
    })

    for r in range(2, ws.max_row + 1):
        ws.cell(r, 3).number_format = "0.###"
        ws.cell(r, 7).number_format = "0.###"
        ws.cell(r, 9).number_format = "0.###"


def write_eas_relation_sheet(wb, relation_rows):
    """输出申请-采购订单关联明细，便于核对是否发生重复计算。"""
    if EAS_RELATION_SHEET_NAME in wb.sheetnames:
        del wb[EAS_RELATION_SHEET_NAME]
    ws = wb.create_sheet(EAS_RELATION_SHEET_NAME)

    ws.append([
        "记录类型",
        "物料编码",
        "物料名称",
        "申请单/匹配申请",
        "申请日期",
        "原申请数量",
        "已转订单数量",
        "未转申请数量(EAS申请)",
        "采购订单/关联订单",
        "采购日期",
        "采购数量",
        "剩余入库数量(EAS在途)",
        "关联方式/说明",
    ])

    for item in relation_rows:
        app_date = item.get("app_date")
        po_date = item.get("po_date")
        ws.append([
            item.get("record_type", ""),
            item.get("code", ""),
            item.get("name", ""),
            item.get("app_doc", ""),
            app_date.strftime("%Y-%m-%d") if app_date else "",
            clean_num(item.get("app_qty")),
            clean_num(item.get("converted_qty")),
            clean_num(item.get("remaining_app_qty")),
            item.get("po_doc", ""),
            po_date.strftime("%Y-%m-%d") if po_date else "",
            clean_num(item.get("po_qty")),
            clean_num(item.get("transit_qty")),
            item.get("relation", ""),
        ])

    style_simple_sheet(ws, {
        "A": 12, "B": 20, "C": 24, "D": 28, "E": 12,
        "F": 14, "G": 14, "H": 20, "I": 28, "J": 12,
        "K": 14, "L": 20, "M": 28,
    })

    for r in range(2, ws.max_row + 1):
        for c in (6, 7, 8, 11, 12):
            ws.cell(r, c).number_format = "0.###"


# ============================================================
# 十一、写总表
# ============================================================
def write_output(template, output, parents, order_qty_map, materials, active_details, eas):
    (
        inventory, latest_price, inbound_info, transit, applications, po_docs,
        relation_rows, raw_applications, converted_applications,
    ) = eas

    result_rows = []

    for code, item in materials.items():
        required_qty = number(item["required_qty"])
        if required_qty <= 0:
            continue

        stock = number(inventory.get(code, 0.0))
        price = latest_price.get(code)
        in_transit = number(transit.get(code, 0.0))
        app_qty = number(applications.get(code, 0.0))

        shortage = max(required_qty - stock, 0.0)
        purchase_gap = max(shortage - in_transit, 0.0)
        application_gap = max(purchase_gap - app_qty, 0.0)
        stock_balance = stock - required_qty

        quota_display, quota_numeric = get_quota_display(item["contributions"])

        stock_balance_sets = None
        transit_sets = None
        total_balance_sets = None
        if quota_numeric not in (None, 0):
            stock_balance_sets = stock_balance / quota_numeric
            transit_sets = in_transit / quota_numeric
            total_balance_sets = (stock_balance + in_transit) / quota_numeric

        stock_balance_amount = None if price is None else stock_balance * price
        transit_amount = None if price is None else in_transit * price
        total_amount = None if price is None else (stock_balance + in_transit) * price

        result_rows.append({
            "code": code,
            "name": item["name"],
            "spec": item["spec"],
            "unit": item["unit"],
            "inventory": stock,
            "price": price,
            "required_qty": required_qty,
            "shortage": shortage,
            "inbound_info": inbound_info.get(code, ""),
            "transit": in_transit,
            "purchase_gap": purchase_gap,
            "applications": app_qty,
            "application_gap": application_gap,
            "orders": "；".join(po_docs.get(code, [])),
            "stock_balance": stock_balance,
            "stock_balance_sets": stock_balance_sets,
            "stock_balance_amount": stock_balance_amount,
            "transit_sets": transit_sets,
            "transit_amount": transit_amount,
            "total_balance_sets": total_balance_sets,
            "total_amount": total_amount,
            "quota_display": quota_display,
            "contributions": item["contributions"],
        })

    wb = load_workbook(str(template), data_only=False)

    if OUTPUT_SHEET_NAME not in wb.sheetnames:
        wb.close()
        raise KeyError("模板中找不到工作表：%s" % OUTPUT_SHEET_NAME)

    ws = wb[OUTPUT_SHEET_NAME]

    # 清理旧数据
    clear_old_data(ws, START_ROW, MAX_CLEAR_ROW)

    # 顶部说明：不再使用一个全局“齐套数量目标”
    ws["A3"] = "订单数量"
    order_summary = []
    for code, info in parents.items():
        qty = number(order_qty_map.get(code, 0))
        if qty > 0:
            order_summary.append("%s=%s" % (code, clean_num(qty)))
    ws["B3"] = "；".join(order_summary)

    ws["B4"] = (
        "核料逻辑：BOM额定消耗为各父项单机用量；先按“单机用量×该父项订单数量”计算需求，"
        "再按子项物料编码汇总。欠料=MAX(用量合计-库存数量,0)；"
        "采购缺口=MAX(欠料-EAS在途,0)；EAS申请只保留尚未转采购订单的申请数量，"
        "已转订单申请不再与EAS在途重复；申请缺口=MAX(采购缺口-EAS申请,0)。"
    )

    # 更正V列表头含义
    ws.cell(5, 22, "单机用量")

    if result_rows:
        end_row = START_ROW + len(result_rows) - 1

        # 复制第6行格式到后续结果行
        for excel_row in range(START_ROW + 1, end_row + 1):
            copy_row_style(ws, START_ROW, excel_row)

        for offset, item in enumerate(result_rows):
            r = START_ROW + offset

            # A:F
            ws.cell(r, 1, item["code"])                    # 欠料数据 / 物料编码
            ws.cell(r, 2, "")                              # 替代料编码
            ws.cell(r, 3, item["name"])                   # 物料名称
            ws.cell(r, 4, item["spec"])                   # 规格型号
            ws.cell(r, 5, clean_num(item["inventory"]))   # 库存数量
            ws.cell(r, 6, item["price"])                  # 入库价

            # G:H —— 本版关键：G是已按各父项订单数量计算后的总需求
            ws.cell(r, 7, clean_num(item["required_qty"]))
            ws.cell(r, 8, clean_num(item["shortage"]))

            # I:N
            ws.cell(r, 9, item["inbound_info"])
            ws.cell(r, 10, clean_num(item["transit"]))
            ws.cell(r, 11, clean_num(item["purchase_gap"]))
            ws.cell(r, 12, clean_num(item["applications"]))
            ws.cell(r, 13, clean_num(item["application_gap"]))
            ws.cell(r, 14, item["orders"])

            # O:V
            ws.cell(r, 15, clean_num(item["stock_balance"]))
            ws.cell(r, 16, item["stock_balance_sets"])
            ws.cell(r, 17, item["stock_balance_amount"])
            ws.cell(r, 18, item["transit_sets"])
            ws.cell(r, 19, item["transit_amount"])
            ws.cell(r, 20, item["total_balance_sets"])
            ws.cell(r, 21, item["total_amount"])
            ws.cell(r, 22, item["quota_display"])

            # V列如果是多个父项的不同单机用量说明，则自动换行
            ws.cell(r, 22).alignment = copy(ws.cell(r, 22).alignment)
            ws.cell(r, 22).alignment = Alignment(
                horizontal=ws.cell(r, 22).alignment.horizontal,
                vertical=ws.cell(r, 22).alignment.vertical,
                text_rotation=ws.cell(r, 22).alignment.text_rotation,
                wrap_text=True,
                shrink_to_fit=ws.cell(r, 22).alignment.shrink_to_fit,
                indent=ws.cell(r, 22).alignment.indent,
            )

        # 数字格式
        for r in range(START_ROW, end_row + 1):
            ws.cell(r, 5).number_format = "0.###"
            ws.cell(r, 6).number_format = "0.000000"
            ws.cell(r, 7).number_format = "0.###"
            ws.cell(r, 8).number_format = "0.###"
            ws.cell(r, 10).number_format = "0.###"
            ws.cell(r, 11).number_format = "0.###"
            ws.cell(r, 12).number_format = "0.###"
            ws.cell(r, 13).number_format = "0.###"
            ws.cell(r, 15).number_format = "0.###"
            ws.cell(r, 16).number_format = "0.00"
            ws.cell(r, 17).number_format = "0.00"
            ws.cell(r, 18).number_format = "0.00"
            ws.cell(r, 19).number_format = "0.00"
            ws.cell(r, 20).number_format = "0.00"
            ws.cell(r, 21).number_format = "0.00"
            if isinstance(ws.cell(r, 22).value, (int, float)):
                ws.cell(r, 22).number_format = "0.###"

        # 避免入库价显示 ######
        if ws.column_dimensions["F"].width is None or ws.column_dimensions["F"].width < 12:
            ws.column_dimensions["F"].width = 12
        if ws.column_dimensions["V"].width is None or ws.column_dimensions["V"].width < 18:
            ws.column_dimensions["V"].width = 18

        # 数据行全部可见
        for r in range(START_ROW, end_row + 1):
            ws.row_dimensions[r].hidden = False

        # 重新设置筛选区域
        try:
            ws.auto_filter.ref = "A5:V%d" % end_row
            ws.auto_filter.filterColumn = []
            ws.auto_filter.sortState = None
        except Exception:
            pass

        # 申请缺口 > 0 时浅灰/浅红提醒（不影响主逻辑）
        red_fill = PatternFill(fill_type="solid", fgColor="FCE8E6")
        try:
            ws.conditional_formatting.add(
                "M%d:M%d" % (START_ROW, end_row),
                CellIsRule(operator="greaterThan", formula=["0"], fill=red_fill),
            )
        except Exception:
            pass
    else:
        end_row = START_ROW

    # 辅助表：订单输入、BOM需求明细、EAS申请-订单关联明细
    write_order_sheet(wb, parents, order_qty_map)
    write_detail_sheet(wb, active_details)
    write_eas_relation_sheet(wb, relation_rows)

    # 重置视图，避免旧模板打开只显示1行
    reset_sheet_view(ws, end_row)
    set_excel_recalculation(wb)

    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(output))
    wb.close()

    return result_rows



# ============================================================
# 十二、Windows桌面界面
# ============================================================
class DesktopApp(object):
    def __init__(self, root):
        import tkinter as tk
        from tkinter import ttk

        self.tk = tk
        self.ttk = ttk
        self.root = root
        self.parents = OrderedDict()
        self.bom_details = []
        self.loaded_bom_path = ""
        self.order_vars = OrderedDict()
        self.order_rows = OrderedDict()
        self.last_output = None
        self.busy = False

        self.root.title("PY订单核料工具 - BOM兼容修正版")
        self.root.geometry("1120x820")
        self.root.minsize(940, 680)

        self.bom_var = tk.StringVar(value=str(BOM_ZIP_PATH) if Path(BOM_ZIP_PATH).exists() else "")
        self.eas_var = tk.StringVar(value=str(SOURCE_ZIP_PATH) if Path(SOURCE_ZIP_PATH).exists() else "")
        self.template_var = tk.StringVar(value=str(TEMPLATE_XLSX_PATH) if Path(TEMPLATE_XLSX_PATH).exists() else "")
        self.output_var = tk.StringVar(value=str(self._default_output_path()))
        self.search_var = tk.StringVar()
        self.only_active_var = tk.BooleanVar(value=False)
        self.count_var = tk.StringVar(value="尚未读取BOM")
        self.status_var = tk.StringVar(value="请选择文件，然后点击“读取BOM”。")

        self._build_ui()

    def _default_output_path(self):
        base = BASE_DIR
        try:
            t = Path(str(TEMPLATE_XLSX_PATH))
            if t.exists():
                base = t.parent
        except Exception:
            pass
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return base / ("PY订单核料表_%s.xlsx" % stamp)

    def _build_ui(self):
        tk = self.tk
        ttk = self.ttk

        outer = ttk.Frame(self.root, padding=10)
        outer.pack(fill="both", expand=True)

        title = ttk.Label(outer, text="PY订单核料工具", font=("Microsoft YaHei", 15, "bold"))
        title.pack(anchor="w", pady=(0, 8))

        file_box = ttk.LabelFrame(outer, text="1. 数据文件", padding=8)
        file_box.pack(fill="x")
        file_box.columnconfigure(1, weight=1)

        self._path_row(file_box, 0, "BOM文件/压缩包", self.bom_var, self._choose_bom)
        self._path_row(file_box, 1, "EAS资料/压缩包", self.eas_var, self._choose_eas)
        self._path_row(file_box, 2, "核料模板", self.template_var, self._choose_template)
        self._path_row(file_box, 3, "输出结果", self.output_var, self._choose_output)

        action_line = ttk.Frame(file_box)
        action_line.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        self.load_btn = ttk.Button(action_line, text="读取BOM", width=14, command=self.load_bom)
        self.load_btn.pack(side="left")
        ttk.Label(
            action_line,
            text="支持 .zip、单个 .xlsx；BOM很多时建议放进一个压缩包。",
        ).pack(side="left", padx=10)

        order_box = ttk.LabelFrame(outer, text="2. 输入各线束订单数量", padding=8)
        order_box.pack(fill="both", expand=True, pady=(10, 0))

        tools = ttk.Frame(order_box)
        tools.pack(fill="x", pady=(0, 6))
        ttk.Label(tools, text="搜索BOM：").pack(side="left")
        search_entry = ttk.Entry(tools, textvariable=self.search_var)
        search_entry.pack(side="left", fill="x", expand=True, padx=(4, 8))
        ttk.Checkbutton(
            tools,
            text="只显示订单数量>0",
            variable=self.only_active_var,
            command=self._apply_filter,
        ).pack(side="left", padx=(0, 10))
        ttk.Button(tools, text="全部清零", command=self._clear_order_qty).pack(side="left")
        ttk.Label(tools, textvariable=self.count_var).pack(side="right", padx=(10, 0))

        header = ttk.Frame(order_box)
        header.pack(fill="x", padx=(2, 18))
        ttk.Label(header, text="父项物料编码", width=24, anchor="w", font=("Microsoft YaHei", 9, "bold")).pack(side="left")
        ttk.Label(header, text="物料名称", width=56, anchor="w", font=("Microsoft YaHei", 9, "bold")).pack(side="left")
        ttk.Label(header, text="订单数量", width=16, anchor="e", font=("Microsoft YaHei", 9, "bold")).pack(side="left")

        body = ttk.Frame(order_box)
        body.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(body, highlightthickness=0, borderwidth=0)
        self.scrollbar = ttk.Scrollbar(body, orient="vertical", command=self.canvas.yview)
        self.order_inner = ttk.Frame(self.canvas)
        self.order_window = self.canvas.create_window((0, 0), window=self.order_inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.order_inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(self.order_window, width=e.width))
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        try:
            self.search_var.trace_add("write", lambda *args: self._apply_filter())
        except AttributeError:
            self.search_var.trace("w", lambda *args: self._apply_filter())

        run_bar = ttk.Frame(outer)
        run_bar.pack(fill="x", pady=(10, 0))
        self.run_btn = ttk.Button(run_bar, text="开始核料", width=18, command=self.start_calculation)
        self.run_btn.pack(side="left")
        self.open_btn = ttk.Button(run_bar, text="打开结果文件", width=16, command=self.open_result, state="disabled")
        self.open_btn.pack(side="left", padx=(8, 0))
        ttk.Label(run_bar, textvariable=self.status_var).pack(side="left", padx=12)

        log_box = ttk.LabelFrame(outer, text="3. 处理日志", padding=6)
        log_box.pack(fill="x", pady=(10, 0))
        log_frame = ttk.Frame(log_box)
        log_frame.pack(fill="x")
        self.log_text = tk.Text(log_frame, height=8, wrap="word")
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.pack(side="left", fill="x", expand=True)
        log_scroll.pack(side="right", fill="y")
        self.log_text.configure(state="disabled")

        foot = ttk.Label(
            outer,
            text="核料规则：先按每个BOM的“单机用量×订单数量”计算需求，再按子项编码汇总；EAS申请会剔除已转采购订单的数量，避免与在途重复。",
            wraplength=1050,
        )
        foot.pack(fill="x", pady=(8, 0))

    def _path_row(self, parent, row, label, variable, command):
        ttk = self.ttk
        ttk.Label(parent, text=label, width=16, anchor="w").grid(row=row, column=0, sticky="w", pady=3)
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", padx=(4, 8), pady=3)
        ttk.Button(parent, text="选择", width=10, command=command).grid(row=row, column=2, sticky="e", pady=3)

    def _choose_bom(self):
        from tkinter import filedialog
        path = filedialog.askopenfilename(
            title="选择BOM压缩包或Excel",
            filetypes=[("BOM文件", "*.zip *.xlsx"), ("ZIP压缩包", "*.zip"), ("Excel", "*.xlsx"), ("所有文件", "*.*")],
        )
        if path:
            self.bom_var.set(path)
            self.loaded_bom_path = ""
            self.status_var.set("BOM文件已更换，请重新点击“读取BOM”。")

    def _choose_eas(self):
        from tkinter import filedialog
        path = filedialog.askopenfilename(
            title="选择EAS资料压缩包或Excel",
            filetypes=[("EAS资料", "*.zip *.xlsx"), ("ZIP压缩包", "*.zip"), ("Excel", "*.xlsx"), ("所有文件", "*.*")],
        )
        if path:
            self.eas_var.set(path)

    def _choose_template(self):
        from tkinter import filedialog
        path = filedialog.askopenfilename(
            title="选择PY订单核料表模板",
            filetypes=[("Excel工作簿", "*.xlsx"), ("所有文件", "*.*")],
        )
        if path:
            self.template_var.set(path)
            # 模板切换时自动给一个新的默认输出名，避免覆盖模板或已打开的结果。
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            self.output_var.set(str(Path(path).parent / ("PY订单核料表_%s.xlsx" % stamp)))

    def _choose_output(self):
        from tkinter import filedialog
        current = self.output_var.get().strip()
        initialdir = str(Path(current).parent) if current else str(BASE_DIR)
        initialfile = Path(current).name if current else "PY订单核料表.xlsx"
        path = filedialog.asksaveasfilename(
            title="保存核料结果",
            defaultextension=".xlsx",
            initialdir=initialdir,
            initialfile=initialfile,
            filetypes=[("Excel工作簿", "*.xlsx")],
        )
        if path:
            self.output_var.set(path)

    def _on_mousewheel(self, event):
        try:
            if event.delta:
                self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except Exception:
            pass

    def _log(self, msg):
        def write():
            self.log_text.configure(state="normal")
            self.log_text.insert("end", str(msg) + "\n")
            self.log_text.see("end")
            self.log_text.configure(state="disabled")
        if threading.current_thread() is threading.main_thread():
            write()
        else:
            self.root.after(0, write)

    def _set_busy(self, value, status=None):
        self.busy = bool(value)
        state = "disabled" if value else "normal"
        self.load_btn.configure(state=state)
        self.run_btn.configure(state=state)
        if status is not None:
            self.status_var.set(status)

    def _clear_order_rows(self):
        for child in self.order_inner.winfo_children():
            child.destroy()
        self.order_vars.clear()
        self.order_rows.clear()
        self.count_var.set("尚未读取BOM")

    def _render_order_rows(self):
        ttk = self.ttk
        tk = self.tk
        self._clear_order_rows()

        defaults = default_order_qty(self.parents)
        for idx, (code, info) in enumerate(self.parents.items()):
            frame = ttk.Frame(self.order_inner)
            frame.grid(row=idx, column=0, sticky="ew", pady=2)
            frame.columnconfigure(1, weight=1)

            ttk.Label(frame, text=code, width=24, anchor="w").grid(row=0, column=0, sticky="w")
            ttk.Label(frame, text=info.get("name", ""), width=56, anchor="w").grid(row=0, column=1, sticky="w")

            var = tk.StringVar(value=str(clean_num(defaults.get(code, 0))))
            ent = ttk.Entry(frame, textvariable=var, width=16, justify="right")
            ent.grid(row=0, column=2, sticky="e", padx=(6, 10))

            try:
                var.trace_add("write", lambda *args: self._apply_filter() if self.only_active_var.get() else None)
            except AttributeError:
                pass

            self.order_vars[code] = var
            self.order_rows[code] = (frame, text(info.get("name", "")))

        self._apply_filter()
        self.root.after_idle(lambda: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

    def _apply_filter(self):
        keyword = self.search_var.get().strip().lower()
        only_active = bool(self.only_active_var.get())
        shown = 0
        active = 0
        for code, pair in self.order_rows.items():
            frame, name = pair
            raw = self.order_vars[code].get().strip()
            qty = number(raw)
            if qty > 0:
                active += 1
            matched = (not keyword) or (keyword in code.lower()) or (keyword in name.lower())
            if only_active and qty <= 0:
                matched = False
            if matched:
                frame.grid()
                shown += 1
            else:
                frame.grid_remove()
        if self.parents:
            self.count_var.set("显示 %d / 共 %d 个BOM；有订单 %d 个" % (shown, len(self.parents), active))
        self.root.after_idle(lambda: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

    def _clear_order_qty(self):
        for var in self.order_vars.values():
            var.set("0")
        self._apply_filter()

    def load_bom(self):
        if self.busy:
            return
        path_text = self.bom_var.get().strip()
        if not path_text:
            self._show_error("请先选择BOM文件或压缩包。")
            return
        path = Path(path_text)
        if not path.exists():
            self._show_error("BOM文件不存在：\n%s" % path)
            return

        self._set_busy(True, "正在读取BOM...")
        self._log("开始读取BOM：%s" % path)

        def worker():
            temp_dirs = []
            try:
                bom_files = xlsx_files(path, temp_dirs)
                if not bom_files:
                    raise RuntimeError("没有找到xlsx格式的BOM文件。")
                diagnostics = []
                parents, details = collect_bom_details(bom_files, diagnostics=diagnostics)
                if not parents or not details:
                    raise RuntimeError("没有识别到有效父项和BOM子项，请检查BOM表头。")
                self.root.after(0, lambda: self._finish_load_bom(path, parents, details, len(bom_files), diagnostics))
            except Exception as exc:
                detail = traceback.format_exc()
                self._log(detail)
                self.root.after(0, lambda e=str(exc): self._load_failed(e))
            finally:
                for td in temp_dirs:
                    try:
                        td.cleanup()
                    except Exception:
                        pass

        threading.Thread(target=worker, daemon=True).start()

    def _finish_load_bom(self, path, parents, details, file_count, diagnostics=None):
        self.parents = parents
        self.bom_details = details
        self.loaded_bom_path = str(path.resolve())
        self._render_order_rows()
        self._set_busy(False, "BOM读取完成：%d 个父项。" % len(parents))
        self._log("读取到BOM Excel：%d 个；识别父项：%d 个；父项-子项明细：%d 条。" % (file_count, len(parents), len(details)))
        for msg in diagnostics or []:
            self._log(msg)

    def _load_failed(self, message):
        self._set_busy(False, "BOM读取失败。")
        self._show_error(message)

    def _get_order_qty_map(self):
        values = OrderedDict()
        for code in self.parents.keys():
            raw = self.order_vars[code].get().strip()
            if raw == "":
                qty = 0.0
            else:
                try:
                    qty = float(raw.replace(",", ""))
                except ValueError:
                    raise ValueError("%s 的订单数量不是有效数字：%s" % (code, raw))
            if qty < 0:
                raise ValueError("%s 的订单数量不能小于0。" % code)
            values[code] = qty
        return values

    def start_calculation(self):
        if self.busy:
            return
        if not self.parents or not self.bom_details:
            self._show_error("请先点击“读取BOM”，确认订单列表已经显示。")
            return
        try:
            current_bom = str(Path(self.bom_var.get().strip()).resolve())
        except Exception:
            current_bom = self.bom_var.get().strip()
        if current_bom != self.loaded_bom_path:
            self._show_error("BOM路径已经变化，请重新点击“读取BOM”。")
            return

        try:
            qty_map = self._get_order_qty_map()
        except Exception as exc:
            self._show_error(str(exc))
            return
        active_orders = [(k, v) for k, v in qty_map.items() if number(v) > 0]
        if not active_orders:
            self._show_error("至少需要填写一个大于0的订单数量。")
            return

        eas_path = Path(self.eas_var.get().strip())
        template_path = Path(self.template_var.get().strip())
        output_text = self.output_var.get().strip()
        if not eas_path.exists():
            self._show_error("EAS资料不存在：\n%s" % eas_path)
            return
        if not template_path.exists():
            self._show_error("核料模板不存在：\n%s" % template_path)
            return
        if not output_text:
            self._show_error("请设置输出结果文件。")
            return
        output_path = Path(output_text)
        if output_path.suffix.lower() != ".xlsx":
            output_path = output_path.with_suffix(".xlsx")
            self.output_var.set(str(output_path))
        if output_path.resolve() == template_path.resolve():
            self._show_error("输出文件不能和模板文件是同一个文件。")
            return

        self._set_busy(True, "正在核料，请稍候...")
        self.open_btn.configure(state="disabled")
        self._log("-" * 60)
        self._log("本次有订单的BOM：%d 个。" % len(active_orders))
        for code, qty in active_orders:
            self._log("  %s  订单数量=%s" % (code, clean_num(qty)))

        def worker():
            temp_dirs = []
            try:
                self._log("计算BOM物料需求...")
                materials, active_details = build_material_requirements(self.bom_details, qty_map)
                if not materials:
                    raise RuntimeError("没有计算出物料需求，请检查订单数量和BOM数据。")
                self._log("需求物料汇总：%d 种。" % len(materials))

                self._log("读取EAS资料并进行申请/在途关联去重...")
                source_files = xlsx_files(eas_path, temp_dirs)
                if not source_files:
                    raise RuntimeError("EAS资料中没有找到xlsx文件。")
                eas = collect_eas(source_files)

                self._log("生成Excel核料结果...")
                result_rows = write_output(
                    template_path,
                    output_path,
                    self.parents,
                    qty_map,
                    materials,
                    active_details,
                    eas,
                )

                relation_rows = eas[6] if len(eas) > 6 else []
                self.root.after(
                    0,
                    lambda: self._finish_calculation(output_path, len(result_rows), len(active_details), len(relation_rows)),
                )
            except PermissionError:
                msg = "输出文件可能正在Excel/WPS中打开，无法覆盖。请先关闭结果文件，或换一个输出文件名。"
                self.root.after(0, lambda: self._calc_failed(msg))
            except Exception as exc:
                detail = traceback.format_exc()
                self._log(detail)
                self.root.after(0, lambda e=str(exc): self._calc_failed(e))
            finally:
                for td in temp_dirs:
                    try:
                        td.cleanup()
                    except Exception:
                        pass

        threading.Thread(target=worker, daemon=True).start()

    def _finish_calculation(self, output_path, material_count, detail_count, relation_count):
        self.last_output = Path(output_path)
        self._set_busy(False, "核料完成。")
        self.open_btn.configure(state="normal")
        self._log("核料完成：总表 %d 种物料；BOM需求明细 %d 条；EAS关联明细 %d 条。" % (material_count, detail_count, relation_count))
        self._log("输出：%s" % output_path)
        from tkinter import messagebox
        messagebox.showinfo(
            "完成",
            "核料完成。\n\n结果文件：\n%s\n\n工作表：总表、订单输入、BOM需求明细、EAS关联明细。" % output_path,
        )

    def _calc_failed(self, message):
        self._set_busy(False, "核料失败。")
        self._show_error(message)

    def open_result(self):
        if not self.last_output or not self.last_output.exists():
            self._show_error("暂时没有可打开的结果文件。")
            return
        try:
            if os.name == "nt":
                os.startfile(str(self.last_output))
            else:
                import subprocess
                subprocess.Popen(["xdg-open", str(self.last_output)])
        except Exception as exc:
            self._show_error("无法自动打开结果文件：\n%s" % exc)

    def _show_error(self, message):
        from tkinter import messagebox
        messagebox.showerror("提示", str(message))


# ============================================================
# 十三、程序入口
# ============================================================
def launch_desktop_app():
    try:
        import tkinter as tk
    except ImportError:
        raise SystemExit("当前Python没有安装 tkinter，无法启动桌面界面。")

    root = tk.Tk()
    DesktopApp(root)
    root.mainloop()


if __name__ == "__main__":
    launch_desktop_app()
