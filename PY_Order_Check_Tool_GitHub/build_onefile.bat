@echo off
setlocal EnableExtensions
cd /d "%~dp0"

REM ============================================================
REM Generic Windows one-file build script for the public repo.
REM Runtime: Python 3.8+ recommended for this project.
REM Before building, place an anonymized/local template.xlsx here.
REM ============================================================

if not exist "template.xlsx" (
    echo ERROR: template.xlsx was not found.
    echo Put your LOCAL/anonymized template.xlsx in this folder first.
    pause
    exit /b 1
)

set "PY_CMD="

REM 1) Optional custom interpreter path supplied by user
if defined PYTHON38_EXE (
    if exist "%PYTHON38_EXE%" (
        set "PY_CMD=\"%PYTHON38_EXE%\""
        goto :python_found
    )
)

REM 2) Try Windows Python Launcher
where py >nul 2>&1
if not errorlevel 1 (
    py -3.8 --version >nul 2>&1
    if not errorlevel 1 (
        set "PY_CMD=py -3.8"
        goto :python_found
    )
)

REM 3) Fall back to python on PATH
where python >nul 2>&1
if not errorlevel 1 (
    set "PY_CMD=python"
    goto :python_found
)

echo ERROR: Python was not found.
echo Install Python or set PYTHON38_EXE to your python.exe path.
pause
exit /b 1

:python_found
echo Using Python: %PY_CMD%
%PY_CMD% --version
if errorlevel 1 goto :failed

if exist ".build_env" rmdir /s /q ".build_env"
%PY_CMD% -m venv ".build_env"
if errorlevel 1 goto :failed

set "BUILD_PY=.build_env\Scripts\python.exe"
"%BUILD_PY%" -m pip install --upgrade pip
if errorlevel 1 goto :failed
"%BUILD_PY%" -m pip install -r requirements-dev.txt
if errorlevel 1 goto :failed

if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "PY_Order_Check_Tool.spec" del /q "PY_Order_Check_Tool.spec"

"%BUILD_PY%" -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --windowed ^
  --onefile ^
  --name "PY_Order_Check_Tool" ^
  --add-data "template.xlsx;." ^
  "app.pyw"
if errorlevel 1 goto :failed

echo.
echo BUILD SUCCESS: dist\PY_Order_Check_Tool.exe
explorer "dist"
pause
exit /b 0

:failed
echo.
echo BUILD FAILED.
pause
exit /b 1
