@echo off
chcp 65001 >nul
setlocal

title RUN TESTS — MA CONTROL
set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
set "PYTHONPATH=%SCRIPT_DIR%"

if exist "%SCRIPT_DIR%runtime\python.exe" (
    set "MA_CONTROL_RUNTIME=%SCRIPT_DIR%runtime\python.exe"
) else (
    set "MA_CONTROL_RUNTIME=python"
)

echo [*] Dang chay toan bo bo test tu dong (FastAPI, SQLite, Logic OB/CF, BBBG, UI Mockup)...
"%MA_CONTROL_RUNTIME%" -m pytest tests -v
set TEST_EXIT_CODE=%errorlevel%

if %TEST_EXIT_CODE% equ 0 (
    echo.
    echo [PASS] Tat ca cac test case deu vuot qua thanh cong!
) else (
    echo.
    echo [FAIL] Co test case bi loi. Ma thoat: %TEST_EXIT_CODE%
)

exit /b %TEST_EXIT_CODE%
