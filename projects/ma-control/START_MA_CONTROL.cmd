@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

title MA CONTROL — Dieu Hanh FG (Webapp Launcher)
echo ===============================================================================
echo            MA CONTROL — DIEU HANH FG (PYTHON + EXCEL-LIKE UI)
echo ===============================================================================
echo.

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
set "PYTHONPATH=%SCRIPT_DIR%"

if exist "%SCRIPT_DIR%runtime\python.exe" (
    set "MA_CONTROL_RUNTIME=%SCRIPT_DIR%runtime\python.exe"
    echo [*] Su dung runtime di kem: !MA_CONTROL_RUNTIME!
) else (
    set "MA_CONTROL_RUNTIME=python"
    echo [*] Su dung Python he thong: python
)

:: Configure database & exports relative to this folder
if not exist "%SCRIPT_DIR%data" mkdir "%SCRIPT_DIR%data"
if not exist "%SCRIPT_DIR%data\bbbg_exports" mkdir "%SCRIPT_DIR%data\bbbg_exports"
set "MA_CONTROL_DB_PATH=%SCRIPT_DIR%data\ma_control_fg.db"
set "MA_CONTROL_EXPORT_DIR=%SCRIPT_DIR%data\bbbg_exports"
set "MA_CONTROL_SAP_MODE=MANUAL"

:: Check and find available port (default 8012)
set "PORT=8012"
netstat -ano | findstr /R /C:":8012 " >nul
if !errorlevel! equ 0 (
    echo [INFO] Cong 8012 dang ban, tu dong chuyen sang cong 8013...
    set "PORT=8013"
    netstat -ano | findstr /R /C:":8013 " >nul
    if !errorlevel! equ 0 (
        echo [INFO] Cong 8013 dang ban, tu dong chuyen sang cong 8014...
        set "PORT=8014"
    )
)

set "MA_CONTROL_PORT=%PORT%"
set "HOST=127.0.0.1"

echo [*] Thu muc lam viec: %SCRIPT_DIR%
echo [*] Co so du lieu:    %MA_CONTROL_DB_PATH%
echo [*] Che do SAP:       %MA_CONTROL_SAP_MODE%
echo [*] Khoi dong tai:    http://%HOST%:%PORT%
echo.

:: Initialize database if needed
"%MA_CONTROL_RUNTIME%" -m app.core.database >nul 2>&1

:: Open browser after 2 seconds in background
start "" cmd /c "timeout /t 2 /nobreak >nul & start http://%HOST%:%PORT%"

:: Launch Uvicorn server
"%MA_CONTROL_RUNTIME%" -m uvicorn app.main:app --host %HOST% --port %PORT%
