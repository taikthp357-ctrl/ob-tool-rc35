@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
title MA CONTROL FG - Enterprise Launcher (Zero-Port)
set "PYTHONUTF8=1"
set "MA_CONTROL_ENV=COMPANY"

echo ======================================================================
echo          MA CONTROL — DIEU HANH FG (ENTERPRISE LAUNCHER)
echo        100%% Zero-Port Mode ^& Corporate Firewall Compliant
echo ======================================================================
echo.

:: 1. Xác định thư mục Share Release (chứa launcher.cmd này)
set "SHARE_ROOT=%~dp0"
if "%SHARE_ROOT:~-1%"=="\" set "SHARE_ROOT=%SHARE_ROOT:~0,-1%"
cd /d "%SHARE_ROOT%"

:: Cấu hình PYTHONPATH để Python tìm thấy package app ngay cả khi chạy từ release có folder payload
set "PYTHONPATH=%SHARE_ROOT%\payload;%SHARE_ROOT%;%PYTHONPATH%"

:: 2. Tìm Python Runtime (ưu tiên embedded runtime nội bộ nếu có)
set "PYTHON_BIN="
if exist "%SHARE_ROOT%\runtime\python.exe" (
    set "PYTHON_BIN=%SHARE_ROOT%\runtime\python.exe"
) else (
    for /f "tokens=*" %%i in ('where python 2^>nul') do (
        set "CANDIDATE=%%i"
        echo !CANDIDATE! | findstr /i /c:"WindowsApps" >nul
        if errorlevel 1 (
            if not defined PYTHON_BIN set "PYTHON_BIN=!CANDIDATE!"
        )
    )
)

if not defined PYTHON_BIN (
    echo [LOI] Khong tim thay Python runtime de khoi chay!
    echo Thu muc "%SHARE_ROOT%\runtime" khong co python.exe va may khong co Python hop le.
    echo Vui long giai nen toan bo file ZIP bao gom thu muc "runtime" di kem.
    pause
    exit /b 1
)

:: 3. Kiểm tra toàn vẹn, staged sync xuống %LOCALAPPDATA% và khởi chạy Edge App Mode
echo [*] Dang khoi chay MA CONTROL FG o che do Zero-Port Edge App Mode...
"!PYTHON_BIN!" -m app.launcher_core --share-dir "!SHARE_ROOT!" %*

if !errorlevel! neq 0 (
    echo.
    echo [CANH BAO] Khong the khoi chay tu Local AppData.
    echo Vui long kiem tra log tai %%LOCALAPPDATA%%\MA-Control-FG\logs\
    pause
    exit /b !errorlevel!
)

echo.
echo [OK] MA CONTROL FG da khoi chay cuc bo thanh cong (Zero-Port).
exit /b 0
