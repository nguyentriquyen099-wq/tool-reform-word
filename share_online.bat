@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul
cd /d "%~dp0"
title TAO DUONG LINK ONLINE CHO MOI NGUOI DUNG
color 0B

echo ==============================================================================
echo       KHOI DONG DUONG LINK ONLINE CONG KHAI - CHUAN HOA WORD (.DOCX)
echo ==============================================================================
echo.

:: 1. Kiem tra xem Streamlit (port 8501) da chay chua
powershell -Command "try { $r = Invoke-WebRequest -Uri 'http://localhost:8501' -UseBasicParsing -TimeoutSec 2; exit 0 } catch { exit 1 }" >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo [INFO] Ung dung Web chua bat. Dang tu dong khoi dong web truoc...
    start "HE THONG CHUAN HOA BAO CAO" run.bat
    echo Dang cho web khoi dong trong 5 giay...
    timeout /t 5 /nobreak >nul
)

:: 2. Chay script chia se online
if exist "venv\Scripts\python.exe" (
    venv\Scripts\python.exe share_online.py
) else (
    python share_online.py
)

pause
