@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul
cd /d "%~dp0"
title HE THONG CHUAN HOA BAO CAO WORD (.DOCX) - TU DONG
color 0A

echo ==============================================================================
echo        HE THONG CHUAN HOA BAO CAO WORD TU DONG - CHUAN VIET NAM VA AI
echo ==============================================================================
echo.

:: 1. Kiem tra Python
echo [1/4] Dang kiem tra moi truong Python...
set PYTHON_CMD=
where python >nul 2>&1
if %ERRORLEVEL% equ 0 (
    set PYTHON_CMD=python
    goto PYTHON_FOUND
)

where py >nul 2>&1
if %ERRORLEVEL% equ 0 (
    set PYTHON_CMD=py
    goto PYTHON_FOUND
)

echo [LOI] Khong tim thay Python tren may tinh cua ban!
echo Vui long cai dat Python 3.9 tro len va tich chon "Add Python to PATH".
echo Tai Python tai: https://www.python.org/downloads/
echo.
pause
exit /b 1

:PYTHON_FOUND
%PYTHON_CMD% --version
echo [OK] Da tim thay Python.
echo.

:: 2. Kiem tra file cau hinh .env
echo [2/4] Kiem tra file cau hinh .env...
if not exist ".env" (
    if exist ".env.example" (
        copy ".env.example" ".env" >nul
        echo [INFO] Da tao file .env tu .env.example.
    )
) else (
    echo [OK] File cau hinh .env da san sang.
)
echo.

:: 3. Kiem tra va tao Virtual Environment
echo [3/4] Kiem tra moi truong ao venv...
if not exist "venv\Scripts\activate.bat" (
    echo [INFO] Dang tao moi truong ao Python venv...
    %PYTHON_CMD% -m venv venv
    if %ERRORLEVEL% neq 0 (
        echo [CANH BAO] Khong the tao venv, he thong se su dung Python mac dinh.
        goto RUN_DIRECT
    )
)

call venv\Scripts\activate.bat
if %ERRORLEVEL% neq 0 (
    echo [CANH BAO] Khong the kich hoat venv, chuyen sang chay truc tiep.
    goto RUN_DIRECT
)

echo [OK] Da kich hoat moi truong ao venv.
echo.

:: 4. Cai dat thu vien neu chua co
if not exist "venv\.installed" (
    echo [INFO] Dang cai dat cac thu vien can thiet tu requirements.txt...
    python -m pip install --upgrade pip
    pip install -r requirements.txt
    if %ERRORLEVEL% equ 0 (
        echo ok > venv\.installed
        echo [OK] Cai dat thu vien hoan tat.
    ) else (
        echo [CANH BAO] Cai dat requirements co the co loi, thu tiep tuc...
    )
) else (
    echo [OK] Cac thu vien da duoc cai dat day du.
)
echo.

:: 5. Khoi dong Streamlit Web App
echo [4/4] Dang khoi dong he thong giao dien Web...
echo [INFO] Ung dung se tu dong mo tren trinh duyet tai http://localhost:8501
echo.
python -m streamlit run app.py
goto SCRIPT_END

:RUN_DIRECT
echo [INFO] Dang kiem tra va cai dat thu vien cho Python he thong...
%PYTHON_CMD% -m pip install -r requirements.txt
echo [INFO] Dang khoi dong Web App bang Python he thong...
%PYTHON_CMD% -m streamlit run app.py

:SCRIPT_END
if %ERRORLEVEL% neq 0 (
    echo.
    echo [THONG BAO] Ung dung da dung lai hoac gap su co.
    pause
)
