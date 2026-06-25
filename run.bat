@echo off
REM Chay Personal Toolbox tu source (khong can build .exe).
REM Lan dau: tu tao moi truong + cai thu vien. Cac lan sau: mo thang.
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo === Lan dau chay: tao moi truong va cai thu vien ===
    python -m venv .venv || goto :err
    call ".venv\Scripts\activate.bat"
    python -m pip install --upgrade pip
    python -m pip install -r requirements.txt || goto :err
) else (
    call ".venv\Scripts\activate.bat"
)

REM pythonw = chay GUI khong kem cua so console den
start "" ".venv\Scripts\pythonw.exe" main.py
exit /b 0

:err
echo.
echo *** Loi: kiem tra da cai Python chua (python --version) ***
pause
exit /b 1
