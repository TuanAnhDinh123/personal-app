@echo off
REM Chay Personal Toolbox tu source (khong can build .exe).
REM Lan dau: tu tao moi truong + cai thu vien. Cac lan sau: mo thang.
cd /d "%~dp0"

REM .venv ghi cung duong dan ban Python da tao ra no; ban do bi go/doi cho thi
REM .venv thanh rac (bao "did not find executable at ...") -> xoa di tao lai
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" -c "" >nul 2>nul || rmdir /s /q ".venv"
)

if not exist ".venv\Scripts\python.exe" (
    echo === Lan dau chay: tao moi truong ===
    python -m venv .venv || goto :err
    call ".venv\Scripts\activate.bat"
    python -m pip install --upgrade pip
) else (
    call ".venv\Scripts\activate.bat"
)

REM Luon dong bo thu vien: cai phan con thieu (du roi se chay rat nhanh)
echo === Kiem tra / cai thu vien ===
python -m pip install -r requirements.txt || goto :err

REM pythonw = chay GUI khong kem cua so console den
start "" ".venv\Scripts\pythonw.exe" main.py
exit /b 0

:err
echo.
echo *** Loi: kiem tra da cai Python chua (python --version) ***
echo *** Luc cai Python nho TICK "Add python.exe to PATH" o man hinh dau tien ***
pause
exit /b 1
