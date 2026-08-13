@echo off
REM Chay Personal Toolbox o CHE DO DEBUG: co cua so console de doc traceback.
REM Khac Personal_Tool.bat: dung python.exe (khong phai pythonw.exe) va KHONG
REM dung "start" -> console nay giu nguyen den khi dong app, moi loi in ra day.
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo === Chua co moi truong: hay chay Personal_Tool.bat mot lan truoc ===
    pause
    exit /b 1
)
call ".venv\Scripts\activate.bat"

REM -u: khong dem output (thay log ngay lap tuc)
REM -X dev: bat che do phat trien cua Python (canh bao + kiem tra chat hon)
set PYTHONFAULTHANDLER=1
REM console cua Windows mac dinh khong doc duoc tieng Viet trong log
set PYTHONIOENCODING=utf-8
chcp 65001 >nul
echo === DEBUG MODE — moi loi se in ra cua so nay va ghi vao:
echo     %%APPDATA%%\PersonalToolbox\debug.log
echo.
".venv\Scripts\python.exe" -u -X dev main.py

echo.
echo === App da dong (exit code %ERRORLEVEL%) — doc traceback o tren ===
pause
