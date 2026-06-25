@echo off
REM Đóng gói Personal Toolbox thành 1 file .exe (double-click file này để build).
cd /d "%~dp0"

echo === Cai dat dependencies + PyInstaller ===
python -m pip install -r requirements.txt pyinstaller || goto :err

echo === Don ban build cu ===
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo === Build ===
python -m PyInstaller --noconfirm --clean personal_app.spec || goto :err

echo.
echo === XONG. File o: dist\personal_app.exe ===
pause
exit /b 0

:err
echo.
echo *** BUILD THAT BAI - xem log o tren ***
pause
exit /b 1
