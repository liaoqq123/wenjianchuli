@echo off
setlocal
cd /d "%~dp0"

if exist ".venv\Scripts\python.exe" (
    set "APP_PYTHON=.venv\Scripts\python.exe"
) else if exist "venv\Scripts\python.exe" (
    set "APP_PYTHON=venv\Scripts\python.exe"
) else (
    set "APP_PYTHON=py -3.14"
)

%APP_PYTHON% -m pip install pyinstaller
%APP_PYTHON% -m PyInstaller --noconfirm --clean --onefile --windowed --name "文件处理器" Main.py

echo.
echo 打包完成：dist\文件处理器.exe
pause
