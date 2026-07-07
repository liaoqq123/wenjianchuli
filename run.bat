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

%APP_PYTHON% Main.py
