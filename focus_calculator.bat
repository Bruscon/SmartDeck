@echo off
cd /d "%~dp0"
python "%~dp0app_focus.py" "calc.exe"
if errorlevel 1 (
    echo Failed to focus calculator with error code %errorlevel%
    timeout /t 3
)