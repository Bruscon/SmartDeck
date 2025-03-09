@echo off
cd /d "%~dp0"
python "%~dp0app_focus.py" "explorer.exe"
if errorlevel 1 (
    echo Failed to focus explorer with error code %errorlevel%
    timeout /t 3
)