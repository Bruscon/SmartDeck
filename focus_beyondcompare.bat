@echo off
cd /d "%~dp0"
python "%~dp0app_focus.py" "C:\Program Files\Beyond Compare 5\BCompare.exe"
if errorlevel 1 (
    echo Failed to focus Beyond Compare with error code %errorlevel%
    timeout /t 3
)