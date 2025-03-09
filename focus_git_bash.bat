@echo off
cd /d "%~dp0"
python "%~dp0app_focus.py" "C:\Program Files\Git\git-bash.exe"
if errorlevel 1 (
    echo Failed to focus Git Bash with error code %errorlevel%
    timeout /t 3
)