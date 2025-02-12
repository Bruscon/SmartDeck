@echo off
cd /d "%~dp0"
python "%~dp0app_focus.py" "C:\Program Files\Sublime Text\sublime_text.exe"
if errorlevel 1 (
    echo Failed to focus Sublime Text with error code %errorlevel%
    timeout /t 3
)