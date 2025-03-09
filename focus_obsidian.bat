@echo off
cd /d "%~dp0"
python "%~dp0app_focus.py" "C:\Users\Nick Brusco\AppData\Local\Programs\Obsidian\Obsidian.exe"
if errorlevel 1 (
    echo Failed to focus Obsidian with error code %errorlevel%
    timeout /t 3
)