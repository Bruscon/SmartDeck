@echo off
cd /d "%~dp0"
python "%~dp0chrome_tab_switcher.py"
if errorlevel 1 (
    echo Failed to focus Chrome with error code %errorlevel%
    timeout /t 3
)