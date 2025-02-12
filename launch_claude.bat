@echo off
cd /d "%~dp0"
python "%~dp0chrome_tab_switcher.py" claude.ai
if errorlevel 1 (
    echo Failed to focus Claude tab with error code %errorlevel%
    timeout /t 3
)