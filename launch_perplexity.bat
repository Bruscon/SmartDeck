@echo off
cd /d "%~dp0"
python "%~dp0chrome_tab_switcher.py" perplexity.ai
if errorlevel 1 (
    echo Failed to focus Perplexity tab with error code %errorlevel%
    timeout /t 3
)