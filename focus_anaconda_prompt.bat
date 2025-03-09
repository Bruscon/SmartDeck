@echo off
cd /d "%~dp0"
python "%~dp0app_focus.py" "C:\Users\Nick Brusco\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Anaconda3 (64-bit)\Anaconda Prompt (Anaconda3).lnk"
if errorlevel 1 (
    echo Failed to focus anaconda prompt with error code %errorlevel%
    timeout /t 3
)