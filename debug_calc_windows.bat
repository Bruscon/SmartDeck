@echo off
cd /d "%~dp0"
python "%~dp0app_focus.py" calc.exe --debug-windows
pause