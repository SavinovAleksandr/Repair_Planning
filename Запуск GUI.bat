@echo off
cd /d "%~dp0"
python "%~dp0gui_svod.py"
if errorlevel 1 pause
