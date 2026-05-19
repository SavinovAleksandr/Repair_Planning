@echo off
rem Diagnostic launcher: window stays open. Use if run_svod.bat closes instantly.
cd /d "%~dp0"
cmd /k run_svod.bat %*
