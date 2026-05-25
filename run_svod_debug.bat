@echo off

rem Diagnostic launcher: console stays open, full error output.

cd /d "%~dp0"

cmd /k "%~dp0run_svod.bat" %*

