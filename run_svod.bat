@echo off
chcp 65001 >nul 2>nul
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"

if /i "%~1"=="--check" goto check
if /i "%~1"=="--setup" goto setup

if "%~1"=="" (
    set "SCRIPT=gui_svod.py"
    set "GUI_MODE=1"
) else (
    set "SCRIPT=build_svod.py"
    set "GUI_MODE=0"
)

call :find_python
if errorlevel 1 goto fail
call :ensure_deps
if errorlevel 1 goto fail

if "!GUI_MODE!"=="1" goto run_gui
goto run_cli

:run_gui
set "PYW=!PYEXE:python.exe=pythonw.exe!"
if /i not "!PYW!"=="!PYEXE!" if exist "!PYW!" (
    echo GUI: "!PYW!" "%SCRIPT%"
    start "" "!PYW!" "%SCRIPT%"
    exit /b 0
)
echo GUI: "!PYEXE!" "%SCRIPT%"
"!PYEXE!" "%SCRIPT%"
set "RC=!ERRORLEVEL!"
if not "!RC!"=="0" goto fail_with_code
exit /b 0

:run_cli
echo.
echo Run: "!PYEXE!" "%SCRIPT%" %*
echo.
"!PYEXE!" "%SCRIPT%" %*
set "RC=!ERRORLEVEL!"
echo.
if not "!RC!"=="0" (
    echo [!] Error code !RC!
    pause
    exit /b !RC!
)
echo [OK] Done.
pause
exit /b 0

:check
echo === Environment check ===
echo Folder: %CD%
echo.
call :find_python
if errorlevel 1 goto fail
echo Python OK: "!PYEXE!"
"!PYEXE!" --version
echo.
set "GUI_MODE=1"
call :ensure_deps
if errorlevel 1 goto fail
echo.
echo All checks passed.
pause
exit /b 0

:setup
echo === Installing dependencies ===
call :find_python
if errorlevel 1 goto fail
"!PYEXE!" -m pip install --upgrade pip
"!PYEXE!" -m pip install -r "%~dp0requirements.txt"
if errorlevel 1 goto fail
echo.
echo Done. Run: run_svod.bat --check
pause
exit /b 0

:find_python
set "PYEXE="
where py >nul 2>nul
if not errorlevel 1 (
    for /f "delims=" %%i in ('py -3 -c "import sys; print(sys.executable)" 2^>nul') do set "PYEXE=%%i"
)
if defined PYEXE exit /b 0
where python >nul 2>nul
if not errorlevel 1 (
    for /f "delims=" %%i in ('python -c "import sys; print(sys.executable)" 2^>nul') do set "PYEXE=%%i"
)
if defined PYEXE exit /b 0
where python3 >nul 2>nul
if not errorlevel 1 (
    for /f "delims=" %%i in ('python3 -c "import sys; print(sys.executable)" 2^>nul') do set "PYEXE=%%i"
)
if defined PYEXE exit /b 0
echo.
echo [ERROR] Python 3 not found.
echo 1. Install from https://www.python.org/downloads/
echo 2. Check "Add python.exe to PATH"
echo 3. Run: run_svod.bat --setup
exit /b 1

:ensure_deps
"!PYEXE!" -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
    echo.
    echo [ERROR] Package openpyxl is missing.
    echo Run: run_svod.bat --setup
    exit /b 1
)
if not "!GUI_MODE!"=="1" exit /b 0
"!PYEXE!" -c "import tkinter" >nul 2>nul
if errorlevel 1 (
    echo.
    echo [ERROR] tkinter is missing (needed for GUI).
    echo Reinstall Python with "tcl/tk and IDLE" enabled.
    echo Or use CLI: run_svod.bat --stage all
    exit /b 1
)
exit /b 0

:fail_with_code
echo.
echo [!] Failed with code !RC!
goto fail

:fail
echo.
echo Press any key to close...
pause >nul
exit /b 1
