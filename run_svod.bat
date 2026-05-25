@echo off

chcp 65001 >nul 2>nul

setlocal EnableExtensions EnableDelayedExpansion

cd /d "%~dp0"

set "ROOT=%~dp0"

set "ROOT=%ROOT:~0,-1%"



if /i "%~1"=="--check" goto check

if /i "%~1"=="--setup" goto setup



if "%~1"=="" (

    set "SCRIPT=%ROOT%\gui_svod.py"

    set "GUI_MODE=1"

) else (

    set "SCRIPT=%ROOT%\build_svod.py"

    set "GUI_MODE=0"

)



call :find_python

if errorlevel 1 goto fail

call :ensure_deps

if errorlevel 1 goto fail



if "!GUI_MODE!"=="1" goto run_gui

goto run_cli



:run_gui

rem НЕ используем «start» — иначе bat закрывается сразу и ошибки pythonw не видны.

set "PYW=!PYEXE:python.exe=pythonw.exe!"

if /i "!PYW!"=="!PYEXE!" set "PYW="

if defined PYW if not exist "!PYW!" set "PYW="



if defined PYW (

    echo Starting GUI ^(pythonw^)...

    "!PYW!" "!SCRIPT!"

) else (

    echo Starting GUI ^(python^)...

    "!PYEXE!" "!SCRIPT!"

)

set "RC=!ERRORLEVEL!"

if not "!RC!"=="0" goto gui_failed

exit /b 0



:gui_failed

echo.

echo [!] GUI failed with code !RC!

if exist "%ROOT%\gui_error.log" (

    echo.

    echo --- gui_error.log ---

    type "%ROOT%\gui_error.log"

    echo --- end ---

)

echo.

echo Try: run_svod_debug.bat

goto fail



:run_cli

echo.

echo Run: "!PYEXE!" "!SCRIPT!" %*

echo.

"!PYEXE!" "!SCRIPT!" %*

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

echo Python: "!PYEXE!"

"!PYEXE!" --version

echo.

set "GUI_MODE=1"

call :ensure_deps

if errorlevel 1 goto fail

echo.

echo Testing GUI import...

"!PYEXE!" -c "import os; os.chdir(r'%ROOT%'); import gui_svod; print('gui_svod OK')"

if errorlevel 1 goto fail

echo.

echo All checks passed. Double-click run_svod.bat to open GUI.

pause

exit /b 0



:setup

echo === Installing dependencies ===

call :find_python

if errorlevel 1 goto fail

set "OFFLINE=%ROOT%\установка\wheels"

if exist "%OFFLINE%\openpyxl-*.whl" (

    echo Offline wheels found in установка\wheels

    "!PYEXE!" -m pip install --no-index --find-links="%OFFLINE%" openpyxl

    if not errorlevel 1 goto setup_ok

    echo Offline install failed, trying online...

)

"!PYEXE!" -m pip install --upgrade pip

"!PYEXE!" -m pip install -r "%ROOT%\requirements.txt"

if errorlevel 1 goto fail

:setup_ok

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

echo 2. Check "Add python.exe to PATH" and "tcl/tk and IDLE"

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

"!PYEXE!" -c "import tkinter; tkinter.Tk().withdraw()" >nul 2>nul

if errorlevel 1 (

    echo.

    echo [ERROR] tkinter is not working (needed for GUI window).

    echo Reinstall Python from python.org with "tcl/tk and IDLE" enabled.

    echo Or use CLI: run_svod.bat --stage all

    exit /b 1

)

exit /b 0



:fail

echo.

echo Press any key to close...

pause >nul

exit /b 1

