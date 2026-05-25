@echo off
chcp 65001 >nul 2>nul
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"
set "ROOT=%~dp0"
set "ROOT=%ROOT:~0,-1%"
set "LOG=%ROOT%\launch_debug.log"
set "DEBUG=0"
if /i "%~1"=="--debug" set "DEBUG=1"
if /i "%~1"=="--check" goto check
if /i "%~2"=="--check" goto check
if /i "%~1"=="--setup" goto setup
if /i "%~2"=="--setup" goto setup
if "%~1"=="" goto run_gui
if /i "%~1"=="--debug" if "%~2"=="" goto run_gui
if /i "%~1"=="--debug" (
    shift
    goto run_cli
)
if not "%~1"=="" goto run_cli
goto run_gui
:run_gui
call :find_python_gui
if errorlevel 1 goto fail
call :resolve_pythonw
echo [%date% %time%] GUI PYW="!PYW!" >> "%LOG%"
if "!DEBUG!"=="1" (
    "!PYEXE!" "%ROOT%\gui_svod.py"
    set "RC=!ERRORLEVEL!"
    echo [%date% %time%] GUI exit !RC! >> "%LOG%"
    if not "!RC!"=="0" goto gui_failed
    exit /b 0
)
start "" "!PYW!" "%ROOT%\gui_svod.py"
exit /b 0
:gui_failed
echo.
echo [!] GUI failed, code !RC!
if exist "%ROOT%\gui_error.log" (
    echo.
    type "%ROOT%\gui_error.log"
)
echo.
echo Log: %LOG%
goto fail
:run_cli
set "SCRIPT=%ROOT%\build_svod.py"
set "GUI_MODE=0"
call :find_python
if errorlevel 1 goto fail
call :ensure_deps
if errorlevel 1 goto fail
echo.
echo Run: "!PYEXE!" "!SCRIPT!" %*
echo.
"!PYEXE!" "!SCRIPT!" %*
set "RC=!ERRORLEVEL!"
echo.
if not "!RC!"=="0" (
    echo [!] Error !RC!
    pause
    exit /b !RC!
)
echo [OK] Done.
pause
exit /b 0
:check
echo === Check ===
echo Folder: %CD%
call :find_python_gui
if errorlevel 1 call :find_python
if errorlevel 1 goto fail
echo Python: "!PYEXE!"
call :resolve_pythonw
echo GUI launcher: "!PYW!"
"!PYEXE!" --version
set "GUI_MODE=1"
call :ensure_deps
if errorlevel 1 goto fail
"!PYEXE!" -c "import os; os.chdir(r'%ROOT%'); import gui_svod; print('OK')"
if errorlevel 1 goto fail
echo All OK.
pause
exit /b 0
:setup
call :find_python
if errorlevel 1 goto fail
set "OFFLINE=%ROOT%\install_wheels"
if exist "%ROOT%\установка\wheels\openpyxl-*.whl" set "OFFLINE=%ROOT%\установка\wheels"
if exist "%OFFLINE%\openpyxl-*.whl" (
    echo Offline install from wheels
    "!PYEXE!" -m pip install --no-index --find-links="%OFFLINE%" openpyxl
    if not errorlevel 1 goto setup_ok
)
"!PYEXE!" -m pip install -r "%ROOT%\requirements.txt"
if errorlevel 1 goto fail
:setup_ok
echo Done.
pause
exit /b 0
:resolve_pythonw
set "PYW="
for %%F in ("!PYEXE!") do set "PYW=%%~dpFpythonw.exe"
if exist "!PYW!" exit /b 0
set "PYW=!PYEXE:python.exe=pythonw.exe!"
if exist "!PYW!" exit /b 0
echo [WARN] pythonw.exe not found, using python.exe >> "%LOG%"
set "PYW=!PYEXE!"
exit /b 0
:find_python_gui
set "PYEXE="
for /f "tokens=2 delims==" %%t in ('assoc .py 2^>nul') do (
    for /f "tokens=1* delims==" %%f in ('ftype %%t 2^>nul') do (
        set "FTYPELINE=%%f=%%g"
    )
)
if defined FTYPELINE (
    for /f tokens^=1^ delims^=^" %%p in ("!FTYPELINE!") do set "PYEXE=%%~p"
)
if defined PYEXE call :validate_python
if defined PYEXE exit /b 0
call :find_python
exit /b %ERRORLEVEL%
:find_python
set "PYEXE="
where python >nul 2>nul
if not errorlevel 1 (
    for /f "delims=" %%i in ('python -c "import sys; print(sys.executable)" 2^>nul') do set "PYEXE=%%i"
)
call :validate_python
if defined PYEXE exit /b 0
where py >nul 2>nul
if not errorlevel 1 (
    for /f "delims=" %%i in ('py -3 -c "import sys; print(sys.executable)" 2^>nul') do set "PYEXE=%%i"
)
call :validate_python
if defined PYEXE exit /b 0
where python3 >nul 2>nul
if not errorlevel 1 (
    for /f "delims=" %%i in ('python3 -c "import sys; print(sys.executable)" 2^>nul') do set "PYEXE=%%i"
)
call :validate_python
if defined PYEXE exit /b 0
echo [ERROR] Python 3 not found.
exit /b 1
:validate_python
if not defined PYEXE exit /b 0
echo !PYEXE! | findstr /i /c:"WindowsApps" >nul 2>nul
if not errorlevel 1 set "PYEXE="
if not defined PYEXE exit /b 0
"!PYEXE!" -c "import sys" >nul 2>nul
if errorlevel 1 set "PYEXE="
exit /b 0
:ensure_deps
"!PYEXE!" -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
    echo [ERROR] pip install openpyxl  OR  run_svod.bat --setup
    exit /b 1
)
if not "!GUI_MODE!"=="1" exit /b 0
"!PYEXE!" -c "import tkinter; r=tkinter.Tk(); r.destroy()" >nul 2>nul
if errorlevel 1 (
    echo [ERROR] tkinter broken - reinstall Python with tcl/tk
    exit /b 1
)
exit /b 0
:fail
pause
exit /b 1
