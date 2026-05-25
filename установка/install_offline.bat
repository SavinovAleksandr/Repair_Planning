@echo off
chcp 65001 >nul 2>nul
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"

echo === Offline install: openpyxl ===
echo Folder: %CD%
echo.

call :find_python
if errorlevel 1 goto fail

set "WHEELS=%~dp0wheels"
if not exist "%WHEELS%\openpyxl-*.whl" (
    echo [ERROR] No wheels in folder:
    echo   %WHEELS%
    echo Copy wheel files there or run download_wheels.sh on dev machine.
    goto fail
)

echo Python: "!PYEXE!"
"!PYEXE!" --version
echo.
echo Installing from local wheels (no internet)...
"!PYEXE!" -m pip install --no-index --find-links="%WHEELS%" openpyxl
if errorlevel 1 goto fail

echo.
echo Checking import...
"!PYEXE!" -c "import openpyxl; print('openpyxl', openpyxl.__version__, 'OK')"
if errorlevel 1 goto fail

echo.
echo [OK] Libraries installed.
echo Now run from project folder: run_svod.bat --check
echo.
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
echo [ERROR] Python 3 not found.
echo Install Python 3.10+ from https://www.python.org/downloads/
echo Check "Add python.exe to PATH" during setup.
exit /b 1

:fail
echo.
pause
exit /b 1
