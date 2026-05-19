@echo off

chcp 65001 >nul 2>nul

setlocal EnableExtensions



rem =============================================================================

rem  Запуск сборщика сводного графика ремонтов (Windows).

rem

rem  Двойной клик — GUI (gui_svod.py).

rem  run_svod.bat --stage all  — консольный режим (build_svod.py).

rem  run_svod.bat --check      — проверка Python и зависимостей.

rem  run_svod.bat --setup      — установка openpyxl.

rem =============================================================================



cd /d "%~dp0"



if /i "%~1"=="--check" goto :check

if /i "%~1"=="--setup" goto :setup



if "%~1"=="" (

    set "SCRIPT=gui_svod.py"

    set "GUI_MODE=1"

) else (

    set "SCRIPT=build_svod.py"

    set "GUI_MODE=0"

)



call :find_python

if errorlevel 1 goto :fail



call :ensure_deps

if errorlevel 1 goto :fail



echo.

echo Запуск: "%PYEXE%" "%SCRIPT%" %*

echo.



"%PYEXE%" "%SCRIPT%" %*

set "RC=%ERRORLEVEL%"



if "%GUI_MODE%"=="1" (

    if not "%RC%"=="0" (

        echo.

        echo [!] GUI завершился с ошибкой (код %RC%). См. сообщение выше.

        pause

    )

    exit /b %RC%

)



echo.

if not "%RC%"=="0" (

    echo [!] Сборка завершилась с ошибкой (код %RC%).

) else (

    echo [OK] Готово. Файл сохранён рядом с этим bat.

)

pause

exit /b %RC%





:check

echo === Проверка окружения ===

echo Папка: %CD%

echo.

call :find_python

if errorlevel 1 goto :fail

echo Python: OK

"%PYEXE%" --version

echo.

set "GUI_MODE=1"

call :ensure_deps

if errorlevel 1 goto :fail

echo.

echo Все проверки пройдены. Можно запускать run_svod.bat двойным кликом.

pause

exit /b 0





:setup

echo === Установка зависимостей ===

call :find_python

if errorlevel 1 goto :fail

"%PYEXE%" -m pip install --upgrade pip

"%PYEXE%" -m pip install -r "%~dp0requirements.txt"

if errorlevel 1 goto :fail

echo.

echo Готово. Запустите: run_svod.bat --check

pause

exit /b 0





:find_python

set "PYEXE="

where py >nul 2>nul

if not errorlevel 1 (

    for /f "delims=" %%i in ('py -3 -c "import sys; print(sys.executable)" 2^>nul') do set "PYEXE=%%i"

    if defined PYEXE exit /b 0

)

where python >nul 2>nul

if not errorlevel 1 (

    for /f "delims=" %%i in ('python -c "import sys; print(sys.executable)" 2^>nul') do set "PYEXE=%%i"

    if defined PYEXE exit /b 0

)

where python3 >nul 2>nul

if not errorlevel 1 (

    for /f "delims=" %%i in ('python3 -c "import sys; print(sys.executable)" 2^>nul') do set "PYEXE=%%i"

    if defined PYEXE exit /b 0

)

echo.

echo [ОШИБКА] На компьютере не найден Python 3.

echo.

echo  1. Скачайте Python 3.10+ с https://www.python.org/downloads/

echo  2. При установке обязательно отметьте "Add python.exe to PATH".

echo  3. Запустите: run_svod.bat --setup

echo  4. Затем: run_svod.bat --check

exit /b 1





:ensure_deps

"%PYEXE%" -c "import openpyxl" >nul 2>nul

if errorlevel 1 (

    echo.

    echo [ОШИБКА] Не установлен пакет openpyxl (нужен для Excel).

    echo Попробуйте: run_svod.bat --setup

    echo Или вручную: "%PYEXE%" -m pip install openpyxl

    exit /b 1

)

if "%GUI_MODE%"=="1" (

    "%PYEXE%" -c "import tkinter" >nul 2>nul

    if errorlevel 1 (

        echo.

        echo [ОШИБКА] Не найден tkinter (нужен для окна с кнопками).

        echo Переустановите Python с python.org и отметьте "tcl/tk and IDLE".

        echo Либо запускайте без GUI: run_svod.bat --stage all

        exit /b 1

    )

)

exit /b 0





:fail

echo.

pause

exit /b 1

