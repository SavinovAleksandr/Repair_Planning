@echo off
chcp 65001 > nul
setlocal

rem =============================================================================
rem  Запуск сборщика сводного графика ремонтов (Windows).
rem
rem  По умолчанию — открывает GUI (gui_svod.py) с кнопками.
rem  Если передать аргументы (например --stage all), запустит CLI (build_svod.py).
rem
rem  Двойной клик — и открывается окно с кнопками.
rem =============================================================================

cd /d "%~dp0"

if "%~1"=="" (
    set "SCRIPT=gui_svod.py"
) else (
    set "SCRIPT=build_svod.py"
)

rem Предпочитаемый запуск через py launcher (ставится вместе с Python на Windows).
where py >nul 2>nul
if %errorlevel%==0 (
    py -3 "%SCRIPT%" %*
    goto :after
)

where python >nul 2>nul
if %errorlevel%==0 (
    python "%SCRIPT%" %*
    goto :after
)

echo.
echo [ОШИБКА] На компьютере не найден Python.
echo  1. Скачайте его с https://www.python.org/downloads/
echo  2. При установке отметьте галочку "Add python.exe to PATH".
echo  3. Откройте командную строку и выполните: pip install openpyxl
echo  4. Повторно запустите этот bat-файл.
pause
exit /b 1

:after
rem В GUI-режиме (без аргументов) окно закроется по крестику — пауза не нужна.
if "%~1"=="" goto :eof

echo.
if %errorlevel% neq 0 (
    echo [!] Сборка завершилась с ошибкой. Проверьте сообщение выше.
) else (
    echo [OK] Готово. Файл сохранён рядом с этим bat.
)
pause
