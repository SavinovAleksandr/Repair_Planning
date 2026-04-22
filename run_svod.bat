@echo off
chcp 65001 > nul
setlocal

rem =============================================================================
rem  Запуск сборщика сводного графика ремонтов (Windows).
rem  Двойной клик — и в этой же папке появится сводный файл.
rem =============================================================================

cd /d "%~dp0"

rem Предпочитаемый запуск через py launcher (ставится вместе с Python на Windows).
where py >nul 2>nul
if %errorlevel%==0 (
    py -3 "build_svod.py" %*
    goto :after
)

where python >nul 2>nul
if %errorlevel%==0 (
    python "build_svod.py" %*
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
echo.
if %errorlevel% neq 0 (
    echo [!] Сборка завершилась с ошибкой. Проверьте сообщение выше.
) else (
    echo [OK] Готово. Файл сохранён рядом с этим bat.
)
pause
