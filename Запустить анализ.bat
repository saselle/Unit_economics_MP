@echo off
chcp 65001 >nul
rem Запуск всего анализа в один клик: откройте этот файл двойным щелчком.
cd /d "%~dp0"

set PY=py
where py >nul 2>nul || set PY=python

%PY% scripts\run_all.py %*

echo.
echo Готово. Отчёты лежат в папке reports.
pause
