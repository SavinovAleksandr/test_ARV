@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Запуск: Модуль оценки АРВ.exe
echo.
"Модуль оценки АРВ.exe"
set EXIT_CODE=%ERRORLEVEL%
echo.
echo Код выхода: %EXIT_CODE%
echo.
pause
