@echo off
rem Запуск GUI excel-compare (см. ТЗ2, раздел 10)
cd /d "%~dp0"
setlocal enabledelayedexpansion

set "PY="
rem 1) python на PATH
python --version >nul 2>&1 && set "PY=python"
if not defined PY py -3 --version >nul 2>&1 && set "PY=py -3"
if not defined PY (
    echo [ОШИБКА] Python не найден в PATH.
    echo Установите Python 3.14+ с python.org и поставьте галочку "Add to PATH".
    pause & exit /b 1
)

"%PY%" -c "import PySide6" >nul 2>&1
if errorlevel 1 (
    echo PySide6 не установлен — устанавливаю в user-окружение ^(без прав админа^)...
    "%PY%" -m pip install --user PySide6 || "%PY%" -m pip install PySide6
)

"%PY%" app\main_gui.py %*
if errorlevel 1 pause
