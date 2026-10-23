@echo off
REM Сборка exe. Требует: pip install pyinstaller (без прав администратора)
cd /d "%~dp0"
python -m PyInstaller --noconfirm app\build_exe.spec
if %errorlevel% neq 0 (
  echo Ошибка сборки. Установите PyInstaller:  pip install pyinstaller
  pause
  exit /b 1
)
echo Готово: dist\excel-compare\excel-compare.exe
pause
