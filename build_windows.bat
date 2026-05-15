@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "PY_CMD="
where py >nul 2>&1 && set "PY_CMD=py -3"
if not defined PY_CMD (
  where python >nul 2>&1 && set "PY_CMD=python"
)
if not defined PY_CMD (
  echo [Ошибка] Не найден Python 3.11+. Установите с python.org, включите PATH.
  exit /b 1
)

set "VENV_PY=%~dp0.venv\Scripts\python.exe"
if not exist "%VENV_PY%" (
  echo Создание .venv...
  %PY_CMD% -m venv .venv
  if errorlevel 1 exit /b 1
)

echo Установка основных зависимостей ^(Flask, openpyxl^)...
"%VENV_PY%" -m pip install --upgrade pip
if errorlevel 1 exit /b 1
"%VENV_PY%" -m pip install -r "%~dp0requirements.txt"
if errorlevel 1 exit /b 1

echo.
echo Playwright для PDF/PNG — по желанию. Запустите install_playwright_windows.bat после сборки.
echo Сборка exe не требует Chromium.
echo.

echo PyInstaller...
"%VENV_PY%" -m pip install pyinstaller>=6.3.0
if errorlevel 1 exit /b 1
"%VENV_PY%" -m PyInstaller "%~dp0build.spec" --noconfirm
if errorlevel 1 (
  echo [Ошибка] Сборка не удалась.
  exit /b 1
)

echo.
echo Готово: %~dp0dist\schedule_changes.exe
echo Excel-экспорт работает сразу. Для PDF/PNG на этом ПК: install_playwright_windows.bat
echo.
pause
endlocal
