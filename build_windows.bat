@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "PY_CMD="
where py >nul 2>&1 && set "PY_CMD=py -3"
if not defined PY_CMD (
  where python >nul 2>&1 && set "PY_CMD=python"
)
if not defined PY_CMD (
  echo [Ошибка] Нужен Python 3.11+ с python.org
  exit /b 1
)

set "VENV_PY=%~dp0.venv\Scripts\python.exe"
if not exist "%VENV_PY%" (
  echo Создание .venv...
  %PY_CMD% -m venv .venv
  if errorlevel 1 exit /b 1
)

echo Установка зависимостей...
"%VENV_PY%" -m pip install --upgrade pip
"%VENV_PY%" -m pip install -r "%~dp0requirements.txt"
if errorlevel 1 exit /b 1

echo.
echo Сборка одного файла schedule_changes.exe ...
echo Chromium на школьных ПК скачается сам при первом запуске ^(интернет^).
echo.
"%VENV_PY%" -m PyInstaller "%~dp0build.spec" --noconfirm
if errorlevel 1 (
  echo [Ошибка] Сборка не удалась.
  exit /b 1
)

echo.
echo ========================================
echo Готово — отправляйте ОДИН файл:
echo   %~dp0dist\schedule_changes.exe
echo.
echo При первом запуске на ПК: интернет, загрузка Chromium ~150 МБ.
echo Дальше — без интернета. База app.db — рядом с exe.
echo ========================================
echo.
pause
endlocal
