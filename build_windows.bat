@echo off
setlocal EnableExtensions
cd /d "%~dp0"

rem --- Python: сначала py -3 (лаунчер Windows), иначе python в PATH ---
set "PY_CMD="
where py >nul 2>&1 && set "PY_CMD=py -3"
if not defined PY_CMD (
  where python >nul 2>&1 && set "PY_CMD=python"
)
if not defined PY_CMD (
  echo [Ошибка] Не найден Python. Установите Python 3.11+ с python.org и включите "Add python.exe to PATH".
  exit /b 1
)

set "VENV_PY=%~dp0.venv\Scripts\python.exe"
if not exist "%VENV_PY%" (
  echo Создание виртуального окружения .venv...
  %PY_CMD% -m venv .venv
  if errorlevel 1 (
    echo [Ошибка] Не удалось создать venv.
    exit /b 1
  )
)

echo Обновление pip и установка зависимостей...
"%VENV_PY%" -m pip install --upgrade pip
if errorlevel 1 exit /b 1
"%VENV_PY%" -m pip install -r "%~dp0requirements.txt"
if errorlevel 1 exit /b 1

echo Установка Chromium для Playwright ^(PDF/PNG^)...
"%VENV_PY%" -m playwright install chromium
if errorlevel 1 exit /b 1

echo Сборка однофайлового exe ^(PyInstaller^)...
"%VENV_PY%" -m PyInstaller "%~dp0build.spec" --noconfirm
if errorlevel 1 (
  echo [Ошибка] Сборка не удалась.
  exit /b 1
)

echo.
echo Готово: %~dp0dist\schedule_changes.exe
echo База app.db появится рядом с exe после первого запуска.
echo.
pause
endlocal
