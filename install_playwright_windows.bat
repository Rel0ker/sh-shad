@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "VENV_PY=%~dp0.venv\Scripts\python.exe"
if not exist "%VENV_PY%" (
  echo Сначала запустите build_windows.bat или создайте .venv вручную.
  pause
  exit /b 1
)

echo Установка пакета playwright...
"%VENV_PY%" -m pip install "playwright>=1.40.0"
if errorlevel 1 (
  echo [Ошибка] pip install playwright
  pause
  exit /b 1
)

rem Папка браузеров в профиле пользователя — меньше проблем с правами в Program Files
set "PLAYWRIGHT_BROWSERS_PATH=%USERPROFILE%\ms-playwright"
echo PLAYWRIGHT_BROWSERS_PATH=%PLAYWRIGHT_BROWSERS_PATH%

echo Скачивание Chromium ^(нужен интернет, ~150 МБ^)...
"%VENV_PY%" -m playwright install chromium
if errorlevel 1 (
  echo.
  echo [Не удалось] Частые причины на Windows:
  echo   - нет интернета или блокирует прокси/антивирус
  echo   - запустите cmd от имени администратора и повторите
  echo   - попробуйте: set HTTPS_PROXY=... если есть корпоративный прокси
  echo.
  echo Excel-экспорт в приложении работает без Playwright.
  pause
  exit /b 1
)

echo.
echo Готово. PDF и PNG в приложении должны заработать.
echo Браузеры лежат в: %PLAYWRIGHT_BROWSERS_PATH%
echo.
pause
endlocal
