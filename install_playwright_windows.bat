@echo off
rem Для разработки без exe: ручная установка Chromium в тот же каталог, что и exe
setlocal
cd /d "%~dp0"
set "VENV_PY=%~dp0.venv\Scripts\python.exe"
if not exist "%VENV_PY%" (
  echo Сначала: build_windows.bat или python -m venv .venv
  pause
  exit /b 1
)
"%VENV_PY%" -m pip install "playwright>=1.40.0"
"%VENV_PY%" -c "from core.playwright_bundle import ensure_chromium_at_startup; ensure_chromium_at_startup()"
pause
endlocal
