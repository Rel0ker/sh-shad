"""
Chromium для PDF/PNG: один exe, браузер в профиле пользователя (не рядом с файлом).
При каждом запуске — проверка; если нет — скачивание через встроенный Playwright.
"""

from __future__ import annotations

import os
import subprocess
import sys
import time
from pathlib import Path

_APP_DIR_NAME = "ScheduleChanges"
_INSTALL_ATTEMPTS = 3
_INSTALL_TIMEOUT_SEC = 900


def persistent_browsers_dir() -> Path:
    """Каталог вне папки с exe — только один файл на рабочий стол/флешку."""
    if sys.platform == "win32":
        base = Path(
            os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local")
        )
    elif sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support"
    else:
        base = Path.home() / ".local" / "share"
    return base / _APP_DIR_NAME / "ms-playwright"


def configure_playwright_environment() -> Path:
    path = persistent_browsers_dir()
    path.mkdir(parents=True, exist_ok=True)
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(path.resolve())
    return path


def chromium_is_ready(browsers_dir: Path | None = None) -> bool:
    root = browsers_dir or persistent_browsers_dir()
    if not root.is_dir():
        return False
    wanted = {
        "chrome.exe",
        "chrome",
        "chrome-headless-shell.exe",
        "chrome-headless-shell",
        "Google Chrome for Testing",
    }
    for f in root.rglob("*"):
        if f.is_file() and f.name in wanted:
            return True
    return False


def bundled_browsers_ready() -> bool:
    configure_playwright_environment()
    return chromium_is_ready()


def _log(msg: str) -> None:
    print(msg, flush=True)


def _playwright_importable() -> bool:
    try:
        import playwright  # noqa: F401

        return True
    except ImportError:
        return False


def _run_playwright_install(browsers_dir: Path) -> bool:
    from playwright._impl._driver import compute_driver_executable, get_driver_env

    driver_executable, driver_cli = compute_driver_executable()
    env = get_driver_env()
    env["PLAYWRIGHT_BROWSERS_PATH"] = str(browsers_dir.resolve())

    proc = subprocess.run(
        [driver_executable, driver_cli, "install", "chromium"],
        env=env,
        timeout=_INSTALL_TIMEOUT_SEC,
    )
    return proc.returncode == 0 and chromium_is_ready(browsers_dir)


def install_chromium_with_retries(browsers_dir: Path) -> bool:
    lock = browsers_dir.parent / ".chromium_install.lock"
    try:
        if lock.exists() and time.time() - lock.stat().st_mtime < _INSTALL_TIMEOUT_SEC:
            _log("Другой процесс уже скачивает Chromium, ожидание…")
            for _ in range(120):
                time.sleep(2)
                if chromium_is_ready(browsers_dir):
                    return True
            return chromium_is_ready(browsers_dir)

        lock.parent.mkdir(parents=True, exist_ok=True)
        lock.write_text(str(os.getpid()), encoding="utf-8")

        for attempt in range(1, _INSTALL_ATTEMPTS + 1):
            _log(f"Загрузка Chromium, попытка {attempt}/{_INSTALL_ATTEMPTS}…")
            try:
                if _run_playwright_install(browsers_dir):
                    return True
            except subprocess.TimeoutExpired:
                _log("Превышено время ожидания загрузки.")
            except OSError as e:
                _log(f"Ошибка запуска установщика: {e}")
            if attempt < _INSTALL_ATTEMPTS:
                time.sleep(3)
        return False
    finally:
        try:
            lock.unlink(missing_ok=True)
        except OSError:
            pass


def ensure_chromium_at_startup() -> bool:
    """
    Вызывать при старте приложения (main).
    Возвращает True, если Chromium готов к PDF/PNG.
    """
    browsers_dir = configure_playwright_environment()

    if chromium_is_ready(browsers_dir):
        return True

    if not _playwright_importable():
        if getattr(sys, "frozen", False):
            _log(
                "[!] PDF/PNG недоступны: сборка без Playwright. "
                "Excel работает."
            )
        return False

    _log("")
    _log("=== Chromium для экспорта PDF/PNG ===")
    _log("Браузер не найден. Нужен интернет (~150 МБ, один раз).")
    _log(f"Каталог: {browsers_dir}")
    _log("")

    if install_chromium_with_retries(browsers_dir):
        _log("Chromium готов.\n")
        return True

    _log("")
    _log(
        "[!] Не удалось установить Chromium (сеть, прокси, антивирус). "
        "Excel и ввод данных работают; PDF/PNG — после успешной загрузки."
    )
    _log("Перезапустите программу при появлении интернета.\n")
    return False
