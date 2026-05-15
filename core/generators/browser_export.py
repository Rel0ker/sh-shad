"""Сообщения об ошибках PDF/PNG."""

from __future__ import annotations

import sys

from core.playwright_bundle import (
    bundled_browsers_ready,
    persistent_browsers_dir,
)


def playwright_missing_message() -> str:
    if getattr(sys, "frozen", False):
        return (
            "Chromium ещё не установлен или загрузка не завершилась.\n\n"
            "1. Закройте программу, проверьте интернет.\n"
            "2. Запустите schedule_changes.exe снова — загрузка начнётся автоматически.\n"
            f"3. Каталог браузера: {persistent_browsers_dir()}\n\n"
            "Пока используйте экспорт Excel."
        )
    if not bundled_browsers_ready():
        return (
            "Для PDF/PNG нужен Chromium.\n"
            "Запустите приложение ещё раз (скачает автоматически) или:\n"
            "  playwright install chromium\n\n"
            "Или используйте экспорт Excel."
        )
    return (
        "Не удалось запустить Chromium.\n"
        "Перезапустите программу или используйте экспорт Excel."
    )


def ensure_playwright() -> None:
    try:
        import playwright  # noqa: F401
    except ImportError as e:
        raise RuntimeError(playwright_missing_message()) from e
    if not bundled_browsers_ready():
        raise RuntimeError(playwright_missing_message())


def playwright_runtime_error(exc: BaseException) -> RuntimeError:
    return RuntimeError(playwright_missing_message() + f"\n\nДетали: {exc}")
