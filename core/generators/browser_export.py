"""Проверка Playwright/Chromium и понятные сообщения об ошибках (Windows и .exe)."""

from __future__ import annotations


def playwright_missing_message() -> str:
    return (
        "Для PDF/PNG нужен Chromium (Playwright).\n\n"
        "На Windows в папке проекта запустите install_playwright_windows.bat\n"
        "или в командной строке:\n"
        "  .venv\\Scripts\\python.exe -m pip install playwright\n"
        "  .venv\\Scripts\\python.exe -m playwright install chromium\n\n"
        "Если скачивание не идёт — прокси, антивирус или нет интернета.\n"
        "Пока используйте экспорт Excel (XLSX) — он работает без браузера."
    )


def ensure_playwright() -> None:
    try:
        import playwright  # noqa: F401
    except ImportError as e:
        raise RuntimeError(playwright_missing_message()) from e


def playwright_runtime_error(exc: BaseException) -> RuntimeError:
    text = str(exc).lower()
    if "executable doesn't exist" in text or "browser" in text and "chromium" in text:
        return RuntimeError(
            playwright_missing_message()
            + "\n\nДетали: Chromium не найден — выполните playwright install chromium."
        )
    return RuntimeError(playwright_missing_message() + f"\n\nДетали: {exc}")
