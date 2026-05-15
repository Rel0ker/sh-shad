"""HTML → PDF: WeasyPrint при наличии библиотек, иначе Playwright (Chromium)."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from flask import Flask

from core.generators.browser_export import ensure_playwright, playwright_runtime_error


def _pdf_weasyprint(app: Flask, template_name: str, context: dict[str, Any]) -> bytes:
    from weasyprint import HTML

    with app.app_context():
        html_str = app.jinja_env.get_template(template_name).render(**context)
    base_url = str(Path(app.root_path).resolve())
    return HTML(string=html_str, base_url=base_url).write_pdf()


def render_pdf_via_playwright(
    url: str,
    *,
    landscape: bool = False,
    margin_mm: int = 8,
) -> bytes:
    ensure_playwright()
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as e:
        from core.generators.browser_export import playwright_missing_message

        raise RuntimeError(playwright_missing_message()) from e

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            try:
                page = browser.new_page()
                page.goto(url, wait_until="networkidle", timeout=60000)
                page.wait_for_selector("#print-area", timeout=30000)
                return page.pdf(
                    print_background=True,
                    format="A4",
                    landscape=landscape,
                    prefer_css_page_size=True,
                    margin={
                        "top": f"{margin_mm}mm",
                        "bottom": f"{margin_mm}mm",
                        "left": f"{margin_mm}mm",
                        "right": f"{margin_mm}mm",
                    },
                )
            finally:
                browser.close()
    except Exception as e:
        raise playwright_runtime_error(e) from e


def render_pdf_bytes(
    app: Flask,
    template_name: str,
    context: dict[str, Any],
    *,
    fallback_url: str | None = None,
    fallback_landscape: bool = False,
    fallback_margin_mm: int = 8,
) -> bytes:
    """
    Сначала WeasyPrint (если установлен). Иначе — Chromium через Playwright.
    Если ничего нет — понятная ошибка (рекомендуется Excel).
    """
    try:
        return _pdf_weasyprint(app, template_name, context)
    except (OSError, ImportError, RuntimeError, Exception):
        if fallback_url:
            return render_pdf_via_playwright(
                fallback_url,
                landscape=fallback_landscape,
                margin_mm=fallback_margin_mm,
            )
        from core.generators.browser_export import playwright_missing_message

        raise RuntimeError(playwright_missing_message()) from None
