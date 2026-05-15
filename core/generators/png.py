"""Скриншот области печати через Playwright (Chromium), если установлен."""

from __future__ import annotations

from core.generators.browser_export import ensure_playwright, playwright_runtime_error


def screenshot_element(
    url: str,
    selector: str = "#print-area",
    width: int = 900,
    device_scale_factor: float = 2.0,
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
                page = browser.new_page(
                    viewport={"width": width, "height": 1200},
                    device_scale_factor=device_scale_factor,
                )
                page.goto(url, wait_until="networkidle", timeout=60000)
                page.wait_for_selector(selector, timeout=30000)
                el = page.locator(selector).first
                return el.screenshot(type="png")
            finally:
                browser.close()
    except Exception as e:
        raise playwright_runtime_error(e) from e
