"""Скриншот области печати через Playwright (Chromium)."""

from __future__ import annotations

from playwright.sync_api import sync_playwright


def screenshot_element(
    url: str,
    selector: str = "#print-area",
    width: int = 900,
    device_scale_factor: float = 2.0,
) -> bytes:
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
            png = el.screenshot(type="png")
            return png
        finally:
            browser.close()
