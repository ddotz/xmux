#!/usr/bin/env python3
"""Render the pixel-aligned SVG source into the Office add-in PNG icon sizes."""

from pathlib import Path

from playwright.sync_api import sync_playwright

ASSETS = Path(__file__).resolve().parent.parent / "public" / "assets"
SIZES = (16, 32, 64, 80)

svg = (ASSETS / "icon.svg").read_text(encoding="utf-8")
html = f"""<!doctype html>
<style>
  html, body {{ margin: 0; width: 100%; height: 100%; background: transparent; }}
  svg {{ display: block; width: 100%; height: 100%; }}
</style>
{svg}
"""

with sync_playwright() as playwright:
    browser = playwright.chromium.launch()
    page = browser.new_page()
    for size in SIZES:
        page.set_viewport_size({"width": size, "height": size})
        page.set_content(html)
        page.screenshot(path=ASSETS / f"icon-{size}.png", omit_background=True)
    browser.close()
