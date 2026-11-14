"""
Minimal CDP background test for Gemini web app.
Connects to an already-running Chrome instance on 127.0.0.1:9222,
finds (or opens) a Gemini tab without stealing OS focus, sends "hi",
waits for the reply to finish, and prints it as UTF-8 to stdout.
"""

from __future__ import annotations

import json
import sys
import time
import urllib.request
from pathlib import Path

from playwright.sync_api import sync_playwright

CDP_URL = "http://127.0.0.1:9222"
GEMINI_URL = "https://gemini.google.com/app"
PROMPT_TEXT = "hi"


def is_cdp_reachable() -> bool:
    try:
        with urllib.request.urlopen(f"{CDP_URL}/json/version", timeout=1.5) as resp:
            return resp.status == 200
    except Exception:
        return False


def find_or_open_gemini_page(context):
    for page in context.pages:
        if "gemini.google.com" in (page.url or "").lower():
            return page
    page = context.new_page()
    page.goto(GEMINI_URL, wait_until="domcontentloaded", timeout=20000)
    return page


def get_response_texts(page) -> list[str]:
    return page.evaluate(
        """() => {
            const nodes = document.querySelectorAll('model-response message-content, model-response .markdown, message-content');
            const out = [];
            for (const n of nodes) {
                const t = (n.innerText || '').trim();
                if (t) out.push(t);
            }
            return out;
        }"""
    )


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    if not is_cdp_reachable():
        print(json.dumps({"ok": False, "error": "CDP_NOT_RUNNING"}))
        return 2

    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(CDP_URL)
        if not browser.contexts:
            print(json.dumps({"ok": False, "error": "NO_BROWSER_CONTEXT"}))
            return 3

        context = browser.contexts[0]
        page = find_or_open_gemini_page(context)

        composer_sel = 'div.ql-editor[contenteditable="true"], rich-textarea [contenteditable="true"], [role="textbox"][contenteditable="true"]'
        try:
            composer = page.locator(composer_sel).first
            composer.wait_for(state="visible", timeout=15000)
        except Exception:
            print(json.dumps({"ok": False, "error": "COMPOSER_NOT_FOUND (check if logged in to Gemini)"}))
            return 4

        before_texts = get_response_texts(page)
        before_count = len(before_texts)

        # Click & insert text purely via CDP inside the tab (no OS focus/mouse/clipboard)
        composer.click()
        page.keyboard.insert_text(PROMPT_TEXT)
        time.sleep(0.15)
        page.keyboard.press("Enter")

        # Wait for a new response to appear and stabilize
        deadline = time.time() + 25.0
        last_text = ""
        stable_ticks = 0

        while time.time() < deadline:
            time.sleep(0.4)
            current_texts = get_response_texts(page)
            if len(current_texts) > before_count:
                latest = current_texts[-1]
                # Check if stop-generating button is still visible
                is_generating = page.evaluate(
                    """() => !!document.querySelector('button[aria-label*="Stop" i], mat-icon[data-mat-icon-name="stop"]')"""
                )
                if latest and latest == last_text and not is_generating:
                    stable_ticks += 1
                    if stable_ticks >= 2:
                        print(json.dumps({"ok": True, "reply": latest}, ensure_ascii=False))
                        return 0
                else:
                    last_text = latest
                    stable_ticks = 0

        if last_text:
            print(json.dumps({"ok": True, "reply": last_text}, ensure_ascii=False))
            return 0

        print(json.dumps({"ok": False, "error": "TIMEOUT_WAITING_FOR_REPLY"}))
        return 5


if __name__ == "__main__":
    raise SystemExit(main())
