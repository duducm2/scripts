"""
Minimal CDP background test for Gemini (consumer) and Gemini Enterprise (Vertex AI Search).
Connects to an already-running Chrome instance on 127.0.0.1:9222,
finds (or opens) the companion tab without stealing OS focus, sends "hi",
waits for the reply to finish, and prints it as UTF-8 to stdout.
"""

from __future__ import annotations

import argparse
import json
import sys
import time
import urllib.request

from playwright.sync_api import sync_playwright

CDP_URL = "http://127.0.0.1:9222"
CONSUMER_URL = "https://gemini.google.com/app"
ENTERPRISE_URL = "https://vertexaisearch.cloud.google.com/u/1/eu/home/cid/bcb383f1-26d8-41fd-9a55-623f7e93de92?pli=1"
PROMPT_TEXT = "hi"


def is_cdp_reachable() -> bool:
    try:
        with urllib.request.urlopen(f"{CDP_URL}/json/version", timeout=1.5) as resp:
            return resp.status == 200
    except Exception:
        return False


def find_or_open_page(context, enterprise: bool):
    needle = "vertexaisearch.cloud.google.com" if enterprise else "gemini.google.com"
    target_url = ENTERPRISE_URL if enterprise else CONSUMER_URL
    for page in context.pages:
        if needle in (page.url or "").lower():
            return page
    page = context.new_page()
    page.goto(target_url, wait_until="domcontentloaded", timeout=25000)
    return page


def get_response_texts(page, enterprise: bool) -> list[str]:
    if not enterprise:
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

    # Enterprise (Vertex AI Search / AskBosch) uses custom web components / shadow DOM.
    # Walk open shadow roots recursively to find response/markdown blocks.
    return page.evaluate(
        """() => {
            function collect(root, out) {
                if (!root) return;
                const els = root.querySelectorAll('ucs-markdown, .markdown, [data-testid*="response" i], [class*="answer" i], [class*="response" i], message-content');
                for (const el of els) {
                    const t = (el.innerText || '').trim();
                    if (t) out.push(t);
                }
                for (const el of root.querySelectorAll('*')) {
                    if (el.shadowRoot) collect(el.shadowRoot, out);
                }
            }
            const out = [];
            collect(document, out);
            return out;
        }"""
    )


def has_stop_button(page) -> bool:
    return page.evaluate(
        """() => {
            function check(root) {
                if (!root) return false;
                if (root.querySelector('button[aria-label*="Stop" i], mat-icon[data-mat-icon-name="stop"], [aria-label*="Stop generating" i]'))
                    return true;
                for (const el of root.querySelectorAll('*')) {
                    if (el.shadowRoot && check(el.shadowRoot)) return true;
                }
                return false;
            }
            return check(document);
        }"""
    )


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    parser = argparse.ArgumentParser()
    parser.add_argument("--enterprise", action="store_true")
    args = parser.parse_args()

    if not is_cdp_reachable():
        print(json.dumps({"ok": False, "error": "CDP_NOT_RUNNING"}))
        return 2

    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(CDP_URL)
        if not browser.contexts:
            print(json.dumps({"ok": False, "error": "NO_BROWSER_CONTEXT"}))
            return 3

        context = browser.contexts[0]
        page = find_or_open_page(context, args.enterprise)

        # Playwright locators pierce open Shadow DOM automatically
        composer_sel = (
            'div.ql-editor[contenteditable="true"], '
            'rich-textarea [contenteditable="true"], '
            '[role="textbox"][contenteditable="true"], '
            'textarea[aria-label], '
            'textarea'
        )
        try:
            composer = page.locator(composer_sel).first
            composer.wait_for(state="visible", timeout=15000)
        except Exception:
            print(json.dumps({"ok": False, "error": "COMPOSER_NOT_FOUND (check if signed in)"}))
            return 4

        before_count = len(get_response_texts(page, args.enterprise))

        composer.click()
        page.keyboard.insert_text(PROMPT_TEXT)
        time.sleep(0.15)
        page.keyboard.press("Enter")

        deadline = time.time() + 25.0
        last_text = ""
        stable_ticks = 0

        while time.time() < deadline:
            time.sleep(0.4)
            current_texts = get_response_texts(page, args.enterprise)
            if len(current_texts) > before_count:
                latest = current_texts[-1]
                if latest and latest == last_text and not has_stop_button(page):
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
