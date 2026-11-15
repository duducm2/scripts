"""
Persistent localhost CDP bridge for AI Companions (Consumer Gemini + Gemini Enterprise).
Runs on 127.0.0.1:8768, keeps a warm Playwright CDP connection to Chrome on 127.0.0.1:9222,
and executes all companion actions in the background without stealing OS focus, mouse, or clipboard.
"""

from __future__ import annotations

import ctypes
from http.server import BaseHTTPRequestHandler, HTTPServer
import json
import os
from pathlib import Path
import subprocess
import sys
import time
import urllib.request

from playwright.sync_api import sync_playwright

HOST = "127.0.0.1"
PORT = 8768
CDP_URL = "http://127.0.0.1:9222"
CONSUMER_URL = "https://gemini.google.com/app"
ENTERPRISE_URL = "https://vertexaisearch.cloud.google.com/u/1/eu/home/cid/bcb383f1-26d8-41fd-9a55-623f7e93de92?pli=1"

COMPOSER_SEL = (
    "#agent-search-prosemirror-editor, "
    '.ProseMirror[contenteditable="true"], '
    'div.ql-editor[contenteditable="true"], '
    'rich-textarea [contenteditable="true"], '
    '[role="textbox"][contenteditable="true"], '
    "textarea[aria-label], "
    "textarea"
)


class CDPBridge:
    def __init__(self) -> None:
        self._pw = None
        self._browser = None

    def _is_cdp_up(self) -> bool:
        try:
            opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
            with opener.open(f"{CDP_URL}/json/version", timeout=1.2) as resp:
                return resp.status == 200
        except Exception:
            return False

    def _launch_cdp_chrome(self, target_url: str, minimized: bool = True) -> bool:
        local_app = os.environ.get("LOCALAPPDATA", "")
        profile_dir = str(Path(local_app) / "GeminiCDP") if local_app else "GeminiCDP"
        args = [
            "chrome.exe",
            f"--user-data-dir={profile_dir}",
            "--remote-debugging-port=9222",
        ]
        if minimized:
            args.append("--start-minimized")
        args.append(target_url)
        try:
            subprocess.Popen(args, shell=True)
        except Exception:
            return False
        deadline = time.time() + 12.0
        while time.time() < deadline:
            if self._is_cdp_up():
                return True
            time.sleep(0.3)
        return False

    def _connect(self, launch_if_down: bool = False, target_url: str = CONSUMER_URL, minimized: bool = True):
        if not self._is_cdp_up():
            self._browser = None
            if not launch_if_down:
                raise RuntimeError("CDP_NOT_RUNNING")
            if not self._launch_cdp_chrome(target_url, minimized=minimized):
                raise RuntimeError("CDP_LAUNCH_FAILED")

        if self._pw is None:
            self._pw = sync_playwright().start()

        if self._browser is not None:
            try:
                if self._browser.is_connected() and self._browser.contexts:
                    return self._browser
            except Exception:
                self._browser = None

        self._browser = self._pw.chromium.connect_over_cdp(CDP_URL)
        return self._browser

    @staticmethod
    def _companion_meta(companion: str) -> tuple[str, str]:
        comp = (companion or "gemini").strip().lower()
        if comp in ("enterprise", "geminienterprise"):
            return "vertexaisearch.cloud.google.com", ENTERPRISE_URL
        return "gemini.google.com", CONSUMER_URL

    def _find_or_open_page(self, companion: str, launch_if_down: bool = False, new_tab: bool = False, minimized: bool = True):
        needle, target_url = self._companion_meta(companion)
        browser = self._connect(launch_if_down=launch_if_down, target_url=target_url, minimized=minimized)
        if not browser.contexts:
            raise RuntimeError("NO_BROWSER_CONTEXT")
        context = browser.contexts[0]

        if not new_tab:
            fallback_page = None
            is_consumer = (needle == "gemini.google.com")
            for page in context.pages:
                u = (page.url or "").lower()
                if needle not in u:
                    continue
                if not is_consumer or "/app" in u or "/gem/" in u:
                    return page
                if fallback_page is None:
                    fallback_page = page
            if fallback_page is not None:
                fallback_page.goto(target_url, wait_until="domcontentloaded", timeout=25000)
                return fallback_page

        page = context.new_page()
        page.goto(target_url, wait_until="domcontentloaded", timeout=25000)
        return page

    @staticmethod
    def _find_composer(page, timeout_s: float = 12.0):
        deadline = time.time() + timeout_s
        while time.time() < deadline:
            for target in [page, *page.frames]:
                try:
                    loc = target.locator(COMPOSER_SEL).locator("visible=true").first
                    if loc.count() > 0:
                        return loc
                except Exception:
                    pass
            time.sleep(0.25)
        return None

    @staticmethod
    def _is_auth_page(page) -> bool:
        try:
            url = (page.url or "").lower()
            if any(k in url for k in ("accounts.google.com", "login.microsoftonline.com", "servicelogin", "/signin")):
                return True
            title = (page.title() or "").lower()
            if any(k in title for k in ("sign in", "fazer login", "entrar")):
                return True
        except Exception:
            pass
        return False

    @staticmethod
    def _find_cdp_chrome_hwnd(title_needle: str = "") -> int:
        user32 = ctypes.windll.user32
        matches: list[int] = []

        @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
        def enum_cb(hwnd, _lparam):
            if not user32.IsWindowVisible(hwnd):
                return True
            length = user32.GetWindowTextLengthW(hwnd)
            if length <= 0:
                return True
            buf = ctypes.create_unicode_buffer(length + 1)
            user32.GetWindowTextW(hwnd, buf, length + 1)
            title = buf.value
            if "Chrome" not in title and "Gemini" not in title and "AskBosch" not in title and "Vertex" not in title:
                return True
            if title_needle and title_needle.lower() in title.lower():
                matches.insert(0, int(hwnd))
            elif "Gemini" in title or "AskBosch" in title or "Vertex AI Search" in title:
                matches.append(int(hwnd))
            return True

        user32.EnumWindows(enum_cb, 0)
        return matches[0] if matches else 0

    @staticmethod
    def _activate_hwnd(hwnd: int) -> None:
        if not hwnd:
            return
        user32 = ctypes.windll.user32
        SW_RESTORE = 9
        if user32.IsIconic(hwnd):
            user32.ShowWindow(hwnd, SW_RESTORE)
        user32.SetForegroundWindow(hwnd)

    @staticmethod
    def _has_stop_button(page) -> bool:
        return bool(
            page.evaluate(
                """() => {
                    const sel =
                        'button.send-button.stop, ' +
                        'button[aria-label^="Stop" i], ' +
                        'button[aria-label^="Interromper" i], ' +
                        'mat-icon[data-mat-icon-name="stop"], ' +
                        '[aria-label="Stop streaming" i], ' +
                        '[aria-label="Stop generating" i], ' +
                        '[aria-label="Interromper transmissão" i]';
                    function isVisible(el) {
                        const r = el.getBoundingClientRect();
                        if (r.width <= 0 || r.height <= 0) return false;
                        const cs = getComputedStyle(el);
                        return cs.display !== 'none' && cs.visibility !== 'hidden';
                    }
                    function check(root) {
                        if (!root) return false;
                        for (const el of root.querySelectorAll(sel)) {
                            if (el.closest('conversations-list, [role="navigation"], nav, aside')) continue;
                            if (isVisible(el)) return true;
                        }
                        for (const el of root.querySelectorAll('*')) {
                            if (el.shadowRoot && check(el.shadowRoot)) return true;
                        }
                        return false;
                    }
                    return check(document);
                }"""
            )
        )

    @staticmethod
    def _get_response_texts(page, enterprise: bool) -> list[str]:
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
        return page.evaluate(
            """() => {
                function collect(root, out) {
                    if (!root) return;
                    const els = root.querySelectorAll('ucs-markdown, ucs-response-markdown, .markdown, .markdown-content, [class*="markdown" i], message-content');
                    for (const el of els) {
                        if (el.closest('button, [role="button"], .ProseMirror')) continue;
                        const t = (el.innerText || '').trim();
                        if (!t) continue;
                        const low = t.toLowerCase();
                        if (low === 'thumb up' || low === 'thumb down' || low === 'thumb_up' || low === 'thumb_down') continue;
                        out.push(t);
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

    @staticmethod
    def _get_code_texts(page) -> list[str]:
        return page.evaluate(
            """() => {
                function collect(root, out) {
                    if (!root) return;
                    const blocks = root.querySelectorAll('pre code, code-block code, .code-block code');
                    for (const b of blocks) {
                        const t = (b.innerText || '').trim();
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

    def ensure(self, data: dict) -> dict:
        companion = data.get("companion", "gemini")
        activate = bool(data.get("activate", False))
        new_tab = bool(data.get("new_tab", False))
        page = self._find_or_open_page(companion, launch_if_down=True, new_tab=new_tab, minimized=not activate)
        if activate:
            try:
                page.bring_to_front()
            except Exception:
                pass
            page_title = ""
            try:
                page_title = page.title() or ""
            except Exception:
                pass
            hwnd = self._find_cdp_chrome_hwnd(page_title)
            self._activate_hwnd(hwnd)
            composer = self._find_composer(page, timeout_s=6.0)
            if composer:
                try:
                    composer.click()
                except Exception:
                    pass
        else:
            page_title = ""
            try:
                page_title = page.title() or ""
            except Exception:
                pass
            hwnd = self._find_cdp_chrome_hwnd(page_title)
        return {"ok": True, "url": page.url, "title": page.title(), "hwnd": hwnd}

    def _toggle_image_gen(self, page, enterprise: bool) -> bool:
        selectors = [
            'button:has-text("Tools")',
            'button:has-text("Select tools")',
            'button[aria-label*="Tools" i]',
        ]
        for sel in selectors:
            try:
                btn = page.locator(sel).locator("visible=true").first
                if btn.count() > 0:
                    btn.click()
                    time.sleep(0.25)
                    for item_sel in ['text="Create image"', 'text="Create images"', 'button:has-text("Create image")']:
                        item = page.locator(item_sel).locator("visible=true").first
                        if item.count() > 0:
                            item.click()
                            return True
            except Exception:
                pass
        return False

    def _attach_files(self, page, files: list[str]) -> bool:
        valid = [f for f in files if f and os.path.isfile(f)]
        if not valid:
            return True
        # 1. Try direct hidden input[type=file] or hidden-local-file-upload-button
        try:
            inp = page.locator('input[type="file"]').first
            if inp.count() > 0:
                inp.set_input_files(valid)
                time.sleep(0.4)
                return True
        except Exception:
            pass
        # 2. Try 2-step "Upload & tools" (+) menu -> "Upload files" (Consumer Gemini)
        for trigger_sel in ['button[aria-label*="Upload & tools" i]', 'button[aria-label*="Open upload" i]', 'button[aria-label*="Add file" i]']:
            try:
                trig = page.locator(trigger_sel).locator("visible=true").first
                if trig.count() > 0:
                    trig.click()
                    time.sleep(0.25)
                    for item_sel in ['button[aria-label*="Upload files" i]', 'button:has-text("Upload files")', '[role="menuitem"]:has-text("Upload files")']:
                        item = page.locator(item_sel).locator("visible=true").first
                        if item.count() > 0:
                            with page.expect_file_chooser(timeout=3000) as fc_info:
                                item.click()
                            fc_info.value.set_files(valid)
                            time.sleep(0.5)
                            return True
            except Exception:
                pass
        # 3. Direct upload button fallback (Enterprise / older UI)
        for sel in ['button.hidden-local-file-upload-button', 'button[aria-label*="Upload" i]', 'button[aria-label*="Add file" i]']:
            try:
                btn = page.locator(sel).first
                if btn.count() > 0:
                    with page.expect_file_chooser(timeout=3000) as fc_info:
                        btn.click(force=True)
                    fc_info.value.set_files(valid)
                    time.sleep(0.5)
                    return True
            except Exception:
                pass
        return False

    @staticmethod
    def _wait_for_upload_ready(page, timeout_s: float = 8.0) -> None:
        deadline = time.time() + timeout_s
        while time.time() < deadline:
            busy = False
            try:
                busy = bool(
                    page.evaluate(
                        """() => {
                            const spinners = document.querySelectorAll('input-container mat-progress-spinner, input-container [role="progressbar"], .input-area mat-progress-spinner, .input-area [role="progressbar"]');
                            for (const s of spinners) {
                                const r = s.getBoundingClientRect();
                                if (r.width > 0 && r.height > 0) return true;
                            }
                            return false;
                        }"""
                    )
                )
            except Exception:
                busy = False
            if not busy:
                return
            time.sleep(0.25)

    def _submit_and_confirm_start(self, page, enterprise: bool) -> tuple[bool, int]:
        reply_count_before = len(self._get_response_texts(page, enterprise))
        page.keyboard.press("Enter")
        deadline = time.time() + 1.2
        while time.time() < deadline:
            if self._has_stop_button(page):
                return True, reply_count_before
            time.sleep(0.15)
        # Fallback: click an enabled visible Send/Submit button if Enter did not trigger generation
        for sel in ['button.send-button:not([disabled])', 'button[aria-label^="Send" i]:not([disabled])', 'button[aria-label^="Enviar" i]:not([disabled])', 'button[aria-label="Submit" i]:not([disabled])']:
            try:
                btn = page.locator(sel).locator("visible=true").first
                if btn.count() > 0:
                    btn.click()
                    break
            except Exception:
                pass
        deadline = time.time() + 2.8
        while time.time() < deadline:
            if self._has_stop_button(page):
                return True, reply_count_before
            time.sleep(0.15)
        return False, reply_count_before

    def paste(self, data: dict) -> dict:
        companion = data.get("companion", "gemini")
        enterprise = companion.strip().lower() in ("enterprise", "geminienterprise")
        text = data.get("text", "") or ""
        files = data.get("files") or []
        submit = bool(data.get("submit", False))
        clear = bool(data.get("clear", False))
        image_gen = bool(data.get("image_gen", False))

        page = self._find_or_open_page(companion, launch_if_down=True, new_tab=False, minimized=True)
        composer = self._find_composer(page, timeout_s=12.0)
        if composer is None:
            if self._is_auth_page(page):
                return {"ok": False, "error": "AUTH_REQUIRED"}
            return {"ok": False, "error": f"COMPOSER_NOT_FOUND on {page.url}"}

        if image_gen:
            self._toggle_image_gen(page, enterprise)
            time.sleep(0.2)

        if files:
            if not self._attach_files(page, files):
                return {"ok": False, "error": "FILE_ATTACH_FAILED"}

        composer.click()
        if clear:
            page.keyboard.press("Control+A")
            page.keyboard.press("Backspace")
            time.sleep(0.05)

        if text:
            page.keyboard.insert_text(text)
            time.sleep(0.12)

        if submit:
            if files:
                self._wait_for_upload_ready(page, timeout_s=8.0)
            started, reply_count_before = self._submit_and_confirm_start(page, enterprise)
            return {
                "ok": True,
                "submitted": True,
                "started": started,
                "reply_count_before": reply_count_before,
                "hwnd": self._find_cdp_chrome_hwnd(),
            }

        return {"ok": True, "submitted": False, "hwnd": self._find_cdp_chrome_hwnd()}

    def submit(self, data: dict) -> dict:
        companion = data.get("companion", "gemini")
        enterprise = companion.strip().lower() in ("enterprise", "geminienterprise")
        page = self._find_or_open_page(companion, launch_if_down=False)
        composer = self._find_composer(page, timeout_s=5.0)
        if composer is None and self._is_auth_page(page):
            return {"ok": False, "error": "AUTH_REQUIRED"}
        if composer is not None:
            composer.click()
        started, reply_count_before = self._submit_and_confirm_start(page, enterprise)
        return {"ok": True, "started": started, "reply_count_before": reply_count_before}

    def status(self, data: dict) -> dict:
        companion = data.get("companion", "gemini")
        enterprise = companion.strip().lower() in ("enterprise", "geminienterprise")
        page = self._find_or_open_page(companion, launch_if_down=False)
        generating = self._has_stop_button(page)
        replies = self._get_response_texts(page, enterprise)
        composer_text = ""
        try:
            composer = self._find_composer(page, timeout_s=1.5)
            if composer is not None:
                composer_text = (composer.inner_text() or "").strip()
        except Exception:
            pass
        return {
            "ok": True,
            "generating": generating,
            "reply_count": len(replies),
            "last_reply_len": len(replies[-1]) if replies else 0,
            "composer_text": composer_text,
            "title": page.title(),
        }

    def copy_last(self, data: dict) -> dict:
        companion = data.get("companion", "gemini")
        enterprise = companion.strip().lower() in ("enterprise", "geminienterprise")
        kind = (data.get("kind") or "message").strip().lower()
        page = self._find_or_open_page(companion, launch_if_down=False)
        items = self._get_code_texts(page) if kind == "code" else self._get_response_texts(page, enterprise)
        if not items:
            return {"ok": False, "error": "NO_RESPONSE_FOUND"}
        return {"ok": True, "text": items[-1]}

    def new_chat(self, data: dict) -> dict:
        companion = data.get("companion", "gemini")
        page = self._find_or_open_page(companion, launch_if_down=True, new_tab=False, minimized=True)
        if self._is_auth_page(page):
            return {"ok": False, "error": "AUTH_REQUIRED"}
        for sel in ['button:has-text("New chat")', 'button[aria-label*="New chat" i]', 'a[aria-label*="New chat" i]']:
            try:
                btn = page.locator(sel).locator("visible=true").first
                if btn.count() > 0:
                    btn.click()
                    time.sleep(0.3)
                    return {"ok": True}
            except Exception:
                pass
        page.keyboard.press("Control+Shift+O")
        time.sleep(0.3)
        if self._find_composer(page, timeout_s=2.0) is not None:
            return {"ok": True}
        return {"ok": False, "error": "NEW_CHAT_NOT_FOUND"}

    def select_model(self, data: dict) -> dict:
        companion = data.get("companion", "gemini")
        model_name = (data.get("model") or "").strip()
        if not model_name:
            return {"ok": False, "error": "EMPTY_MODEL"}
        page = self._find_or_open_page(companion, launch_if_down=True, new_tab=False, minimized=True)
        if self._is_auth_page(page):
            return {"ok": False, "error": "AUTH_REQUIRED"}
        # Open model picker button
        for sel in [
            'button[aria-label*="Choose model" i]',
            'button[aria-label*="model" i]',
            '[data-test-id="bard-mode-menu-button"]',
            'button:has-text("Pro")',
            'button:has-text("Flash")',
            'button:has-text("Hydrogen")',
            'button:has-text("Argon")',
        ]:
            try:
                btn = page.locator(sel).locator("visible=true").first
                if btn.count() > 0:
                    btn.click()
                    time.sleep(0.25)
                    opt = page.locator(f'text="{model_name}"').locator("visible=true").first
                    if opt.count() == 0:
                        opt = page.locator(f'button:has-text("{model_name}"), [role="menuitem"]:has-text("{model_name}"), [role="option"]:has-text("{model_name}")').locator("visible=true").first
                    if opt.count() > 0:
                        opt.click()
                        time.sleep(0.2)
                        return {"ok": True, "model": model_name}
            except Exception:
                pass
        return {"ok": False, "error": f"MODEL_NOT_FOUND:{model_name}"}

    def read_aloud(self, data: dict) -> dict:
        companion = data.get("companion", "gemini")
        page = self._find_or_open_page(companion, launch_if_down=False)
        for sel in ['button[aria-label*="Listen" i]', 'button[aria-label*="Read aloud" i]']:
            try:
                btns = page.locator(sel).locator("visible=true")
                cnt = btns.count()
                if cnt > 0:
                    btns.nth(cnt - 1).click()
                    return {"ok": True}
            except Exception:
                pass
        return {"ok": False, "error": "READ_ALOUD_NOT_FOUND"}


BRIDGE = CDPBridge()


class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        return

    def _send_json(self, code: int, payload: dict) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path == "/health":
            self._send_json(200, {"ok": True})
            return
        self._send_json(404, {"ok": False, "error": "NOT_FOUND"})

    def do_POST(self):
        length = int(self.headers.get("Content-Length", "0") or "0")
        raw = self.rfile.read(length) if length > 0 else b"{}"
        try:
            data = json.loads(raw.decode("utf-8")) if raw else {}
        except Exception:
            self._send_json(400, {"ok": False, "error": "INVALID_JSON"})
            return

        routes = {
            "/ensure": BRIDGE.ensure,
            "/paste": BRIDGE.paste,
            "/submit": BRIDGE.submit,
            "/status": BRIDGE.status,
            "/copy_last": BRIDGE.copy_last,
            "/new_chat": BRIDGE.new_chat,
            "/select_model": BRIDGE.select_model,
            "/read_aloud": BRIDGE.read_aloud,
        }
        fn = routes.get(self.path)
        if fn is None:
            self._send_json(404, {"ok": False, "error": "NOT_FOUND"})
            return
        try:
            res = fn(data)
            self._send_json(200, res)
        except Exception as exc:
            self._send_json(500, {"ok": False, "error": str(exc)})


def main() -> int:
    os.environ["NO_PROXY"] = "127.0.0.1,localhost"
    os.environ["no_proxy"] = "127.0.0.1,localhost"
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    server = HTTPServer((HOST, PORT), Handler)
    server.serve_forever()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
