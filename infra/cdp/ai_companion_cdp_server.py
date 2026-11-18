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
import shutil
import subprocess
import sys
import threading
import time
import urllib.request
import winreg

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
        self._active_pages: dict[str, object] = {}
        self._connect_lock = threading.Lock()

    def _is_cdp_up(self) -> bool:
        try:
            opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
            with opener.open(f"{CDP_URL}/json/version", timeout=1.2) as resp:
                return resp.status == 200
        except Exception:
            return False

    @staticmethod
    def _find_chrome_exe() -> str:
        found = shutil.which("chrome.exe")
        if found:
            return found
        candidates = [
            os.path.join(os.environ.get("PROGRAMFILES", r"C:\Program Files"), "Google", "Chrome", "Application", "chrome.exe"),
            os.path.join(os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)"), "Google", "Chrome", "Application", "chrome.exe"),
            os.path.join(os.environ.get("LOCALAPPDATA", ""), "Google", "Chrome", "Application", "chrome.exe"),
        ]
        for path in candidates:
            if path and os.path.isfile(path):
                return path
        for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                with winreg.OpenKey(root, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe") as key:
                    val, _ = winreg.QueryValueEx(key, "")
                    if val and os.path.isfile(val):
                        return val
            except Exception:
                pass
        return "chrome.exe"

    def _launch_cdp_chrome(self, target_url: str, minimized: bool = True) -> bool:
        local_app = os.environ.get("LOCALAPPDATA", "")
        profile_dir = str(Path(local_app) / "GeminiCDP") if local_app else "GeminiCDP"

        if not self._is_cdp_up():
            try:
                subprocess.run(
                    ["powershell.exe", "-NoProfile", "-Command", 
                     "Get-CimInstance Win32_Process -Filter \"Name='chrome.exe'\" | Where-Object { $_.CommandLine -match 'GeminiCDP' } | Invoke-CimMethod -MethodName Terminate"],
                    shell=False,
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
                )
                time.sleep(0.5)
            except Exception:
                pass

        chrome_exe = self._find_chrome_exe()
        args = [
            chrome_exe,
            f"--user-data-dir={profile_dir}",
            "--remote-debugging-port=9222",
        ]
        if minimized:
            args.append("--start-minimized")
        args.append(target_url)
        try:
            creation_flags = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
            subprocess.Popen(args, shell=False, creationflags=creation_flags)
        except Exception:
            return False
        deadline = time.time() + 25.0
        while time.time() < deadline:
            if self._is_cdp_up():
                return True
            time.sleep(0.3)
        return False

    def _connect(self, launch_if_down: bool = False, target_url: str = CONSUMER_URL, minimized: bool = True):
        with self._connect_lock:
            if self._browser is not None:
                try:
                    if self._browser.is_connected() and self._browser.contexts:
                        return self._browser
                except Exception:
                    self._browser = None

            if not self._is_cdp_up():
                self._browser = None
                if not launch_if_down:
                    raise RuntimeError("CDP_NOT_RUNNING")
                if not self._launch_cdp_chrome(target_url, minimized=minimized):
                    raise RuntimeError("CDP_LAUNCH_FAILED")

            if self._pw is None:
                self._pw = sync_playwright().start()

            last_exc = None
            deadline = time.time() + 6.0
            while time.time() < deadline:
                try:
                    browser = self._pw.chromium.connect_over_cdp(CDP_URL)
                    if browser.contexts:
                        self._browser = browser
                        return self._browser
                except Exception as exc:
                    last_exc = exc
                time.sleep(0.25)
            if last_exc is not None:
                raise last_exc
            raise RuntimeError("NO_BROWSER_CONTEXT")

    @staticmethod
    def _companion_meta(companion: str) -> tuple[str, str]:
        comp = (companion or "gemini").strip().lower()
        if comp in ("enterprise", "geminienterprise"):
            return "vertexaisearch.cloud.google.com", ENTERPRISE_URL
        if comp in ("whatsapp", "wa"):
            return "web.whatsapp.com", "https://web.whatsapp.com"
        if comp in ("teams", "msteams"):
            return "teams.microsoft.com", "https://teams.microsoft.com"
        return "gemini.google.com", CONSUMER_URL

    @staticmethod
    def _get_hwnd_title(hwnd: int) -> str:
        if not hwnd:
            return ""
        user32 = ctypes.windll.user32
        length = user32.GetWindowTextLengthW(hwnd)
        if length <= 0:
            return ""
        buf = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, buf, length + 1)
        return buf.value

    @staticmethod
    def _get_mru_target_ids() -> list[str]:
        try:
            opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
            with opener.open(f"{CDP_URL}/json/list", timeout=1.0) as resp:
                items = json.loads(resp.read().decode("utf-8"))
            return [
                str(item["id"])
                for item in items
                if item.get("type") == "page" and item.get("id")
            ]
        except Exception:
            return []

    @staticmethod
    def _get_page_target_id(page) -> str:
        cdp = None
        try:
            cdp = page.context.new_cdp_session(page)
            info = cdp.send("Target.getTargetInfo")
            return str(info.get("targetInfo", {}).get("targetId", ""))
        except Exception:
            return ""
        finally:
            if cdp is not None:
                try:
                    cdp.detach()
                except Exception:
                    pass

    @staticmethod
    def _strip_chrome_title_suffix(title: str) -> str:
        t = (title or "").strip()
        for suffix in (" - Google Chrome", " - Chromium"):
            if t.endswith(suffix):
                return t[: -len(suffix)].strip()
        return t

    @staticmethod
    def _page_has_focus(page) -> bool:
        try:
            return bool(page.evaluate("() => document.hasFocus()"))
        except Exception:
            return False

    def _pick_current_page(self, comp: str, matching_pages: list, active_title: str = "") -> object:
        if len(matching_pages) == 1:
            self._active_pages[comp] = matching_pages[0]
            return matching_pages[0]

        # 1. Active Chrome window title always reflects the currently selected tab.
        win_title = self._strip_chrome_title_suffix(active_title) or self._strip_chrome_title_suffix(
            self._get_hwnd_title(self._find_cdp_chrome_hwnd())
        )
        if win_title:
            title_matches = []
            for p in matching_pages:
                try:
                    t = (p.title() or "").strip()
                    if t and (t == win_title or win_title.startswith(t)):
                        title_matches.append(p)
                except Exception:
                    pass
            if len(title_matches) == 1:
                self._active_pages[comp] = title_matches[0]
                return title_matches[0]
            if len(title_matches) > 1:
                for p in title_matches:
                    if self._page_has_focus(p):
                        self._active_pages[comp] = p
                        return p
                tracked = self._active_pages.get(comp)
                if tracked in title_matches:
                    return tracked
                self._active_pages[comp] = title_matches[-1]
                return title_matches[-1]

        # 2. If Chrome is focused on an AI tab whose title differs slightly, document.hasFocus() wins.
        for p in matching_pages:
            if self._page_has_focus(p):
                self._active_pages[comp] = p
                return p

        # 3. Last tracked active page for this companion.
        tracked = self._active_pages.get(comp)
        if tracked in matching_pages:
            return tracked

        # 4. Fallback: /json/list order, then newest tab.
        mru_ids = self._get_mru_target_ids()
        if mru_ids:
            rank = {tid: idx for idx, tid in enumerate(mru_ids)}
            best_page = None
            best_rank = len(mru_ids) + 1
            for p in matching_pages:
                tid = self._get_page_target_id(p)
                if tid in rank and rank[tid] < best_rank:
                    best_rank = rank[tid]
                    best_page = p
            if best_page is not None:
                self._active_pages[comp] = best_page
                return best_page

        self._active_pages[comp] = matching_pages[-1]
        return matching_pages[-1]

    def _find_or_open_page(self, companion: str, launch_if_down: bool = False, new_tab: bool = False, minimized: bool = True, active_title: str = ""):
        comp = (companion or "gemini").strip().lower()
        needle, target_url = self._companion_meta(companion)
        browser = self._connect(launch_if_down=launch_if_down, target_url=target_url, minimized=minimized)
        if not browser.contexts:
            raise RuntimeError("NO_BROWSER_CONTEXT")
        context = browser.contexts[0]

        if not new_tab:
            matching_pages = []
            fallback_page = None
            is_consumer = (needle == "gemini.google.com")
            for page in context.pages:
                u = (page.url or "").lower()
                if needle not in u:
                    continue
                if not is_consumer or "/app" in u or "/gem/" in u:
                    matching_pages.append(page)
                elif fallback_page is None:
                    fallback_page = page
            if matching_pages:
                return self._pick_current_page(comp, matching_pages, active_title=active_title)
            if fallback_page is not None:
                fallback_page.goto(target_url, wait_until="domcontentloaded", timeout=35000)
                self._active_pages[comp] = fallback_page
                return fallback_page

        page = context.new_page()
        page.goto(target_url, wait_until="domcontentloaded", timeout=35000)
        try:
            page.bring_to_front()
        except Exception:
            pass
        self._active_pages[comp] = page
        return page

    @staticmethod
    def _find_composer(page, timeout_s: float = 12.0):
        deadline = time.time() + timeout_s
        while time.time() < deadline:
            try:
                loc = page.locator(COMPOSER_SEL).locator("visible=true").first
                if loc.count() > 0:
                    return loc
            except Exception:
                pass
            for frame in page.frames:
                if frame == page.main_frame:
                    continue
                try:
                    loc = frame.locator(COMPOSER_SEL).locator("visible=true").first
                    if loc.count() > 0:
                        return loc
                except Exception:
                    pass
            time.sleep(0.1)
        return None

    @staticmethod
    def _is_auth_page(page) -> bool:
        try:
            url = (page.url or "").lower()
            if any(k in url for k in ("accounts.google.com", "login.microsoftonline.com", "login.live.com", "servicelogin", "/signin")):
                return True
            if "gemini.google.com" in url or "vertexaisearch.cloud.google.com" in url:
                return False
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
        chrome_fallback: list[int] = []

        @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
        def enum_cb(hwnd, _lparam):
            if not user32.IsWindowVisible(hwnd):
                return True
            cls_buf = ctypes.create_unicode_buffer(64)
            user32.GetClassNameW(hwnd, cls_buf, 64)
            if cls_buf.value != "Chrome_WidgetWin_1":
                return True
            length = user32.GetWindowTextLengthW(hwnd)
            if length <= 0:
                return True
            buf = ctypes.create_unicode_buffer(length + 1)
            user32.GetWindowTextW(hwnd, buf, length + 1)
            title = buf.value
            if title_needle and title_needle.lower() in title.lower():
                matches.insert(0, int(hwnd))
            elif "Gemini" in title or "AskBosch" in title or "Vertex" in title:
                matches.append(int(hwnd))
            elif "Chrome" in title:
                chrome_fallback.append(int(hwnd))
            return True

        user32.EnumWindows(enum_cb, 0)
        if matches:
            return matches[0]
        return chrome_fallback[0] if chrome_fallback else 0

    @staticmethod
    def _restore_window_via_cdp(page) -> None:
        cdp = None
        try:
            cdp = page.context.new_cdp_session(page)
            info = cdp.send("Browser.getWindowForTarget")
            win_id = info.get("windowId")
            bounds = info.get("bounds", {})
            if win_id is not None and bounds.get("windowState") == "minimized":
                cdp.send(
                    "Browser.setWindowBounds",
                    {"windowId": win_id, "bounds": {"windowState": "normal"}},
                )
        except Exception:
            pass
        finally:
            if cdp is not None:
                try:
                    cdp.detach()
                except Exception:
                    pass

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
                    const seenPres = new Set();
                    const blocks = root.querySelectorAll('pre code, code-block code, .code-block code');
                    for (const b of blocks) {
                        const t = (b.innerText || '').trim();
                        if (t) {
                            out.push(t);
                            const p = b.closest('pre');
                            if (p) seenPres.add(p);
                        }
                    }
                    const pres = root.querySelectorAll('code-block pre, .code-block pre, pre');
                    for (const p of pres) {
                        if (seenPres.has(p)) continue;
                        const t = (p.innerText || '').trim();
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
            self._restore_window_via_cdp(page)
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
            if companion.strip().lower() not in ("whatsapp", "wa", "teams", "msteams"):
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
    def _count_attachment_chips(page) -> int:
        try:
            return int(
                page.evaluate(
                    """() => {
                        function isVisible(el) {
                            const r = el.getBoundingClientRect();
                            if (r.width <= 0 || r.height <= 0) return false;
                            const cs = getComputedStyle(el);
                            return cs.display !== 'none' && cs.visibility !== 'hidden';
                        }
                        const chipSel =
                            'button[aria-label*="Remove" i], button[aria-label*="Remover" i], ' +
                            'button[aria-label*="Delete attachment" i], button[aria-label*="Remove file" i], ' +
                            'ucs-file-chip, ucs-attachment, .file-preview, .attachment-chip, ' +
                            '[data-test-id*="file-preview" i], [data-test-id*="attachment" i]';
                        let count = 0;
                        function scan(root) {
                            if (!root) return;
                            for (const el of root.querySelectorAll(chipSel)) {
                                if (el.closest('conversations-list, [role="navigation"], nav, aside')) continue;
                                if (isVisible(el)) count++;
                            }
                            for (const el of root.querySelectorAll('*')) {
                                if (el.shadowRoot) scan(el.shadowRoot);
                            }
                        }
                        scan(document);
                        return count;
                    }"""
                )
            )
        except Exception:
            return 0

    @staticmethod
    def _wait_for_upload_ready(page, before_chips: int = 0, timeout_s: float = 25.0) -> None:
        # Wait through the initial gap where the web app hasn't mounted the upload chip/spinner yet,
        # then wait until any upload spinners clear and the Send button is enabled.
        time.sleep(0.5)
        deadline = time.time() + timeout_s
        appear_deadline = time.time() + min(5.0, timeout_s)
        saw_activity = False
        stable_since = 0.0

        while time.time() < deadline:
            state = None
            try:
                state = page.evaluate(
                    """() => {
                        function isVisible(el) {
                            const r = el.getBoundingClientRect();
                            if (r.width <= 0 || r.height <= 0) return false;
                            const cs = getComputedStyle(el);
                            return cs.display !== 'none' && cs.visibility !== 'hidden';
                        }
                        const busySel =
                            'mat-progress-spinner, mat-progress-bar, mat-spinner, ' +
                            'md-circular-progress, md-linear-progress, ' +
                            '[role="progressbar"], [aria-busy="true"], .uploading, ' +
                            '[class*="spinner" i], [class*="progressbar" i]';
                        function hasBusy(root) {
                            if (!root) return false;
                            for (const el of root.querySelectorAll(busySel)) {
                                if (el.closest('conversations-list, [role="navigation"], nav, aside')) continue;
                                if (isVisible(el)) return true;
                            }
                            for (const el of root.querySelectorAll('*')) {
                                if (el.shadowRoot && hasBusy(el.shadowRoot)) return true;
                            }
                            return false;
                        }
                        const chipSel =
                            'button[aria-label*="Remove" i], button[aria-label*="Remover" i], ' +
                            'button[aria-label*="Delete attachment" i], button[aria-label*="Remove file" i], ' +
                            'ucs-file-chip, ucs-attachment, .file-preview, .attachment-chip, ' +
                            '[data-test-id*="file-preview" i], [data-test-id*="attachment" i]';
                        let chips = 0;
                        function scanChips(root) {
                            if (!root) return;
                            for (const el of root.querySelectorAll(chipSel)) {
                                if (el.closest('conversations-list, [role="navigation"], nav, aside')) continue;
                                if (isVisible(el)) chips++;
                            }
                            for (const el of root.querySelectorAll('*')) {
                                if (el.shadowRoot) scanChips(el.shadowRoot);
                            }
                        }
                        scanChips(document);

                        const sendSels = [
                            'button.send-button:not(.stop)',
                            'button[aria-label^="Send" i]',
                            'button[aria-label^="Enviar" i]',
                            'button[aria-label="Submit" i]',
                        ];
                        let sawSend = false;
                        let sawDisabledSend = false;
                        function scanSend(root) {
                            if (!root) return;
                            for (const sel of sendSels) {
                                for (const btn of root.querySelectorAll(sel)) {
                                    if (!isVisible(btn)) continue;
                                    sawSend = true;
                                    if (btn.disabled || btn.getAttribute('aria-disabled') === 'true')
                                        sawDisabledSend = true;
                                }
                            }
                            for (const el of root.querySelectorAll('*')) {
                                if (el.shadowRoot) scanSend(el.shadowRoot);
                            }
                        }
                        scanSend(document);
                        return {
                            busy: hasBusy(document),
                            chips: chips,
                            sawSend: sawSend,
                            sawDisabledSend: sawDisabledSend,
                        };
                    }"""
                )
            except Exception:
                return

            busy = bool(state.get("busy"))
            chips = int(state.get("chips") or 0)
            saw_send = bool(state.get("sawSend"))
            saw_disabled_send = bool(state.get("sawDisabledSend"))

            if busy or saw_disabled_send or chips > before_chips:
                saw_activity = True

            if not busy and saw_send and not saw_disabled_send:
                if chips > before_chips or time.time() >= appear_deadline:
                    if stable_since <= 0.0:
                        stable_since = time.time()
                    elif time.time() - stable_since >= 0.4:
                        return
                else:
                    stable_since = 0.0
            else:
                stable_since = 0.0

            time.sleep(0.25)

    def _atomic_click_send_if_idle(self, page, enterprise: bool, allow_empty_text: bool = False) -> str:
        """Click Send/Submit atomically in JS only if not already generating and composer/attachments are non-empty."""
        return str(
            page.evaluate(
                """(args) => {
                    const { enterprise, allowEmptyText } = args;
                    function isVisible(el) {
                        const r = el.getBoundingClientRect();
                        if (r.width <= 0 || r.height <= 0) return false;
                        const cs = getComputedStyle(el);
                        return cs.display !== 'none' && cs.visibility !== 'hidden';
                    }
                    const stopSel =
                        'button.send-button.stop, ' +
                        'button[aria-label^="Stop" i], ' +
                        'button[aria-label^="Interromper" i], ' +
                        'mat-icon[data-mat-icon-name="stop"], ' +
                        '[aria-label="Stop streaming" i], ' +
                        '[aria-label="Stop generating" i], ' +
                        '[aria-label="Interromper transmissão" i]';
                    function hasStop(root) {
                        if (!root) return false;
                        for (const el of root.querySelectorAll(stopSel)) {
                            if (el.closest('conversations-list, [role="navigation"], nav, aside')) continue;
                            if (isVisible(el)) return true;
                        }
                        for (const el of root.querySelectorAll('*')) {
                            if (el.shadowRoot && hasStop(el.shadowRoot)) return true;
                        }
                        return false;
                    }
                    if (hasStop(document)) return 'already_generating';

                    // If composer is already empty (and no attachments), Enter already submitted — never click Send/Stop.
                    let composerText = '';
                    if (!enterprise) {
                        const comp = document.querySelector('rich-textarea .ql-editor, div[contenteditable="true"][role="textbox"], textarea');
                        if (comp) composerText = (comp.innerText || comp.value || '').trim();
                    } else {
                        function findText(root) {
                            if (!root) return '';
                            const ed = root.querySelector('.ProseMirror, div[contenteditable="true"], textarea');
                            if (ed) return (ed.innerText || ed.value || '').trim();
                            for (const el of root.querySelectorAll('*')) {
                                if (el.shadowRoot) {
                                    const t = findText(el.shadowRoot);
                                    if (t) return t;
                                }
                            }
                            return '';
                        }
                        composerText = findText(document);
                    }
                    if (!composerText && !allowEmptyText) return 'composer_empty';

                    const sendSels = [
                        'button.send-button:not(.stop):not([disabled])',
                        'button[aria-label^="Send" i]:not([disabled])',
                        'button[aria-label^="Enviar" i]:not([disabled])',
                        'button[aria-label="Submit" i]:not([disabled])',
                    ];
                    function findAndClick(root) {
                        if (!root) return '';
                        for (const sel of sendSels) {
                            for (const btn of root.querySelectorAll(sel)) {
                                if (!isVisible(btn)) continue;
                                if (btn.disabled || btn.getAttribute('aria-disabled') === 'true') continue;
                                const cls = (btn.className || '').toString().toLowerCase().split(/\\s+/);
                                const lbl = (btn.getAttribute('aria-label') || '').trim().toLowerCase();
                                if (cls.includes('stop') || lbl.startsWith('stop') || lbl.startsWith('interromper')) {
                                    return 'stop_found';
                                }
                                if (btn.querySelector('mat-icon[data-mat-icon-name="stop"]')) {
                                    return 'stop_found';
                                }
                                btn.click();
                                return 'clicked';
                            }
                        }
                        for (const el of root.querySelectorAll('*')) {
                            if (el.shadowRoot) {
                                const res = findAndClick(el.shadowRoot);
                                if (res) return res;
                            }
                        }
                        return '';
                    }
                    return findAndClick(document) || 'not_found';
                }""",
                {"enterprise": enterprise, "allowEmptyText": allow_empty_text},
            )
        )

    @staticmethod
    def _insert_text_in_composer(page, composer, text: str, clear: bool = False) -> None:
        inserted = False
        try:
            inserted = bool(
                composer.evaluate(
                    """(el, payload) => {
                        const { text, clear } = payload;
                        el.focus();
                        const isInput = el.tagName === 'TEXTAREA' || el.tagName === 'INPUT';
                        if (clear) {
                            if (isInput) {
                                el.value = '';
                                el.dispatchEvent(new Event('input', { bubbles: true }));
                            } else {
                                const sel = window.getSelection();
                                const range = document.createRange();
                                range.selectNodeContents(el);
                                sel.removeAllRanges();
                                sel.addRange(range);
                                document.execCommand('delete', false, null);
                            }
                        }
                        if (!text) return true;
                        if (isInput) {
                            const start = el.selectionStart ?? el.value.length;
                            const end = el.selectionEnd ?? el.value.length;
                            el.setRangeText(text, start, end, 'end');
                            el.dispatchEvent(new Event('input', { bubbles: true }));
                            return (el.value || '').trim().length > 0;
                        }
                        const sel = window.getSelection();
                        const range = document.createRange();
                        range.selectNodeContents(el);
                        range.collapse(false);
                        sel.removeAllRanges();
                        sel.addRange(range);
                        document.execCommand('insertText', false, text);
                        return ((el.innerText || el.textContent || '').trim().length > 0);
                    }""",
                    {"text": text, "clear": clear},
                )
            )
        except Exception:
            inserted = False

        if not inserted and (text or clear):
            try:
                composer.click()
            except Exception:
                pass
            if clear:
                page.keyboard.press("Control+A")
                page.keyboard.press("Backspace")
                time.sleep(0.03)
            if text:
                page.keyboard.insert_text(text)

    def _submit_and_confirm_start(self, page, enterprise: bool, has_attachments: bool = False) -> tuple[bool, int]:
        reply_count_before = len(self._get_response_texts(page, enterprise))

        # Prefer atomic JS Send-button click first: works reliably even when Chrome is minimized/unfocused.
        click_res = "not_found"
        btn_deadline = time.time() + (4.0 if has_attachments else 0.45)
        while time.time() < btn_deadline:
            click_res = self._atomic_click_send_if_idle(page, enterprise, allow_empty_text=has_attachments)
            if click_res in ("already_generating", "stop_found", "composer_empty"):
                return True, reply_count_before
            if click_res == "clicked":
                deadline = time.time() + 2.5
                while time.time() < deadline:
                    if self._has_stop_button(page) or len(self._get_response_texts(page, enterprise)) > reply_count_before:
                        return True, reply_count_before
                    time.sleep(0.08)
                return True, reply_count_before
            time.sleep(0.04)

        # Fallback: keyboard Enter (e.g. if a UI variant uses a non-matching Send button).
        page.keyboard.press("Enter")
        deadline = time.time() + 1.5
        while time.time() < deadline:
            if self._has_stop_button(page) or len(self._get_response_texts(page, enterprise)) > reply_count_before:
                return True, reply_count_before
            time.sleep(0.1)
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

        before_chips = 0
        if files:
            before_chips = self._count_attachment_chips(page)
            if not self._attach_files(page, files):
                return {"ok": False, "error": "FILE_ATTACH_FAILED"}

        self._insert_text_in_composer(page, composer, text, clear=clear)

        if submit:
            if files:
                self._wait_for_upload_ready(page, before_chips=before_chips, timeout_s=25.0)
            started, reply_count_before = self._submit_and_confirm_start(page, enterprise, has_attachments=bool(files))
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
            try:
                composer.evaluate("el => el.focus()")
            except Exception:
                pass
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
        page = self._find_or_open_page(companion, launch_if_down=True, new_tab=False, minimized=True)
        items = self._get_code_texts(page) if kind == "code" else self._get_response_texts(page, enterprise)
        if not items:
            return {"ok": False, "error": "NO_RESPONSE_FOUND"}
        return {"ok": True, "text": items[-1]}

    def new_chat(self, data: dict) -> dict:
        companion = data.get("companion", "gemini")
        active_title = (data.get("active_title") or "").strip()
        page = self._find_or_open_page(companion, launch_if_down=True, new_tab=False, minimized=True, active_title=active_title)
        composer = self._find_composer(page, timeout_s=6.0)
        if composer is None and self._is_auth_page(page):
            return {"ok": False, "error": "AUTH_REQUIRED"}
        for sel in [
            'button:has-text("New chat")',
            'button[aria-label*="New chat" i]',
            'a[aria-label*="New chat" i]',
            'button[aria-label*="New conversation" i]',
            'button:has-text("New conversation")',
            'button[aria-label*="Novo chat" i]',
            'button:has-text("Novo chat")',
        ]:
            try:
                btn = page.locator(sel).locator("visible=true").first
                if btn.count() > 0:
                    try:
                        btn.evaluate("el => el.click()")
                    except Exception:
                        btn.click()
                    time.sleep(0.25)
                    return {"ok": True}
            except Exception:
                pass
        page.keyboard.press("Control+Shift+O")
        time.sleep(0.3)
        if self._find_composer(page, timeout_s=2.0) is not None:
            return {"ok": True}
        return {"ok": False, "error": "NEW_CHAT_NOT_FOUND"}

    def close_tab(self, data: dict) -> dict:
        companion = data.get("companion", "gemini")
        comp = (companion or "gemini").strip().lower()
        needle, _ = self._companion_meta(companion)
        browser = self._connect(launch_if_down=False)
        if not browser.contexts:
            return {"ok": False, "error": "NO_BROWSER_CONTEXT"}
        context = browser.contexts[0]
        is_consumer = (needle == "gemini.google.com")
        matching_pages = []
        for page in context.pages:
            u = (page.url or "").lower()
            if needle not in u:
                continue
            if not is_consumer or "/app" in u or "/gem/" in u:
                matching_pages.append(page)
        if len(matching_pages) <= 1:
            return {"ok": False, "error": "ONLY_ONE_TAB"}
        target = self._pick_current_page(comp, matching_pages)
        remaining = [p for p in matching_pages if p != target]
        try:
            target.close()
        except Exception as exc:
            return {"ok": False, "error": str(exc)}
        if remaining:
            self._active_pages[comp] = remaining[-1]
        else:
            self._active_pages.pop(comp, None)
        return {"ok": True, "remaining": len(remaining)}

    def select_model(self, data: dict) -> dict:
        companion = data.get("companion", "gemini")
        model_name = (data.get("model") or "").strip()
        if not model_name:
            return {"ok": False, "error": "EMPTY_MODEL"}
        page = self._find_or_open_page(companion, launch_if_down=True, new_tab=False, minimized=True)
        composer = self._find_composer(page, timeout_s=6.0)
        if composer is None and self._is_auth_page(page):
            return {"ok": False, "error": "AUTH_REQUIRED"}
        # Open model picker button
        opt_sel = (
            f'button:has-text("{model_name}"), '
            f'[role="menuitem"]:has-text("{model_name}"), '
            f'[role="menuitemradio"]:has-text("{model_name}"), '
            f'[role="option"]:has-text("{model_name}")'
        )
        for sel in [
            'button[aria-label*="Choose model" i]',
            'button[aria-label*="model" i]',
            'button[aria-label*="modelo" i]',
            '[data-test-id="bard-mode-menu-button"]',
            'button:has-text("Pro")',
            'button:has-text("Flash")',
            'button:has-text("Fast")',
            'button:has-text("Thinking")',
            'button:has-text("Auto")',
            'button:has-text("Hydrogen")',
            'button:has-text("Argon")',
            'button:has-text("Gemini")',
        ]:
            try:
                btn = page.locator(sel).locator("visible=true").first
                if btn.count() > 0:
                    try:
                        btn.evaluate("el => el.click()")
                    except Exception:
                        btn.click()
                    time.sleep(0.25)
                    opt = page.locator(f'text="{model_name}"').locator("visible=true").first
                    if opt.count() == 0:
                        opt = page.locator(opt_sel).locator("visible=true").first
                    if opt.count() > 0:
                        try:
                            opt.evaluate("el => el.click()")
                        except Exception:
                            opt.click()
                        time.sleep(0.2)
                        return {"ok": True, "model": model_name}
                    try:
                        page.keyboard.press("Escape")
                    except Exception:
                        pass
            except Exception:
                pass
        return {"ok": False, "error": f"MODEL_NOT_FOUND:{model_name}"}

    def read_aloud(self, data: dict) -> dict:
        companion = data.get("companion", "gemini")
        page = self._find_or_open_page(companion, launch_if_down=True, new_tab=False, minimized=True)
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

    @staticmethod
    def _find_first_visible(page, selectors: list[str], timeout_s: float = 6.0):
        deadline = time.time() + timeout_s
        while time.time() < deadline:
            for sel in selectors:
                try:
                    loc = page.locator(sel).locator("visible=true").first
                    if loc.count() > 0:
                        return loc
                except Exception:
                    pass
            time.sleep(0.12)
        return None

    def _chat_jump_whatsapp(self, page, contact: str, text: str, files: list[str], submit: bool) -> dict:
        search_sels = [
            'div[contenteditable="true"][data-tab="3"]',
            'div[role="textbox"][aria-label*="Search" i]',
            'div[role="textbox"][aria-label*="Pesquisar" i]',
            'div[role="textbox"][title*="Search" i]',
            '#side div[role="textbox"][contenteditable="true"]',
            'input[aria-label*="Search" i]',
        ]
        composer_sels = [
            'footer div[contenteditable="true"][data-tab="10"]',
            'footer div[role="textbox"][contenteditable="true"]',
            'footer div[contenteditable="true"]',
        ]

        composer = None
        last_err = "WHATSAPP_SEARCH_NOT_FOUND"
        for attempt in range(1, 4):
            search_loc = self._find_first_visible(page, search_sels, timeout_s=18.0 if attempt == 1 else 6.0)
            if search_loc is None:
                try:
                    if page.locator('canvas[aria-label*="Scan" i], div[data-ref]').count() > 0 or self._is_auth_page(page):
                        return {"ok": False, "error": "AUTH_REQUIRED"}
                except Exception:
                    pass
                last_err = "WHATSAPP_SEARCH_NOT_FOUND"
                time.sleep(0.5)
                continue

            self._insert_text_in_composer(page, search_loc, contact, clear=True)
            time.sleep(0.45)

            clicked = False
            click_deadline = time.time() + 5.0
            while time.time() < click_deadline:
                try:
                    clicked = bool(
                        page.evaluate(
                            """(contact) => {
                                const pane = document.querySelector('#pane-side');
                                if (!pane) return false;
                                const tokens = contact.toLowerCase().split(/\\s+/).filter(Boolean);
                                const rows = Array.from(pane.querySelectorAll('[role="listitem"], [role="row"]'));
                                let best = null;
                                for (const row of rows) {
                                    const titleEl = row.querySelector('span[title]');
                                    const label = ((titleEl && titleEl.getAttribute('title')) || row.innerText || '').toLowerCase();
                                    if (!label) continue;
                                    if (tokens.every(t => label.includes(t))) {
                                        best = titleEl || row.querySelector('[role="gridcell"]') || row;
                                        break;
                                    }
                                }
                                if (!best && rows.length > 0) {
                                    const firstTitle = rows[0].querySelector('span[title]');
                                    best = firstTitle || rows[0].querySelector('[role="gridcell"]') || rows[0];
                                }
                                if (!best) return false;
                                for (const evt of ['mousedown', 'mouseup', 'click']) {
                                    best.dispatchEvent(new MouseEvent(evt, { bubbles: true, cancelable: true, view: window }));
                                }
                                return true;
                            }""",
                            contact,
                        )
                    )
                except Exception:
                    clicked = False
                if clicked:
                    break
                time.sleep(0.15)

            if not clicked:
                last_err = f"CONTACT_NOT_FOUND:{contact}"
                try:
                    page.keyboard.press("Escape")
                except Exception:
                    pass
                time.sleep(0.3)
                continue

            composer = self._find_first_visible(page, composer_sels, timeout_s=2.5)
            if composer is None:
                try:
                    search_loc.press("Enter")
                except Exception:
                    pass
                composer = self._find_first_visible(page, composer_sels, timeout_s=3.5)
            if composer is not None:
                break
            last_err = "WHATSAPP_COMPOSER_NOT_FOUND"
            time.sleep(0.3)

        if composer is None:
            return {"ok": False, "error": last_err}

        if files:
            attached = False
            for _ in range(3):
                if self._attach_files(page, files):
                    attached = True
                    break
                time.sleep(0.3)
            if not attached:
                return {"ok": False, "error": "FILE_ATTACH_FAILED"}
            time.sleep(0.4)
        if text:
            for _ in range(3):
                self._insert_text_in_composer(page, composer, text, clear=False)
                time.sleep(0.1)
                try:
                    if (composer.inner_text() or "").strip():
                        break
                except Exception:
                    break

        if submit and (text or files):
            send_btn = self._find_first_visible(
                page,
                ['button[aria-label="Send" i]', 'button[aria-label="Enviar" i]', 'span[data-icon="send"]'],
                timeout_s=2.5,
            )
            if send_btn is not None:
                try:
                    send_btn.evaluate("el => el.click()")
                except Exception:
                    send_btn.click()
            else:
                page.keyboard.press("Enter")
        return {"ok": True}

    def _chat_jump_teams(self, page, contact: str, text: str, files: list[str], submit: bool) -> dict:
        search_sels = [
            'input[id*="search" i]',
            '[data-tid="search-box"] input',
            'input[placeholder*="Search" i]',
            'input[placeholder*="Pesquisar" i]',
        ]
        composer_sels = [
            '[data-tid="ckeditor"]',
            'div[role="textbox"][contenteditable="true"]',
        ]
        composer = None
        last_err = "TEAMS_SEARCH_NOT_FOUND"
        for attempt in range(1, 4):
            search_loc = self._find_first_visible(page, search_sels, timeout_s=15.0 if attempt == 1 else 6.0)
            if search_loc is None:
                if self._is_auth_page(page):
                    return {"ok": False, "error": "AUTH_REQUIRED"}
                last_err = "TEAMS_SEARCH_NOT_FOUND"
                time.sleep(0.5)
                continue

            self._insert_text_in_composer(page, search_loc, contact, clear=True)
            time.sleep(0.45)

            result_loc = self._find_first_visible(
                page,
                [
                    '[data-tid*="search" i] [role="option"]',
                    '[role="listbox"] [role="option"]',
                    '[data-tid*="autosuggest" i] [role="listitem"]',
                ],
                timeout_s=5.0,
            )
            if result_loc is not None:
                try:
                    result_loc.evaluate("el => el.click()")
                except Exception:
                    result_loc.click()
            else:
                page.keyboard.press("ArrowDown")
                page.keyboard.press("Enter")

            composer = self._find_first_visible(page, composer_sels, timeout_s=5.0)
            if composer is not None:
                break
            last_err = "TEAMS_COMPOSER_NOT_FOUND"
            try:
                page.keyboard.press("Escape")
            except Exception:
                pass
            time.sleep(0.3)

        if composer is None:
            return {"ok": False, "error": last_err}

        if files:
            attached = False
            for _ in range(3):
                if self._attach_files(page, files):
                    attached = True
                    break
                time.sleep(0.3)
            if not attached:
                return {"ok": False, "error": "FILE_ATTACH_FAILED"}
            time.sleep(0.4)
        if text:
            for _ in range(3):
                self._insert_text_in_composer(page, composer, text, clear=False)
                time.sleep(0.1)
                try:
                    if (composer.inner_text() or "").strip():
                        break
                except Exception:
                    break

        if submit and (text or files):
            send_btn = self._find_first_visible(
                page,
                ['button[data-tid="newMessageCommands-send"]', 'button[aria-label="Send" i]', 'button[aria-label="Enviar" i]'],
                timeout_s=2.5,
            )
            if send_btn is not None:
                try:
                    send_btn.evaluate("el => el.click()")
                except Exception:
                    send_btn.click()
            else:
                page.keyboard.press("Enter")
        return {"ok": True}

    def chat_jump(self, data: dict) -> dict:
        app = (data.get("app") or "").strip().lower()
        contact = (data.get("contact") or "").strip()
        text = data.get("text", "") or ""
        files = data.get("files") or []
        submit = bool(data.get("submit", False))
        activate = bool(data.get("activate", False))
        if app not in ("whatsapp", "teams"):
            return {"ok": False, "error": f"UNSUPPORTED_CHAT_APP:{app}"}
        if not contact:
            return {"ok": False, "error": "EMPTY_CONTACT"}

        page = self._find_or_open_page(app, launch_if_down=True, new_tab=False, minimized=not activate)
        res = (
            self._chat_jump_whatsapp(page, contact, text, files, submit)
            if app == "whatsapp"
            else self._chat_jump_teams(page, contact, text, files, submit)
        )
        if not res.get("ok"):
            return res

        hwnd = 0
        if activate:
            self._restore_window_via_cdp(page)
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
        return {"ok": True, "hwnd": hwnd}


def _file_sig() -> tuple[int, int]:
    try:
        st = os.stat(__file__)
        return (int(st.st_size), int(st.st_mtime))
    except Exception:
        return (0, 0)


_STARTUP_FILE_SIG = _file_sig()
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
            self._send_json(200, {
                "ok": True,
                "stale": _file_sig() != _STARTUP_FILE_SIG,
                "cdp_up": BRIDGE._is_cdp_up(),
            })
            return
        self._send_json(404, {"ok": False, "error": "NOT_FOUND"})

    def do_POST(self):
        if self.path == "/shutdown":
            self._send_json(200, {"ok": True})
            os._exit(0)
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
            "/close_tab": BRIDGE.close_tab,
            "/select_model": BRIDGE.select_model,
            "/read_aloud": BRIDGE.read_aloud,
            "/chat/jump": BRIDGE.chat_jump,
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


def _is_healthy_bridge_up() -> bool:
    try:
        opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
        with opener.open(f"http://{HOST}:{PORT}/health", timeout=0.8) as resp:
            if resp.status != 200:
                return False
            data = json.loads(resp.read().decode("utf-8"))
            return bool(data.get("ok")) and not bool(data.get("stale"))
    except Exception:
        return False


def main() -> int:
    os.environ["NO_PROXY"] = "127.0.0.1,localhost"
    os.environ["no_proxy"] = "127.0.0.1,localhost"
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    server = None
    last_exc = None
    deadline = time.time() + 6.0
    while time.time() < deadline:
        try:
            server = HTTPServer((HOST, PORT), Handler)
            break
        except OSError as exc:
            last_exc = exc
            if _is_healthy_bridge_up():
                return 0
            time.sleep(0.2)
    if server is None:
        try:
            log_path = os.path.join(os.environ.get("TEMP", "."), "ai_companion_cdp_server.log")
            with open(log_path, "a", encoding="utf-8") as fh:
                fh.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] bind failed: {last_exc}\n")
        except Exception:
            pass
        raise last_exc if last_exc is not None else RuntimeError("HTTPServer bind failed")
    server.serve_forever()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
