import json
import os
import socket
import sys
import threading
import time
import urllib.request
from pathlib import Path
from unittest import mock

CDP_DIR = Path(__file__).resolve().parents[2] / "cdp"
if str(CDP_DIR) not in sys.path:
    sys.path.insert(0, str(CDP_DIR))

import ai_companion_cdp_server as srv


def test_query_cdp_port_pid_finds_local_listener():
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.bind(("127.0.0.1", 0))
    sock.listen(1)
    try:
        port = sock.getsockname()[1]
        pid = srv.CDPBridge._query_cdp_port_pid(port)
        assert pid == os.getpid()
    finally:
        sock.close()


def test_connect_reuses_live_cdp_even_when_hwnd_is_zero():
    bridge = srv.CDPBridge()
    fake_page = mock.MagicMock()
    fake_ctx = mock.MagicMock()
    fake_ctx.pages = [fake_page]
    fake_browser = mock.MagicMock()
    fake_browser.is_connected.return_value = True
    fake_browser.contexts = [fake_ctx]
    bridge._browser = fake_browser

    with (
        mock.patch.object(bridge, "_is_cdp_up", return_value=True),
        mock.patch.object(bridge, "_find_cdp_chrome_hwnd", return_value=0),
        mock.patch.object(bridge, "_launch_cdp_chrome") as launch_mock,
    ):
        got = bridge._connect(launch_if_down=True, target_url="https://web.whatsapp.com", minimized=True)
        assert got is fake_browser
        launch_mock.assert_not_called()


def test_launch_cdp_chrome_never_kills_live_cdp_without_force():
    bridge = srv.CDPBridge()
    with (
        mock.patch.object(bridge, "_is_cdp_up", return_value=True),
        mock.patch.object(srv.subprocess, "run") as run_mock,
        mock.patch.object(srv.subprocess, "Popen") as popen_mock,
    ):
        assert bridge._launch_cdp_chrome("https://web.whatsapp.com", minimized=True, force_restart=False) is True
        run_mock.assert_not_called()
        popen_mock.assert_not_called()


def test_health_does_not_block_while_worker_is_busy():
    started = threading.Event()
    release = threading.Event()

    def slow_ensure(_data):
        started.set()
        release.wait(timeout=3.0)
        return {"ok": True, "url": "https://web.whatsapp.com", "title": "WhatsApp", "hwnd": 0}

    server = srv.SingleInstanceHTTPServer(("127.0.0.1", 0), srv.Handler)
    port = server.server_address[1]
    t_srv = threading.Thread(target=server.serve_forever, daemon=True)
    t_srv.start()

    try:
        with (
            mock.patch.object(srv.BRIDGE, "ensure", side_effect=slow_ensure),
            mock.patch.object(srv.BRIDGE, "_is_cdp_up", return_value=True),
        ):
            post_res = {}

            def do_post():
                req = urllib.request.Request(
                    f"http://127.0.0.1:{port}/ensure",
                    data=b'{"companion":"whatsapp"}',
                    headers={"Content-Type": "application/json"},
                    method="POST",
                )
                opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
                with opener.open(req, timeout=5.0) as resp:
                    post_res["status"] = resp.status
                    post_res["body"] = json.loads(resp.read().decode("utf-8"))

            t_post = threading.Thread(target=do_post, daemon=True)
            t_post.start()
            assert started.wait(timeout=2.0)

            t0 = time.monotonic()
            opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
            with opener.open(f"http://127.0.0.1:{port}/health", timeout=1.0) as resp:
                elapsed = time.monotonic() - t0
                health = json.loads(resp.read().decode("utf-8"))
            assert resp.status == 200
            assert elapsed < 0.5
            assert health["ok"] is True
            assert health["cdp_up"] is True
            assert health["busy"] is True

            release.set()
            t_post.join(timeout=3.0)
            assert post_res["status"] == 200
            assert post_res["body"]["ok"] is True
    finally:
        server.shutdown()
        server.server_close()


def test_find_or_open_page_uses_pinned_target_id_first():
    bridge = srv.CDPBridge()
    page1 = mock.MagicMock()
    page1.is_closed.return_value = False
    page1.url = "https://gemini.google.com/app/chat1"
    page2 = mock.MagicMock()
    page2.is_closed.return_value = False
    page2.url = "https://gemini.google.com/app/chat2"

    fake_ctx = mock.MagicMock()
    fake_ctx.pages = [page1, page2]
    fake_browser = mock.MagicMock()
    fake_browser.contexts = [fake_ctx]

    def fake_tid(p):
        return "TID_1" if p is page1 else "TID_2"

    with (
        mock.patch.object(bridge, "_connect", return_value=fake_browser),
        mock.patch.object(bridge, "_get_page_target_id", side_effect=fake_tid),
        mock.patch.object(bridge, "_pick_current_page") as pick_mock,
    ):
        got = bridge._find_or_open_page("gemini", target_id="TID_2")
        assert got is page2
        assert bridge._active_pages["gemini"] is page2
        pick_mock.assert_not_called()


def test_pick_current_page_minimized_prefers_active_pages_over_json_list_order():
    bridge = srv.CDPBridge()
    page1 = mock.MagicMock()
    page2 = mock.MagicMock()
    bridge._active_pages["enterprise"] = page2

    with (
        mock.patch.object(bridge, "_order_pages_by_cdp_mru", return_value=[page1, page2]),
        mock.patch.object(bridge, "_ensure_tab_tracker"),
        mock.patch.object(bridge, "_is_cdp_chrome_foreground", return_value=False),
        mock.patch.object(bridge, "_find_cdp_chrome_hwnd", return_value=0),
        mock.patch.object(bridge, "_page_is_visible", return_value=False),
        mock.patch.object(bridge, "_get_page_last_active_at", return_value=0),
    ):
        got = bridge._pick_current_page("enterprise", [page1, page2])
        assert got is page2


def test_select_model_returns_target_id():
    bridge = srv.CDPBridge()
    page2 = mock.MagicMock()
    page2.url = "https://vertexaisearch.cloud.google.com/u/1/eu/home/cid/123"
    opt_loc = mock.MagicMock()
    opt_loc.count.return_value = 1
    btn_loc = mock.MagicMock()
    btn_loc.count.return_value = 1

    def fake_locator(sel):
        m = mock.MagicMock()
        if "Hydrogen" in sel or "menuitem" in sel or "text=" in sel:
            m.locator.return_value.first = opt_loc
        else:
            m.locator.return_value.first = btn_loc
        return m

    page2.locator.side_effect = fake_locator

    with (
        mock.patch.object(bridge, "_find_or_open_page", return_value=page2),
        mock.patch.object(bridge, "_get_page_target_id", return_value="TID_2"),
        mock.patch.object(bridge, "_find_composer", return_value=mock.MagicMock()),
    ):
        res = bridge.select_model({"companion": "enterprise", "model": "Gemini 4 Hydrogen", "target_id": "TID_2"})
        assert res["ok"] is True
        assert res["target_id"] == "TID_2"


def test_launch_cdp_chrome_waits_for_existing_process_instead_of_spawning():
    bridge = srv.CDPBridge()
    cdp_states = iter([False, False, True])

    with (
        mock.patch.object(bridge, "_is_cdp_up", side_effect=lambda: next(cdp_states, True)),
        mock.patch.object(bridge, "_has_cdp_chrome_process", return_value=True),
        mock.patch.object(srv.time, "sleep"),
        mock.patch.object(srv.subprocess, "Popen") as popen_mock,
    ):
        assert bridge._launch_cdp_chrome("https://web.whatsapp.com", minimized=True, force_restart=False) is True
        popen_mock.assert_not_called()


def test_find_or_open_page_waits_for_cold_launch_tab_instead_of_opening_duplicate():
    bridge = srv.CDPBridge()
    bridge._last_launch_ts = time.time()

    calls = {"n": 0}

    class TransitionPage:
        def is_closed(self):
            return False

        @property
        def url(self):
            calls["n"] += 1
            return "about:blank" if calls["n"] <= 2 else "https://web.whatsapp.com/"

    page = TransitionPage()
    fake_ctx = mock.MagicMock()
    fake_ctx.pages = [page]
    fake_browser = mock.MagicMock()
    fake_browser.contexts = [fake_ctx]

    with (
        mock.patch.object(bridge, "_connect", return_value=fake_browser),
        mock.patch.object(bridge, "_pick_current_page", return_value=page),
        mock.patch.object(srv.time, "sleep"),
    ):
        got = bridge._find_or_open_page("whatsapp", launch_if_down=True)
        assert got is page
        fake_ctx.new_page.assert_not_called()


def test_find_or_open_page_reuses_blank_tab_instead_of_new_page():
    bridge = srv.CDPBridge()
    blank_page = mock.MagicMock()
    blank_page.is_closed.return_value = False
    blank_page.url = "about:blank"

    fake_ctx = mock.MagicMock()
    fake_ctx.pages = [blank_page]
    fake_browser = mock.MagicMock()
    fake_browser.contexts = [fake_ctx]

    now = [1000.0]

    def fake_time():
        now[0] += 0.6
        return now[0]

    with (
        mock.patch.object(bridge, "_connect", return_value=fake_browser),
        mock.patch.object(bridge, "_ensure_tab_tracker"),
        mock.patch.object(srv.time, "time", side_effect=fake_time),
        mock.patch.object(srv.time, "sleep"),
    ):
        got = bridge._find_or_open_page("whatsapp", launch_if_down=False)
        assert got is blank_page
        blank_page.goto.assert_called_once_with("https://web.whatsapp.com", wait_until="domcontentloaded", timeout=35000)
        fake_ctx.new_page.assert_not_called()


def test_whatsapp_ready_reports_loading_then_ready_and_brings_to_front():
    bridge = srv.CDPBridge()
    wa_page = mock.MagicMock()
    wa_page.is_closed.return_value = False
    wa_page.url = "https://web.whatsapp.com/"

    fake_ctx = mock.MagicMock()
    fake_ctx.pages = [wa_page]
    fake_browser = mock.MagicMock()
    fake_browser.contexts = [fake_ctx]

    states = iter([
        {"auth_required": False, "search_visible": False, "chat_list_ready": False, "startup_loading": True},
        {"auth_required": False, "search_visible": True, "chat_list_ready": False, "startup_loading": False},
        {"auth_required": False, "search_visible": True, "chat_list_ready": True, "startup_loading": False},
    ])

    with (
        mock.patch.object(bridge, "_is_cdp_up", return_value=True),
        mock.patch.object(bridge, "_connect", return_value=fake_browser),
        mock.patch.object(bridge, "_pick_current_page", return_value=wa_page),
        mock.patch.object(bridge, "_inspect_whatsapp_state", side_effect=lambda _p: next(states)),
    ):
        r1 = bridge.whatsapp_ready({})
        assert r1 == {"ok": True, "ready": False, "auth_required": False, "stage": "loading"}
        r2 = bridge.whatsapp_ready({})
        assert r2 == {"ok": True, "ready": False, "auth_required": False, "stage": "syncing"}
        r3 = bridge.whatsapp_ready({})
        assert r3 == {"ok": True, "ready": True, "auth_required": False, "stage": "ready"}
        assert wa_page.bring_to_front.call_count == 3


def test_whatsapp_ready_detects_auth_required():
    bridge = srv.CDPBridge()
    wa_page = mock.MagicMock()
    wa_page.is_closed.return_value = False
    wa_page.url = "https://web.whatsapp.com/"

    fake_ctx = mock.MagicMock()
    fake_ctx.pages = [wa_page]
    fake_browser = mock.MagicMock()
    fake_browser.contexts = [fake_ctx]

    with (
        mock.patch.object(bridge, "_is_cdp_up", return_value=True),
        mock.patch.object(bridge, "_connect", return_value=fake_browser),
        mock.patch.object(bridge, "_pick_current_page", return_value=wa_page),
        mock.patch.object(
            bridge,
            "_inspect_whatsapp_state",
            return_value={"auth_required": True, "search_visible": False, "chat_list_ready": False, "startup_loading": False},
        ),
    ):
        res = bridge.whatsapp_ready({})
        assert res == {"ok": True, "ready": False, "auth_required": True, "stage": "auth_required"}


def test_wait_whatsapp_ready_waits_past_10s_until_hydrated():
    bridge = srv.CDPBridge()
    wa_page = mock.MagicMock()
    now = [0.0]

    def fake_time():
        now[0] += 3.0
        return now[0]

    call_count = [0]

    def fake_inspect(_page):
        call_count[0] += 1
        if call_count[0] <= 5:  # > 15s elapsed
            return {"auth_required": False, "search_visible": False, "chat_list_ready": False, "startup_loading": True}
        return {"auth_required": False, "search_visible": True, "chat_list_ready": True, "startup_loading": False}

    with (
        mock.patch.object(bridge, "_inspect_whatsapp_state", side_effect=fake_inspect),
        mock.patch.object(srv.time, "time", side_effect=fake_time),
        mock.patch.object(srv.time, "sleep"),
    ):
        ok, err = bridge._wait_whatsapp_ready(wa_page, timeout_s=90.0)
        assert ok is True
        assert err == ""
        wa_page.bring_to_front.assert_called_once()



