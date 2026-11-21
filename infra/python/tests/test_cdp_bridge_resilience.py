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
