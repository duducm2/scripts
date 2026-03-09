#!/usr/bin/env python3
"""
Persistent Gemini IPC daemon. Listens on a local TCP port (default 29512) for
length-prefixed JSON frames. AHK connects once and streams requests; Python
handles API/JSON/streaming. Boundary: AHK = hotkeys, UIA, clipboard; Python =
network, JSON, text pipelines. No RunWait from AHK per shortcut.
"""

import json
import socket
import struct
import sys
from protocol import encode_message, decode_message, validate_request, CMD_PING

DEFAULT_PORT = 29512
MAX_FRAME = 1024 * 1024


def handle_request(req: dict) -> dict:
    """Synchronous request handler. Extend for API/streaming work."""
    if not validate_request(req):
        return {"id": req.get("id", ""), "ok": False, "error": "invalid request"}
    cmd = req.get("id")
    if cmd == CMD_PING:
        return {"id": cmd, "ok": True, "pong": True}
    return {"id": cmd, "ok": False, "error": "unknown command"}


def read_frame_from_socket(conn: socket.socket) -> dict | None:
    header = conn.recv(4)
    if len(header) < 4:
        return None
    (length,) = struct.unpack(">I", header)
    if length > MAX_FRAME:
        return None
    payload = b""
    while len(payload) < length:
        chunk = conn.recv(length - len(payload))
        if not chunk:
            return None
        payload += chunk
    try:
        return json.loads(payload.decode("utf-8"))
    except (json.JSONDecodeError, UnicodeDecodeError):
        return None


def serve_connection(conn: socket.socket) -> None:
    try:
        while True:
            msg = read_frame_from_socket(conn)
            if msg is None:
                break
            resp = handle_request(msg)
            conn.sendall(encode_message(resp))
    except (ConnectionResetError, BrokenPipeError, OSError):
        pass
    finally:
        conn.close()


def run_tcp_server(port: int = DEFAULT_PORT) -> None:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        s.bind(("127.0.0.1", port))
        s.listen(1)
        while True:
            conn, _ = s.accept()
            serve_connection(conn)


if __name__ == "__main__":
    port = int(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_PORT
    run_tcp_server(port)
