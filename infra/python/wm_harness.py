#!/usr/bin/env python3
"""Quick latency harness for WM daemon IPC (optional; AHK harness is canonical)."""
import sys
import time
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from wm_protocol import encode_message, decode_message, make_request

try:
    import win32file
    import struct
except ImportError:
    print("pip install pywin32 required")
    sys.exit(1)

PIPE_NAME = r"\\.\pipe\wm_automation"


def main():
    try:
        handle = win32file.CreateFile(
            PIPE_NAME,
            win32file.GENERIC_READ | win32file.GENERIC_WRITE,
            0,
            None,
            win32file.OPEN_EXISTING,
            0,
            None,
        )
    except OSError as e:
        print("Connect failed:", e)
        sys.exit(1)
    latencies = []
    n = 50
    for i in range(n):
        req = make_request(
            f"Ping-{i}", "Ping", "", {"clientTs": int(time.time() * 1000), "n": i}, 5000
        )
        data = encode_message(req)
        t0 = time.perf_counter()
        win32file.WriteFile(handle, data)
        err, header = win32file.ReadFile(handle, 4)
        if err or len(header) < 4:
            print("Read header failed")
            break
        (length,) = struct.unpack(">I", header)
        err, payload = win32file.ReadFile(handle, length)
        if err or len(payload) < length:
            print("Read body failed")
            break
        t1 = time.perf_counter()
        latencies.append((t1 - t0) * 1000)
    win32file.CloseHandle(handle)
    latencies.sort()
    p50 = latencies[int(n * 0.5)] if latencies else 0
    p95 = latencies[int(n * 0.95)] if latencies else 0
    print(
        f"WM IPC latency (Python, {len(latencies)} rounds): p50={p50:.2f} ms, p95={p95:.2f} ms"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
