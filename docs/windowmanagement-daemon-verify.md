# WindowManagement Daemon – Cutover Verification

This document describes how to verify the polyglot refactor (daemon + IPC) and how to roll back safely.

## Rollback (legacy paths)

- All daemon-backed paths are guarded by **feature flags** (in `infra/ipc/WMIPC.ahk`):
  - `WM_USE_DAEMON`
  - `WM_USE_PIPE_IPC`
  - `WM_USE_EVENT_HOOK_CACHE`
- Default is **all off**. With flags off, the script behaves as before (legacy only).
- To **roll back** after enabling: set all three to `false` in `infra/ipc/WMIPC.ahk` (or override before the include) and reload the script. No daemon dependency.

## Enabling the daemon

1. Start the daemon: `python infra/python/wm_daemon.py` (keep it running).
2. In `infra/ipc/WMIPC.ahk` set:
   - `WM_USE_DAEMON := true`
   - `WM_USE_PIPE_IPC := true`
   - `WM_USE_EVENT_HOOK_CACHE := true`
3. Reload `WindowManagement.ahk`.

For intermittent **monitor 1 / wrong-screen** behavior when the daemon is on, see [windowmanagement-monitor1-debug.md](windowmanagement-monitor1-debug.md) (enumeration compare + `WM_DEBUG_MONITOR_MAP` logging).

## Verification checklist

### Functional parity

- **Monitor cycle / minimize** (e.g. hotkeys for monitor 1/2/3): uses daemon `GetVisibleWindowsByMonitor` when flags on, else legacy `GetVisibleWindowsOnMonitor`.
- **Monitor close** (Shift+Ctrl+Alt+Win+A/S/D/F): always uses **legacy** `GetVisibleWindowsOnMonitor(..., skipDaemon)` so the window list matches AHK `MonitorGet` for that monitor (avoids daemon/Enum order drift when closing from other displays).
- **Cursor project activation**: Activate by project path; same behavior; uses daemon `GetCursorWindows` / `ResolveProjectWindow` when flags on.
- **Preview window activation**: Same behavior; uses daemon `GetPreviewWindows` when flags on.
- **Cursor window selector sub-menu**: Same behavior; uses daemon `GetCursorWindows` when flags on.
- **Foreground window / cursor centering**: Timer runs at 250 ms when daemon on (daemon `GetForegroundWindowState`), 100 ms when off (legacy `WinExist("A")`).

### Performance

- With daemon: no repeated `WinGetList("ahk_exe Cursor.exe")` on hot paths; visibility list comes from daemon cache.
- With daemon: no 100 ms polling; 250 ms daemon-backed foreground check only when daemon is used.

### Reliability

- **Daemon down**: Each IPC call has a timeout; on failure the script falls back to the legacy path (e.g. `WinGetList`, local `GetVisibleWindowsOnMonitor`). No AHK restart required.
- **Reconnect**: If the pipe breaks, the next request will reconnect (see `WMIPC_SendRequest` in `infra/ipc/WMIPC.ahk`).
- **Latency**: Use `infra/ipc/WM_IPC_Harness.ahk` (or `python infra/python/wm_harness.py`) to measure Ping RTT; transport is sub-millisecond when daemon is running.

## Files

- **Daemon**: `infra/python/wm_daemon.py`, `infra/python/wm_protocol.py`, `infra/python/wm_hooks.py`
- **AHK client**: `infra/ipc/WMIPC.ahk`
- **Harness**: `infra/ipc/WM_IPC_Harness.ahk`, `infra/python/wm_harness.py`

## Python IPC unit tests (all daemons)

From the repo root: `pip install -r infra/python/requirements.txt` then `python -m pytest infra/python/tests`. This checks framed JSON for WM, ShiftKeys, AppLauncher, and Gemini protocols (no pipe server or AHK required). WM daemon verification above is separate (functional + harness).
