# Monitor 1 intermittent issues — bottleneck identification

**Purpose:** Operational checklist and evidence capture for intermittent failures on the **leftmost** display (ordinal 1) in [WindowManagement.ahk](../WindowManagement.ahk). Aligns with [efficiency-canon.md](efficiency-canon.md) §3–4 (determinism, silent fallbacks, IPC contracts).

**Related:** [windowmanagement-daemon-verify.md](windowmanagement-daemon-verify.md), [display-environment.md](display-environment.md).

---

## Step 0: Classify each incident (required)

Record **one line per occurrence** before deeper debugging:

| Symptom class                | What to write                                                         |
| ---------------------------- | --------------------------------------------------------------------- |
| **Wrong physical screen**    | e.g. “Ctrl+Alt+Win+A moved window to monitor 3, not leftmost.”        |
| **Empty / false no windows** | e.g. “Toast: no windows on monitor 1; Chrome visible on left screen.” |
| **Cursor / flash**           | e.g. “Red flash centered on primary, window was on M1.”               |
| **Move / maximize**          | e.g. “Window snapped back after maximize on M1 only.”                 |

**Incident log (append rows):**

| Date/time | Symptom class | Hotkey / action | Daemon flags on? (Y/N) | Notes |
| --------- | ------------- | --------------- | ---------------------- | ----- |
|           |               |                 |                        |       |

---

## Step 1: A/B — daemon path vs legacy

`GetVisibleWindowsOnMonitor` uses the WM automation daemon when **all** of these are true (see [infra/ipc/WMIPC.ahk](../infra/ipc/WMIPC.ahk)):

- `WM_USE_DAEMON := true`
- `WM_USE_PIPE_IPC := true`
- `WM_USE_EVENT_HOOK_CACHE := true`

**Procedure:**

1. Reproduce with flags **on** (current failing setup if applicable). Note whether the issue appears.
2. Set all three to **false**, restart [WindowManagement.ahk](../WindowManagement.ahk) (and ensure no daemon required for other flows you need).
3. Reproduce the same action.

**Interpretation:**

- If the problem **disappears** with the daemon off, prioritize **AHK monitor index vs Python `EnumDisplayMonitors` order** (see Step 2 and [infra/python/compare_monitor_enumeration.py](../infra/python/compare_monitor_enumeration.py)).
- If it **persists**, focus on legacy geometry (`MonitorGet`, `MonitorFromPoint`, sleeps, OS “remember window locations”).

---

## Step 2: Compare enumeration maps (one snapshot)

Run **both** from the same machine state (no display topology change between runs):

1. **Python** (requires `pywin32`; same stack as `wm_daemon.py`):

   ```text
   cd infra/python
   python compare_monitor_enumeration.py
   ```

2. **AutoHotkey** (mirrors `GetMonitorIndexByOrder` sort in [WindowManagement.ahk](../WindowManagement.ahk)):

   ```text
   "C:\Path\To\AutoHotkey64.exe" infra\tools\MonitorEnumerationSnapshot.ahk
   ```

   Output file: [infra/tools/MonitorEnumerationSnapshot-out.txt](../infra/tools/MonitorEnumerationSnapshot-out.txt) (next to the script).

**Check:** For **ordinal 1** (leftmost by `cx`, then `cy`), compare:

- AHK **resolved index** `idx` and work-area rect.
- Python **EnumDisplayMonitors** slot `i` and `HMONITOR`.

If ordinal 1’s AHK `idx` is `k` but Python’s `k`th monitor is a **different** rectangle or handle class, the daemon’s `GetVisibleWindowsByMonitor(k)` can target the wrong screen while legacy AHK path stays correct.

**Mitigation in tree:** [WindowManagement.ahk](../WindowManagement.ahk) filters daemon window lists so each HWND’s `MonitorFromWindow` matches the `HMONITOR` from `MonitorGet(mon)` work-area center; on mismatch it falls back to legacy enumeration.

---

## Step 3: NDJSON logging

Optional `WM_DEBUG_MONITOR_MAP` / `debug-2f65b1.log` hooks were **removed** after the monitor-close fix. Use Steps 1–2 and temporary local logging if you need new evidence.

---

## Step 4: Map observations to canon taxonomy

| If you observe…                         | Canon bucket ([efficiency-canon.md](efficiency-canon.md) §3)                      |
| --------------------------------------- | --------------------------------------------------------------------------------- |
| Daemon vs legacy different results      | Repeated enumeration; non-deterministic fallbacks (silent catch + alternate path) |
| Wrong monitor only when daemon succeeds | Hardcoded literals / **contract mismatch** at IPC boundary                        |
| Cursor wrong on M1 only                 | Polling (`MonitorActiveWindow` interval) + race with focus                        |
| Move fails sporadically                 | Blocking sleeps; timing vs bounded waits elsewhere                                |

---

## Step 5: Evidence table (fill after investigation)

| Symptom | Hypothesis | Test | Result | Fix direction |
| ------- | ---------- | ---- | ------ | ------------- |
|         |            |      |        |               |

**Suggested fix directions:**

- **Done in repo:** AHK-side validation + legacy fallback (see note above).
- Further hardening: align Python `EnumDisplayMonitors` ordering with AHK, or pass **HMONITOR** / rect from AHK in the IPC payload.
- Harden `MonitorFromPoint` packing for negative virtual-screen coordinates if legacy path mis-fires on the left monitor only.
- Increase or replace fixed `Sleep` in `MoveWinToMonitor` with bounded condition waits if races correlate with animation.
