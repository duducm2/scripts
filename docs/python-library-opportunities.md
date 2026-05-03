# Python library opportunities

This report surveys the AutoHotkey-centric [scripts](../README.md) repository and its small [python](../python/) footprint. The goal is to highlight where **modern, well-maintained Python libraries** would add the most value without fighting the architecture (AHK owns hotkeys, UIA, and low-latency UI; Python owns optional daemons and IPC).

---

## 1. Current state

### 1.1 Python layout and dependencies

The `python/` tree hosts **persistent IPC daemons** and shared protocol code:

| Role | Files |
|------|--------|
| Daemons | [wm_daemon.py](../python/wm_daemon.py), [applauncher_daemon.py](../python/applauncher_daemon.py), [shiftkeys_daemon.py](../python/shiftkeys_daemon.py), [gemini_daemon.py](../python/gemini_daemon.py) |
| Protocols / framing | [protocol.py](../python/protocol.py), [wm_protocol.py](../python/wm_protocol.py), [al_protocol.py](../python/al_protocol.py), [shiftkeys_protocol.py](../python/shiftkeys_protocol.py) |
| Window / hook helpers | [wm_hooks.py](../python/wm_hooks.py), [al_window_enum.py](../python/al_window_enum.py) |
| ShiftKeys context / UIA | [shiftkeys_context.py](../python/shiftkeys_context.py), [shiftkeys_uia.py](../python/shiftkeys_uia.py) (UIA find/wait via **pywinauto** for daemon `FindElement` / `WaitElementState`) |
| JSON + envelope validation | [ipc_wire.py](../python/ipc_wire.py) (**orjson** + **pydantic**) used by `*_protocol.py` |
| Daemon logging | [daemon_logging.py](../python/daemon_logging.py) (**structlog**, JSON lines to stderr) |
| Tests | [python/tests/](python/tests/) (**pytest**), protocol round-trips |
| Harness / diagnostics | [wm_harness.py](../python/wm_harness.py), [compare_monitor_enumeration.py](../python/compare_monitor_enumeration.py) |

Declared dependencies are in [requirements.txt](../python/requirements.txt): **pywin32**, **lingua-language-detector**, **orjson**, **pydantic**, **pywinauto**, **structlog**, **pytest** (pytest is primarily for development/CI; daemons do not require it at runtime). **lingua** is used in [gemini_daemon.py](../python/gemini_daemon.py) for `OP_DETECT_LANG`.

### 1.2 JSON and IPC

On the Python side, framed payloads use **orjson** and minimal request validation via **pydantic** in [ipc_wire.py](../python/ipc_wire.py) (length-prefixed UTF-8 frames unchanged on the wire). On the AutoHotkey side, several IPC clients implement **hand-built JSON** and simplified decoders—for example [AppLauncherIPC.ahk](../aux/AppLauncherIPC.ahk) (`AL_IPC_JsonEncodeRequest`, `AL_IPC_DecodeResponse`). Similar patterns appear in [WMIPC.ahk](../aux/WMIPC.ahk) and [ShiftKeysIPC.ahk](../aux/ShiftKeysIPC.ahk). That keeps dependencies zero in AHK but increases the risk of edge-case bugs (escaping, nested objects, unicode) whenever the protocol evolves.

### 1.3 Where automation actually lives

- **UIA and browser automation** live in AHK with [UIA-v2/Lib/UIA.ahk](../UIA-v2/Lib/UIA.ahk) and related includes. Representative flows: [Gemini.ahk](../Gemini.ahk) (copy last message, async lookup, TTS), [GeminiToCursorBridge.ahk](../GeminiToCursorBridge.ahk), [WindowManagement.ahk](../WindowManagement.ahk).
- **Audio** for Spotify uses COM/WASAPI directly in AHK: [SpotifyWASAPI.ahk](../SpotifyWASAPI.ahk).
- **Bootstrap** in [Act.ahk](../Act.ahk) uses `RunWait` with `git fetch` / `git pull`; no Python today.
- **PowerShell** is spawned from [Utils.ahk](../Utils.ahk) for several utilities (for example mic volume and recycle-bin flows). Those are candidates for consolidation **only** if you explicitly want Python as a second scripting surface with tests and structured errors.

---

## 2. High-ROI library opportunities (by theme)

### 2.1 IPC protocols: validation, speed, and safety

**Shipped (Python):** [ipc_wire.py](../python/ipc_wire.py) already provides **orjson** serialization and **pydantic**-backed minimal request validation for all daemon `*_protocol.py` modules. Framing on the wire is unchanged.

**Still open / next steps:**

| Opportunity | Suggested libraries | Where it helps |
|-------------|---------------------|----------------|
| Binary or stricter schemas (optional) | **msgspec** (or richer **pydantic** models per op) | Only if you need smaller payloads or stronger response typing than today’s dict envelopes |
| Property-based tests for framing and unicode | **hypothesis** | Length-prefix boundaries, parity with AHK hand-built JSON in [aux/AppLauncherIPC.ahk](../aux/AppLauncherIPC.ahk), [aux/WMIPC.ahk](../aux/WMIPC.ahk), [aux/ShiftKeysIPC.ahk](../aux/ShiftKeysIPC.ahk) |

**Practical note:** You can keep AHK as the MMF writer and add a **small Python CLI used only in tests** to round-trip frames, instead of moving runtime JSON generation off AHK—unless profiling shows AHK string building is a bottleneck.

### 2.2 ShiftKeys daemon: UIA path (implemented; integration follow-ups)

[shiftkeys_uia.py](../python/shiftkeys_uia.py) is **implemented** with **pywinauto** (UIA): `find_element`, `wait_element_state`, and `find_gemini_stop_button` back the daemon’s `FindElement` / `WaitElementState` ops. **Optional next steps:** (1) expose `FindElement` / `WaitElementState` from [aux/ShiftKeysIPC.ahk](../aux/ShiftKeysIPC.ahk) so hotkeys actually hit the Python path today; (2) replace the **timer stub** for `OP_WATCH_UI_STATE` in [shiftkeys_daemon.py](../python/shiftkeys_daemon.py) with real UIA polling if you want non-blocking Gemini watch to match AHK’s blocking `WaitForStopResponseButton_Gemini`; (3) tune descendant-walk limits or add a dedicated STA worker if COM threading on the pipe thread ever flakes.

### 2.3 Observability and developer experience

**Shipped:** [daemon_logging.py](../python/daemon_logging.py) (**structlog**, JSON per line to stderr) on every IPC request in all four daemons; baseline **pytest** protocol round-trips in [python/tests/](python/tests/).

**Still open:**

| Opportunity | Suggested libraries | Notes |
|-------------|---------------------|--------|
| Cross-process log correlation | Same field names as AHK NDJSON ([GeminiToCursorBridge.ahk](../GeminiToCursorBridge.ahk) `Bridge_Log`, agent logs in [Shift keys.ahk](../Shift keys.ahk)) | Optional convention pass; daemons already emit `req_id`, `op`, `ok`, `ms_ms` |
| Nicer CLI for harnesses | **typer** or **click** | [wm_harness.py](../python/wm_harness.py), [compare_monitor_enumeration.py](../python/compare_monitor_enumeration.py) |
| Linting and static typing | **ruff**, **mypy** or **pyright** | Broader than the current small pytest suite |
| More pytest coverage | **pytest** | Daemon `handle_request` with mocks, edge payloads, regression tests for [ipc_wire.py](../python/ipc_wire.py) |

### 2.4 Resilience and concurrency (optional)

| Opportunity | Suggested libraries | Tradeoff |
|-------------|---------------------|----------|
| Retries with backoff inside daemon handlers | **tenacity** | Prefer **inside Python**, not from AHK `RunWait` loops |
| Async pipe handling | **asyncio** (and possibly a Windows event loop helper if needed) | Only worth it if threading becomes a measured bottleneck; current hotkey QPS is usually low |

### 2.5 Fuzzy text and window matching

If you **centralize** more window-title or path-segment matching in Python (for example alongside [wm_daemon.py](../python/wm_daemon.py) or app-launcher enumeration), **rapidfuzz** gives high-quality fuzzy ratios without pulling in heavy ML stacks. Today, heuristics live in AHK (for example [GeminiToCursorBridge.ahk](../GeminiToCursorBridge.ahk) segment matching).

### 2.6 Data files and reporting

The README documents CSV and INI under `data/` (Pomodoro log, Wikipedia completion, scroll positions). Parsing and analytics in AHK are fine for one-off reads; if you add **batch reporting** or dashboards, **polars** is a strong default for CSV. If you automate the Excel habit tracker opened from [Act.ahk](../Act.ahk), **openpyxl** is the usual choice.

### 2.7 Git and bootstrap (low priority)

[Act.ahk](../Act.ahk) already shells to Git successfully. **GitPython** or disciplined **subprocess** wrappers only pay off if you need structured handling of conflicts, auth prompts, or unified exit reporting—not for a straight `git pull` that either works or fails visibly.

### 2.8 Audio

**pycaw** and friends can control session volume in Python but largely **duplicate** what [SpotifyWASAPI.ahk](../SpotifyWASAPI.ahk) already does well in-process. Recommendation: **keep WASAPI in AHK** unless you need Python-side metering or multi-app analytics.

### 2.9 HTTP and LLM APIs (future)

Gemini workflows in this repo are **browser- and UIA-driven**, not REST clients. If [gemini_daemon.py](../python/gemini_daemon.py) ever runs queued tasks that call HTTP APIs, **httpx** plus **pydantic** response models is a modern, maintainable stack compared to raw `urllib`.

---

## 3. Per-script notes (requested entry scripts)

| Script | Python angle |
|--------|----------------|
| [Act.ahk](../Act.ahk) | Optional richer Git handling; otherwise unchanged |
| [AppLaunchers.ahk](../AppLaunchers.ahk) | **polars** (or scripts) if you add analytics over CSV/INI |
| [CheatSheetRich.ahk](../CheatSheetRich.ahk) | Native RichEdit Win32; **no Python** |
| [Gemini.ahk](../Gemini.ahk) | UIA hot path; Python for side tasks (language detect already in daemon) |
| [GeminiToCursorBridge.ahk](../GeminiToCursorBridge.ahk) | Bridge logic stays AHK; optional **structlog** correlation if daemons participate in debugging |
| [Microsoft Teams.ahk](../Microsoft Teams.ahk), [Outlook.ahk](../Outlook.ahk) | UIA-first; Python COM automation only if you deliberately add a heavy service layer |
| [mousemaster.ahk](../mousemaster.ahk) | Low-level input; **remain AHK** |
| [Shift keys.ahk](../Shift keys.ahk) | **shiftkeys_uia** is **done** in Python (daemon `FindElement` / `WaitElementState`); AHK does **not** yet call those ops over IPC—optional wiring in [aux/ShiftKeysIPC.ahk](../aux/ShiftKeysIPC.ahk). Optional **rapidfuzz** if fuzzy helpers move into the daemon |
| [Spotify.ahk](../Spotify.ahk), [SpotifyWASAPI.ahk](../SpotifyWASAPI.ahk) | Current WASAPI approach is appropriate |
| [Utils.ahk](../Utils.ahk) | Optional small Python CLIs with **typer** to replace scattered PowerShell where testability matters |
| [WindowManagement.ahk](../WindowManagement.ahk) | Already integrated with [wm_daemon.py](../python/wm_daemon.py); WM protocol uses **ipc_wire** (pydantic + orjson) |

---

## 4. Suggested priority order

1. **Done:** **pydantic** and **orjson** on shared protocol encode/decode ([ipc_wire.py](../python/ipc_wire.py)).
2. **Done:** **Real [shiftkeys_uia.py](../python/shiftkeys_uia.py)** (pywinauto UIA) for daemon `FindElement` / `WaitElementState`.
3. **Done:** **structlog** plus **pytest** — [daemon_logging.py](../python/daemon_logging.py) wired into all four daemons; [python/tests/](python/tests/) protocol round-trips (`python -m pytest python/tests`).
4. **rapidfuzz** if you move more window matching into Python for a single source of truth.
5. GitPython, openpyxl, httpx, and async refactors only when a **concrete feature** justifies the dependency and process boundary.

---

## 5. AHK versus Python boundary

```mermaid
flowchart TB
  subgraph ahk [AHK hot path]
    UIA[UIA v2]
    Input[Hooks and hotkeys]
    WASAPI[WASAPI COM]
  end
  subgraph py [Python daemons]
    Pipes[Named pipes / MMF]
    JsonStack[ipc_wire orjson pydantic]
    Proto[Op handlers]
    Logs[structlog stderr]
    Lingua[lingua detect]
    ShiftKeysUIA[shiftkeys_uia pywinauto]
  end
  ahk -->|IPC frames| Pipes
  Pipes --> JsonStack
  JsonStack --> Proto
  Proto --> Logs
  Proto --> Lingua
  Proto --> ShiftKeysUIA
```

**Rule of thumb:** keep anything that must complete in tens of milliseconds on a keypress in AHK; use Python for caching, enumeration, background tasks, validation-heavy JSON, and library ecosystems that are painful to reimplement in AHK.

---

## Appendix A: AutoHotkey files in this repository

Entry and domain scripts:

- [Act.ahk](../Act.ahk), [Shift keys.ahk](../Shift%20keys.ahk), [Utils.ahk](../Utils.ahk), [WindowManagement.ahk](../WindowManagement.ahk), [AppLaunchers.ahk](../AppLaunchers.ahk)
- [Gemini.ahk](../Gemini.ahk), [Microsoft Teams.ahk](../Microsoft%20Teams.ahk), [Outlook.ahk](../Outlook.ahk), [Spotify.ahk](../Spotify.ahk), [SpotifyWASAPI.ahk](../SpotifyWASAPI.ahk)
- [GeminiToCursorBridge.ahk](../GeminiToCursorBridge.ahk), [CheatSheetRich.ahk](../CheatSheetRich.ahk), [mousemaster.ahk](../mousemaster.ahk)

Supporting and library paths:

- IPC and harnesses: [aux/AppLauncherIPC.ahk](../aux/AppLauncherIPC.ahk), [aux/WMIPC.ahk](../aux/WMIPC.ahk), [aux/ShiftKeysIPC.ahk](../aux/ShiftKeysIPC.ahk), [aux/GeminiIPC.ahk](../aux/GeminiIPC.ahk), [aux/AL_IPC_Harness.ahk](../aux/AL_IPC_Harness.ahk), [aux/WM_IPC_Harness.ahk](../aux/WM_IPC_Harness.ahk)
- UIA v2: [UIA-v2/Lib/UIA.ahk](../UIA-v2/Lib/UIA.ahk), [UIA-v2/Lib/UIA_Browser.ahk](../UIA-v2/Lib/UIA_Browser.ahk), [UIA-v2/UIATreeInspector.ahk](../UIA-v2/UIATreeInspector.ahk)
- Other: [Lib/Media.ahk](../Lib/Media.ahk), [tools/MonitorEnumerationSnapshot.ahk](../tools/MonitorEnumerationSnapshot.ahk), [env.ahk](../env.ahk) (local / not always in repo)

---

## Appendix B: Python modules (roles)

| Module | Role |
|--------|------|
| [wm_daemon.py](../python/wm_daemon.py) | Named-pipe server for window-management IPC (structlog per request) |
| [daemon_logging.py](../python/daemon_logging.py) | Shared structlog JSON configuration for daemons |
| [wm_protocol.py](../python/wm_protocol.py), [wm_hooks.py](../python/wm_hooks.py) | WM message schema and optional hook-backed cache |
| [wm_harness.py](../python/wm_harness.py) | Client harness for latency / ping checks |
| [applauncher_daemon.py](../python/applauncher_daemon.py) | MMF-based app launcher daemon |
| [al_protocol.py](../python/al_protocol.py), [al_window_enum.py](../python/al_window_enum.py) | AL framing and window enumeration helpers |
| [shiftkeys_daemon.py](../python/shiftkeys_daemon.py) | Named-pipe server for ShiftKeys automation offload |
| [shiftkeys_protocol.py](../python/shiftkeys_protocol.py) | ShiftKeys message schema |
| [shiftkeys_context.py](../python/shiftkeys_context.py) | Foreground / context hooks (pywin32 COM message pump) |
| [shiftkeys_uia.py](../python/shiftkeys_uia.py) | UIA find/wait for daemon ops (pywinauto; COM init per call) |
| [ipc_wire.py](../python/ipc_wire.py) | orjson + pydantic envelope validation for protocols |
| [gemini_daemon.py](../python/gemini_daemon.py) | Gemini IPC daemon; task queue and **lingua** language detection |
| [protocol.py](../python/protocol.py) | Gemini daemon JSON framing helpers |
| [compare_monitor_enumeration.py](../python/compare_monitor_enumeration.py) | Diagnostic compare for monitor enumeration vs AHK |

---

*Generated as a codebase survey; implementing any library change should be done in small steps with daemon harnesses and AHK feature flags, per [windowmanagement-daemon-verify.md](windowmanagement-daemon-verify.md).*
