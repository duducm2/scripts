# AppLaunchers.ahk — revision notes

## Purpose
Consolidates **application/website launcher hotkeys** and related automation helpers (e.g., open/activate Cursor, open Desktop in Explorer, Wikipedia flows, etc.).

## Runtime model / methodology
- **AutoHotkey v2**, `#SingleInstance Force`.
- Central pattern: **hotkey → resolve target window/process → activate reliably → perform bounded automation**.
- Uses **feature flags** to phase in more advanced mechanisms (daemon/IPC/event hooks/FSM), while keeping a legacy path.
- Cleans up on exit (`OnExit`) to avoid leaving hooks/IPC in a bad state.

## Key entry points
- **Auto-execute section**:
  - Declares feature flags (`AL_USE_DAEMON`, `AL_USE_MMF_IPC`, `AL_USE_EVENT_HOOKS`, `AL_USE_WIKI_FSM`).
  - Includes:
    - `env.ahk`
    - `aux\AppLauncherIPC.ahk` (IPC setup/teardown)
    - `UIA-v2` libs
    - `Utils.ahk`
  - Registers `OnExit(AL_AppLaunchersExit, 1)` which calls:
    - `AL_RemoveInputGuard()`
    - `AL_UnregisterForegroundHook()`
    - `AL_IPC_Teardown()`
- **Representative hotkeys**:
  - `#!+n` (Win+Alt+Shift+N): open/activate **Cursor** with window-title heuristics; optionally uses IPC `ResolveCursorTargets`.
  - `+#e` (Shift+Win+E): open/activate **Desktop in Explorer**, maximize, refresh, select first item.

## Internal structure
- **Target resolution**:
  - IPC-enabled branch: `AL_IPC_Call("ResolveCursorTargets", ...)`.
  - Legacy branch: `WinGetList()` over `Cursor.exe` / `Code.exe`, skip “preview”, pick “best” title match.
- **Activation strategy**:
  - `WinActivate` + `WinWaitActive`, with additional Win32 calls in some flows (e.g., `SwitchToThisWindow`, `SetForegroundWindow`) when needed.
- **UX**:
  - Uses overlays and notifications from `Utils.ahk` (`StandardLoadingBar_Show`, `ShowCenteredOverlay_Utils`).
- **Safety**:
  - “safe input guard” infrastructure (replacing `BlockInput`) is present behind the scenes.

## Dependencies
- `env.ahk`: environment paths/behavior differences.
- `aux\AppLauncherIPC.ahk`: IPC and potential daemon integration.
- `UIA-v2`: UI automation for browser / structured UI interactions.
- `Utils.ahk`: overlays, helpers, scheduling.

## Notable patterns (as used here)
- **Feature-flagged rollout**: keep legacy resolution path; enable daemon/IPC progressively.
- **Bounded waits**: time-limited polling loops instead of unbounded sleeps.
- **Verify after action**: check window existence/activeness before sending keys.

## DIP investigation output
### Findings summary
(paste here)

### Recommended changes
(paste here)

### Risks / regressions to watch
(paste here)

## DIP investigation inquiry
You are investigating `AppLaunchers.ahk`, which triggers app/window activation and small automations via hotkeys. The critical constraints are reliability (correct target window), speed (low latency), and minimizing user interruption (avoid stealing focus unexpectedly or blocking typing).

Please research and report on:
- **Modern / alternative web-based methods**:
  - Replacements for “launcher hotkeys” using modern tooling (PowerToys, Windows shortcuts, AutoHotkey alternatives, web-to-desktop deep links).
  - For web workflows: browser extensions, automation via Playwright, or OS-level URI handlers.
- **Code efficiency optimizations**:
  - Faster window/process discovery than repeated `WinGetList` scans (WinEvent hooks, cached HWND maps, event-driven updates).
  - More deterministic Cursor targeting (beyond title heuristics).
- **Parallelism / async feasibility in AutoHotkey**:
  - Best practice patterns: multi-process daemons, named pipes/MMF, message-based queues, cooperative timers.
  - How to keep launch flows responsive while background work occurs.
- **Background macro feasibility without interrupting typing**:
  - Feasibility of performing “open app / activate / click / paste / return focus” flows asynchronously without disrupting user typing.
  - Constraints imposed by Windows foreground rules, accessibility APIs, and AHK’s single-threaded event loop.

