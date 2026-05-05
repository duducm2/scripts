# AIB Allow Watcher — Architecture Reference

> Last updated: 2026-05-05

## Overview

The AIB Allow Watcher is a centralized background monitor that detects when VS Code or Cursor surfaces an "Allow" tool-call confirmation button and automatically clicks it after a brief user-confirmation window. A single global 350 ms timer polls all enrolled IDE windows simultaneously; stopping the watcher immediately terminates monitoring across all windows.

---

## Trigger Sources

| Source | Hotkey / Event | AHK Call |
|--------|---------------|----------|
| `chat_submit` | Enter in VS Code chat input | `AIB_ArmAllowWatcher("chat_submit", hwnd, "ide", false, true)` |
| `chat_submit_cursor` | Enter in Cursor composer | `AIB_ArmAllowWatcher("chat_submit_cursor", hwnd, "cursor", true, true)` |
| `start_implementation` | "Start Implementation" button disappears from UIA tree | `AIB_ArmAllowWatcher("start_implementation", 0, "ide", false, true)` |
| `rapid_fire_hotkey` | Win+Alt+Shift+9 manual toggle | `AIB_ArmAllowWatcher("rapid_fire_hotkey", hwnd, "vscode", true, true)` |

**Scope semantics:**
- `"ide"` — enroll all open `code.exe` and `cursor.exe` windows (used by `chat_submit` and `start_implementation`)
- `"vscode"` — only `code.exe` windows
- `"cursor"` — only `cursor.exe` windows, pinned to specific hwnd
- `pinToTarget := false` — allow the tick loop to auto-enroll additional windows that open after arming

---

## Lifecycle

```
User action (Enter / button click / hotkey)
    ↓
AIB_ArmAllowWatcher()
    ├─ Calls AIB_StartAllowWatcher() — starts timer if not already running
    ├─ Calls AIB_AllowWatcherEnsureState() for each window in scope
    ├─ Shows temporary trigger banner (1.3 s)
    └─ Creates persistent ⏳ emoji indicator (bottom-left of active monitor)
    ↓
AIB_AllowWatcherTick()  [every 350 ms]
    ├─ Auto-enrolls any new IDE windows opened after arming
    ├─ Round-robin across all enrolled windows (fairness)
    ├─ For each window:
    │   ├─ Check 8-minute session timeout → remove if expired
    │   ├─ Detect Allow button (AIB_WindowHasAllowButton)
    │   │   ├─ FOUND: run decision flow (2 s prep + 3 s Y/N prompt)
    │   │   │   ├─ Y / timeout → click Allow, re-arm for next occurrence
    │   │   │   └─ N → remove this window's session
    │   │   └─ NOT FOUND (after seeing it): track absenceStreak
    │   │       ├─ send-ready check (non-rapid-fire): remove session if Send button ready
    │   │       └─ absenceStreak ≥ 2: reset seenAllow, wait for next Allow
    │   └─ Reposition persistent banner if active monitor changed
    └─ If all sessions empty → AIB_StopAllowWatcher()
    ↓
AIB_StopAllowWatcher()
    ├─ Kills timer
    ├─ Clears ALL sessions (g_AIB_AllowWatcherStateByHwnd reset)
    ├─ Destroys persistent emoji banner
    └─ Optionally shows completion toast
```

---

## Session Model

Each enrolled window has an independent session stored in `g_AIB_AllowWatcherStateByHwnd` (Map, keyed by hwnd string):

```
{
    seenAllow:      false       ; Has Allow button been spotted yet?
    absenceStreak:  0           ; Consecutive ticks without Allow after first sight
    lastSeenTick:   <tickcount> ; Last poll time
    startedTick:    <tickcount> ; When session was created
    source:         "chat_submit" | "start_implementation" | "rapid_fire_hotkey"
    scope:          "ide" | "vscode" | "cursor"
    timeoutMs:      480000      ; 8 minutes (g_AIB_AllowWatcherTimeoutMs)
}
```

Sessions are created by `AIB_AllowWatcherEnsureState()` and deleted on:
- Timeout (8 min from startedTick)
- User presses N at decision prompt
- Send button becomes ready again (non-rapid-fire sources)
- Global stop (`AIB_StopAllowWatcher`)

---

## Global Timer

| Variable | Value | Purpose |
|----------|-------|---------|
| `g_AIB_AllowWatcherTimer` | Function ref | Single 350 ms repeating timer |
| `g_AIB_AllowWatcherActive` | bool | Master active flag |
| `g_AIB_AllowWatcherPromptLock` | bool | Mutex — only one decision overlay at a time |
| `g_AIB_AllowWatcherRoundRobinOffset` | int | Rotates per tick for polling fairness |

**One timer, all windows** — the timer is started once when the first session is created and killed once all sessions are gone or the watcher is stopped. New windows opening after arming are auto-enrolled on the next tick.

---

## Termination

| Method | Behavior |
|--------|----------|
| Win+Alt+Shift+9 (toggle) | If active: calls `AIB_StopAllowWatcher` — kills ALL sessions immediately |
| All sessions expire naturally | `AIB_AllowWatcherTick` calls `AIB_StopAllowWatcher` automatically |
| User presses N | Removes only that window's session; other sessions continue |
| 8-minute timeout per window | Session auto-removed; others unaffected |

---

## Persistent Flow Indicator

A tight emoji-only GUI (`⏳`) shows at the bottom-left of the **active monitor** for the duration of the flow:

- Created in `AIB_PersistentBannerCreate()` at arm time
- Repositioned in `AIB_PersistentBannerMonitorUpdate()` each tick if active monitor changes
- Destroyed in `AIB_PersistentBannerDestroy()` when watcher stops
- Semi-transparent (alpha 210), no caption, AutoSize to emoji dimensions
- Stored in `g_AIB_FlowBannerHwnd`; `g_AIB_FlowBannerLastMonitorIdx` tracks last known monitor

---

## Start Implementation Probe

A separate 900 ms probe (`AIB_StartImplementationProbeTick`) scans all IDE windows for the "Start Implementation" UIA button. When a window transitions from button-present → button-absent in the foreground, it arms the watcher with `"ide"` scope (enrolling all open VS Code/Cursor windows).

The probe is blocked while the watcher is active (`g_AIB_AllowWatcherActive = true`).

Button detection uses Document-level UIA traversal (VS Code embeds buttons inside a `RootWebArea` Document element not reachable via direct window `FindAll`).

---

## Key Functions

| Function | File | Purpose |
|----------|------|---------|
| `AIB_ArmAllowWatcher` | AppLaunchers.ahk | Public entry point: start watcher + banners |
| `AIB_StartAllowWatcher` | AppLaunchers.ahk | Start timer, enroll windows |
| `AIB_StopAllowWatcher` | AppLaunchers.ahk | Kill timer, clear all state, destroy banner |
| `AIB_AllowWatcherTick` | AppLaunchers.ahk | Main 350 ms loop |
| `AIB_AllowWatcherEnsureState` | AppLaunchers.ahk | Create/update per-window session |
| `AIB_RunAllowDecisionFlow` | AppLaunchers.ahk | 2 s prep + 3 s Y/N decision overlay |
| `AIB_WindowHasAllowButton` | AppLaunchers.ahk | UIA scan for Allow button |
| `AIB_WindowHasStartImplementationButton` | AppLaunchers.ahk | UIA Document scan for Start Implementation |
| `AIB_StartImplementationProbeTick` | AppLaunchers.ahk | 900 ms probe for button transition |
| `AIB_PersistentBannerCreate` | AppLaunchers.ahk | Create ⏳ emoji indicator |
| `AIB_PersistentBannerMonitorUpdate` | AppLaunchers.ahk | Reposition if monitor changed |
| `AIB_PersistentBannerDestroy` | AppLaunchers.ahk | Destroy indicator |
| `AIB_GetWatcherWindowHwnds` | AppLaunchers.ahk | Enumerate target IDE windows by scope |

---

## Debug Log

Log file: `docs/aib-allow-runtime-debug.log`

Key signal patterns:

| Pattern | Meaning |
|---------|---------|
| `trigger-accepted source=... sessions=N` | Watcher armed, N windows enrolled |
| `start-impl-probe ide-windows=N` | Probe running, N IDE windows visible |
| `start-impl-btn doc-scan hwnd=... doc-btns=N` | Document traversal scanning N buttons |
| `start-impl-btn found hwnd=...` | Start Implementation button detected |
| `start-impl-probe arm scope=ide hwnd=...` | Button transition detected, arming watcher |
| `tick allow-detected hwnd=...` | Allow button found in a window |
| `tick send-ready-complete hwnd=...` | Session completed (Send ready again) |
| `tick session-timeout hwnd=...` | Session expired after 8 min |
| `banner-created hwnd=... x=... y=...` | Persistent indicator placed |
| `banner-repositioned monitor=...` | Indicator moved to new monitor |
| `stop reason=...` | Watcher stopped, all sessions cleared |
