
# AIB Allow Watcher — Architecture Reference

> Last updated: 2026-05-05

## Overview

The AIB Allow Watcher is a centralized background monitor that detects the AI chat tool-call confirmation **Allow** button in **VS Code** and automatically clicks it after a brief user-confirmation window. Act starts it in a persistent VS Code-only mode at launch, and a single global 350 ms timer polls all enrolled windows simultaneously until the user turns it off.

---

## Trigger Sources

| Source | Hotkey / Event | AHK Call |
|--------|---------------|----------|
| `persistent_startup` | Act launches AppLaunchers with `/StartPersistentAllowWatcher` | `AIB_ArmAllowWatcher("persistent_startup", 0, "vscode", false, false)` |
| `chat_submit` | Enter in VS Code chat input | `AIB_ArmAllowWatcher("chat_submit", hwnd, "vscode", false, true)` |
| `chat_submit_cursor` | Enter in Cursor composer | Not armed (Enter is sent, watcher remains VS Code-only) |
| `start_implementation` | "Start Implementation" button disappears from UIA tree | `AIB_ArmAllowWatcher("start_implementation", 0, "vscode", false, true)` |
| `persistent_hotkey` | Win+Alt+Shift+9 manual toggle | `AIB_ArmAllowWatcher("persistent_hotkey", hwnd, "vscode", true, true)` |

**Scope semantics:**
- `"ide"` — normalized to `"vscode"` when VS Code-only restriction is enabled
- `"vscode"` — only `code.exe` windows
- `"cursor"` — normalized to `"vscode"` when VS Code-only restriction is enabled
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
    │   ├─ Check 8-minute session timeout → refresh state instead of removing when persistent mode is on
    │   ├─ Detect Allow button (AIB_WindowHasAllowButton)
    │   │   ├─ FOUND: run decision flow (2 s prep + 3 s Y/N prompt)
    │   │   │   ├─ Y / timeout → click Allow, re-arm for next occurrence
    │   │   │   └─ N → in persistent mode, skip this current Allow until it disappears
    │   │   └─ NOT FOUND (after seeing it): track absenceStreak
    │   │       ├─ send-ready check: reset session if Send button ready during persistent mode
    │   │       └─ absenceStreak ≥ 2: reset seenAllow, wait for next Allow
    │   └─ Reposition persistent banner if active monitor changed
    └─ If all sessions empty and persistent mode is off → AIB_StopAllowWatcher()
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
    skipCurrentAllow: false     ; Suppress current Allow until it disappears after N
    absenceStreak:  0           ; Consecutive ticks without Allow after first sight
    lastSeenTick:   <tickcount> ; Last poll time
    startedTick:    <tickcount> ; When session was created
    source:         "persistent_startup" | "persistent_hotkey" | "chat_submit" | "start_implementation"
    scope:          "vscode"
    timeoutMs:      480000      ; 8 minutes (g_AIB_AllowWatcherTimeoutMs)
}
```

Sessions are created by `AIB_AllowWatcherEnsureState()` and deleted on:
- Window close / disappearance from the current window list
- Timeout when persistent mode is off
- User presses N at decision prompt when persistent mode is off
- Send button becomes ready again when persistent mode is off
- Global stop (`AIB_StopAllowWatcher`)

When persistent mode is on, timeout/send-ready events reset the window session instead of removing it, and `N` sets `skipCurrentAllow := true` until the current Allow prompt disappears.

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

## Allow Selector Guardrails

The watcher only accepts an Allow candidate when **all** of the following are true:

- Button name contains standalone token `allow` (word boundary check)
- Button is inside AI chat confirmation context (`chat-confirmation-widget-container` and/or "Chat Confirmation Dialog" ancestry)
- Button name does **not** match negative filters (`proceed without executing`, `do not allow`, `don't allow`, etc.)

Search order:

1. Exact confirmation container scan (`AIB_FindAllowButtonInChatConfirmation`)
2. Framed button scan within confirmation bounds (`AIB_GetBestAllowButton`)
3. Document-subtree fallback scan for Electron/WebView tree drift (`AIB_FindAllowButtonInDocumentSubtree`)
4. Shortcut fallback (`Ctrl+Enter`) only when explicit confirmation hint exists
5. Optional OCR fallback (feature-flagged)

---

## OCR Last Resort (Optional)

An OCR fallback path is available but off by default:

- Flag: `g_AIB_AllowWatcherEnableOcrFallback := false`
- Helper script: `tools/Get-AllowPromptOcr.ps1`
- Cooldown: per-window debounce via `g_AIB_AllowWatcherOcrCooldownMs`

OCR path is allowed only when:

- Chat confirmation hint is present
- OCR text in the confirmation frame contains standalone `allow`
- OCR text also includes confirmation markers (`ctrl+enter`, `control+enter`, or `chat confirmation required`)
- OCR text does not match negative filters

If accepted, watcher performs one guarded click at OCR hit coordinates, then re-verifies prompt state.

---

## Termination

| Method | Behavior |
|--------|----------|
| Win+Alt+Shift+9 (toggle) | If active: calls `AIB_StopAllowWatcher`; if inactive: enables persistent mode immediately |
| Act startup | Launches AppLaunchers with `/StartPersistentAllowWatcher`, enabling persistent mode quietly |
| User presses N | In persistent mode, skips only the current Allow in that window until it disappears |
| 8-minute timeout per window | In persistent mode, session is refreshed instead of removed |

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

A separate 900 ms probe (`AIB_StartImplementationProbeTick`) scans IDE windows for the "Start Implementation" UIA button. When a window transitions from button-present → button-absent in the foreground, it arms the watcher with `"vscode"` scope.

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
| `AIB_IsPersistentAllowWatcherMode` | AppLaunchers.ahk | Report whether persistent mode is enabled |
| `AIB_WindowHasAllowButton` | AppLaunchers.ahk | UIA scan for Allow button |
| `AIB_WindowHasStartImplementationButton` | AppLaunchers.ahk | UIA Document scan for Start Implementation |
| `AIB_StartImplementationProbeTick` | AppLaunchers.ahk | 900 ms probe for button transition |
| `AIB_PersistentBannerCreate` | AppLaunchers.ahk | Create ⏳ emoji indicator |
| `AIB_PersistentBannerMonitorUpdate` | AppLaunchers.ahk | Reposition if monitor changed |
| `AIB_PersistentBannerDestroy` | AppLaunchers.ahk | Destroy indicator |
| `AIB_GetWatcherWindowHwnds` | AppLaunchers.ahk | Enumerate target IDE windows by scope |
| `AIB_IsAllowNegativeButtonName` | AppLaunchers.ahk | Reject non-target labels containing allow |
| `AIB_FindAllowButtonInDocumentSubtree` | AppLaunchers.ahk | Fallback search in Document nodes |
| `AIB_TryOcrAllowClick` | AppLaunchers.ahk | OCR-backed last-resort click path |
| `AIB_RunAllowOcrProbe` | AppLaunchers.ahk | Run PowerShell OCR helper and parse hit |

---

## Debug Log

Log file: `docs/aib-allow-runtime-debug.log`

Key signal patterns:

| Pattern | Meaning |
|---------|---------|
| `trigger-accepted source=... sessions=N` | Watcher armed, N windows enrolled |
| `tick allow-skip-armed hwnd=...` | User pressed N and current Allow will be ignored until it disappears |
| `tick allow-skip-cleared hwnd=...` | Skipped Allow disappeared; future prompts can be handled again |
| `start-impl-probe ide-windows=N` | Probe running, N IDE windows visible |
| `start-impl-btn doc-scan hwnd=... doc-btns=N` | Document traversal scanning N buttons |
| `start-impl-btn found hwnd=...` | Start Implementation button detected |
| `start-impl-probe arm scope=vscode hwnd=...` | Button transition detected, arming watcher |
| `tick allow-detected hwnd=...` | Allow button found in a window |
| `click-ocr hwnd=... x=... y=... text='...'` | OCR fallback fired a guarded click |
| `tick send-ready-complete hwnd=...` | Session completed (Send ready again) |
| `tick session-timeout hwnd=...` | Session expired after 8 min |
| `banner-created hwnd=... x=... y=...` | Persistent indicator placed |
| `banner-repositioned monitor=...` | Indicator moved to new monitor |
| `stop reason=...` | Watcher stopped, all sessions cleared |
