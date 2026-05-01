# Microsoft Teams.ahk — revision notes

## Purpose
Consolidates **Microsoft Teams** hotkeys and automation helpers: resolving the correct Teams window (meeting vs chat), activating it reliably, and using UIA to find/click specific controls.

## Runtime model / methodology
- **AutoHotkey v2**, `#SingleInstance Force`.
- Focuses on **window resolution + activation reliability**:
  - Multiple candidate process names (`ms-teams.exe`, `Teams.exe`, `MSTeams.exe`).
  - Caches HWNDs for performance and stability, with validity checks.
  - Uses multi-attempt activation strategies (WinRestore/WinActivate + WinWaitActive + Win32 calls).
- Uses **UIA-v2** with caching (`UIA.CreateCacheRequest`) for repeated control queries.

## Key entry points
- **Auto-execute section**:
  - Includes `UIA-v2\Lib\UIA.ahk` and `Utils.ahk`.
  - Configures feature flags and activation parameters:
    - `TEAMS_USE_HWND_CACHE`
    - `TEAMS_ACTIVATION_ATTEMPTS`
    - `TEAMS_ACTIVATION_WAIT_MS`
    - `TEAMS_PROCESSES`
- **HWND cache class**:
  - `TeamsHwndCache` with `IsValid`, `InvalidateMeeting`, `InvalidateChat`.
- **Window resolution helpers**:
  - `ResolveTeamsMeetingHwnd()`
  - `ActivateTeamsMeetingWindow()`
  - `ActivateTeamsChatWindow()`
- **Activation strategy helper**:
  - `ActivateWindowWithRetry(hwnd, attempts, waitMs)`
- **UIA list-item helpers**:
  - `FindListItemByNames(...)`
  - `WaitListItemByNames(...)`

## Internal structure
- **Meeting vs chat identification**:
  - Title heuristics (exclude “Chat |”, “Sharing control bar |”, include meeting indicators).
  - Regex-based fallbacks for “... | Microsoft Teams ...”.
- **Activation robustness**:
  - Uses escalating attempts: `WinActivate`/`WinWaitActive`, then `SetForegroundWindow`, then `BringWindowToTop`.
- **UIA performance**:
  - Uses a toggle cache request: `TEAMS_TOGGLE_CACHE := UIA.CreateCacheRequest(["Name", "AutomationId"], ["Toggle"])`.

## Dependencies
- `UIA-v2` for control discovery.
- `Utils.ahk` for overlays/banners and shared UX behaviors.

## Notable patterns (as used here)
- **Cache + validity checks** to reduce repeated window scans.
- **Bounded waits** and retries with clear failure overlays.
- **Prefer API-like activation** (Win32 calls) when normal activation fails.

## DIP investigation output
### Findings summary
(paste here)

### Recommended changes
(paste here)

### Risks / regressions to watch
(paste here)

## DIP investigation inquiry
You are investigating `Microsoft Teams.ahk`, which automates Teams window targeting and UI actions. Key constraints: choosing the correct Teams surface (meeting/chat), reliable activation, and minimizing disruption to the user.

Please research and report on:
- **Modern / alternative methods**:
  - Teams automation alternatives: Graph API (where applicable), Teams app shortcuts, accessibility APIs.
  - If UI actions must be automated: compare UIA vs Power Automate Desktop vs WinAppDriver vs Playwright (for web Teams).
- **Code efficiency optimizations**:
  - More robust meeting/chat detection than title heuristics (UIA root properties, window class patterns, process metadata).
  - Event-driven HWND cache updates (WinEvent hooks) vs repeated scans.
- **Parallelism / async feasibility in AutoHotkey**:
  - How to run window resolution/UIA queries without blocking hotkey responsiveness.
  - Multi-process worker patterns.
- **Complex background macros without interrupting typing**:
  - Feasibility of muting/unmuting or UI actions via UIA Invoke without bringing Teams foreground.
  - Foreground restrictions and mitigation strategies (queued actions, passive confirmations).

