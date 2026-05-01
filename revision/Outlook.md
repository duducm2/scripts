# Outlook.ahk — revision notes

## Purpose
Consolidates **Outlook** hotkeys and automation helpers for both **classic Outlook** (`OUTLOOK.EXE`) and **new Outlook/Monarch** (`olk.exe`). It focuses on reliably locating the correct Outlook surface (mail, calendar, reminders) and invoking actions like **Read Aloud** with UIA where possible.

## Runtime model / methodology
- **AutoHotkey v2**, `#SingleInstance Force`.
- Uses **UIA-v2** for UI-bound actions (e.g., finding “Read Aloud” button in EN/PT).
- Uses **HWND caching** for performance and stability, with validity checks and invalidation.
- Strong use of **feature flags** to phase in improved backends:
  - `OUTLOOK_USE_HWND_CACHE`
  - `OUTLOOK_USE_STATE_WAITS`
  - `OUTLOOK_USE_COM_CORE`
  - `OUTLOOK_USE_WMCOMMAND_READALOUD`
  - `OUTLOOK_USE_UIA_READALOUD`

## Key entry points
- **Auto-execute section**:
  - Includes `UIA-v2\Lib\UIA.ahk` and `Utils.ahk`.
  - Defines process/window identification constants for classic vs new Outlook.
  - Defines timeout constants and feature flags.
- **Surface detection helpers**:
  - `OutlookTitleIsMailModule(title)`
  - `OutlookTitleIsCalendarModule(title)`
  - `OutlookTitleIsExcludedMainSurface(title)`
  - `OutlookWinActive()`
- **UI invocation abstraction**:
  - `InvokeReadAloudStart(hwnd)`:
    - If `OUTLOOK_USE_UIA_READALOUD`: uses UIA to find a button by name (`Read Aloud`, `Ler em voz alta`).
    - Optional WM_COMMAND path (placeholder / behind flag).
    - Falls back to synthetic `Alt+1`.
- **HWND cache**:
  - `class OutlookHwndCache` with `GetMailboxHwnd`, `GetCalendarHwnd`, `GetReminderHwnd` and resolvers scanning windows for both executables.

## Internal structure
- **Dual-process support**:
  - Treats both `outlook.exe` and `olk.exe` as valid Outlook processes.
  - Recognizes classic main class (`rctrl_renwnd32`) and new “Outlook Host”.
- **Reliability**:
  - Avoids blind key injection when UIA is available.
  - Uses “state waits” (flagged) to verify transitions rather than fixed sleeps.
- **Localization**:
  - Supports English + Portuguese UI strings for critical controls and titles.

## Dependencies
- `UIA-v2` for UI discovery/invocation.
- `Utils.ahk` for overlays/UX and shared helpers.

## Notable patterns (as used here)
- **Backend abstraction**: UIA vs WM_COMMAND vs keystroke fallback behind a single function (`InvokeReadAloudStart`).
- **HWND caching with invalidation** to reduce repeated scans and avoid HWND reuse pitfalls.

## DIP investigation output
### Findings summary
(paste here)

### Recommended changes
(paste here)

### Risks / regressions to watch
(paste here)

## DIP investigation inquiry
You are investigating `Outlook.ahk`, which targets both classic and new Outlook, including localized UI. Key constraints: reliably selecting the correct Outlook surface (mail/calendar/reminders), invoking actions without brittle keystrokes, and minimizing disruption to the user’s current typing.

Please research and report on:
- **Modern / alternative methods**:
  - Using Outlook APIs/COM (where supported) for actions instead of UI automation.
  - Microsoft Graph for message/calendar operations (when feasible).
  - Power Automate Desktop vs UIA vs COM trade-offs.
- **Code efficiency optimizations**:
  - Faster/safer window discovery and surface classification.
  - Event-driven caching (WinEvent hooks) to avoid repeated window list scans.
  - More reliable Read Aloud invocation (UIA patterns, automation IDs, command invocation).
- **Parallelism / async feasibility in AutoHotkey**:
  - Architectures to keep hotkeys responsive while UIA/COM calls run (worker process/daemon, queued actions).
- **Complex background macros without interrupting typing**:
  - Feasibility of invoking “Read Aloud” or switching modules without foreground activation.
  - Practical mitigation strategies if foreground activation is unavoidable (restore focus deterministically, passive overlays, batching).

