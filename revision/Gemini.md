# Gemini.ahk — revision notes

## Purpose
Automates **Gemini (web UI in Chrome)** workflows: activating the right tab/window, focusing the prompt field, copying the last assistant response, and triggering “listen/read aloud” style actions. It emphasizes robustness against UI changes and race conditions.

## Runtime model / methodology
- **AutoHotkey v2**, `#SingleInstance Force`.
- Heavy use of **UI Automation** via `UIA-v2` (`UIA.ahk`, `UIA_Browser.ahk`) to avoid brittle coordinate clicking.
- Uses many **named constants** for timeouts/retries and minimizes “magic sleeps”.
- Incorporates **bounded polling loops** and **retry strategies** (including exponential backoff in some flows).
- Feature flags for phased refactors:
  - `GEMINI_USE_WIN_EVENT_HOOK`
  - `GEMINI_USE_PYTHON_IPC`
  - performance logging toggle (`GEMINI_PERF_LOG_ENABLED`)

## Key entry points
- **Auto-execute section**:
  - Includes `UIA-v2`, `env.ahk`, `Utils.ahk`, and IPC helpers (`aux\WMIPC.ahk`, `aux\GeminiIPC.ahk`).
  - Defines extensive constants for timings, retries, clipboard sync.
- **Hotkeys (from scan)**:
  - `#!+o`, `#!+p`, `#!+7`, `#!+8`, `#!+i` (Win+Alt+Shift+...) are bound in this script and represent the user-facing Gemini actions (open/focus, copy, TTS, etc.).
- **Core helper functions (examples of “public” utilities)**:
  - `CopyLastGeminiMessageWithRetry(...)`: layered retry wrapper.
  - `GetLastGeminiCopyButton(uia)`: chooses the visually-lowest “Copy” button (reduces UI drift issues).
  - `FindGeminiPauseResumeButton(uia, which)`: finds TTS controls.
  - Scoped discovery helpers (`GetGeminiSearchRoot`, `GetGeminiMoreOptionsButtonsScoped`, `FindGeminiTextToSpeechMenuItem`).

## Internal structure
- **UI element discovery**:
  - Centralized “find all buttons then filter” approach to reduce repeated tree walks.
  - Fallback strategies: strict ControlType constants first, then more permissive string-based searches.
  - Scoping searches to main pane to avoid full-document traversal when possible.
- **Clipboard correctness**:
  - Uses `GetClipboardSequenceNumber()` (Win32) for low-latency “clipboard changed” detection instead of fixed sleeps.
  - Minimum clipboard length guard (`GEMINI_TRANSFER_MIN_CLIPBOARD_LENGTH`).
- **Observability**:
  - Optional perf logging to `.cursor\gemini_perf.log`.

## Dependencies
- `UIA-v2` (`UIA.ahk`, `UIA_Browser.ahk`).
- `Utils.ahk` for shared UIA helpers (e.g., locating prompt field), overlays, and shared UX.
- `aux\WMIPC.ahk` / `aux\GeminiIPC.ahk` for IPC/daemon phases.
- `env.ahk` for environment-specific behavior and paths.

## Notable patterns (as used here)
- **Bounded waits + retries** everywhere UI state is volatile.
- **Verification layers**: prefer confirming UI elements/windows over assuming success after a sleep.
- **Refactor via feature flags** rather than big-bang rewrites.

## DIP investigation output
### Findings summary
(paste here)

### Recommended changes
(paste here)

### Risks / regressions to watch
(paste here)

## DIP investigation inquiry
You are investigating `Gemini.ahk`, which automates Gemini in a browser using UIA. The critical constraints are robustness to DOM/accessibility tree changes, low latency, and—importantly—support for “background” style macros that don’t disrupt the user’s active typing.

Please research and report on:
- **Modern / alternative web-based methods** to achieve the same objectives:
  - Browser extension approaches (content scripts + messaging).
  - Playwright/Puppeteer automation (local daemon) vs UIA.
  - Gemini APIs (if available/appropriate) as a replacement for UI automation.
- **Code efficiency optimizations** for UIA automation:
  - Minimizing tree traversal (`FindAll`) cost; caching; scoping roots.
  - More reliable selectors than button name/class heuristics.
  - Event-driven readiness (WinEvent hooks, accessibility events) vs polling.
- **Parallelism / async feasibility in AutoHotkey**:
  - Viable designs: separate helper process/daemon for UIA queries; named pipes; cooperative timers.
  - How to keep AHK hotkeys responsive during long UIA operations.
- **Complex background macro feasibility without interrupting typing**:
  - Is it feasible to run a sequence like “activate Gemini tab → move mouse/click/focus → paste → submit → restore focus” asynchronously while the user keeps typing elsewhere?
  - Windows foreground restrictions (SetForegroundWindow rules), input injection limitations, and possible mitigations (e.g., UIA Invoke without activation, background message posting, browser extension side-channel, clipboard + deferred paste).

