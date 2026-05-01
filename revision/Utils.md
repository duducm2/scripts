# Utils.ahk — revision notes

## Purpose
Shared utilities used across the automation suite: overlays/banners, media/sound helpers, UIA helpers (especially for Gemini/Chrome), and general-purpose automation primitives.

This file is effectively the **core shared library** for the repo’s AHK scripts.

## Runtime model / methodology
- **AutoHotkey v2** library module, included by many scripts.
- Mixes:
  - UI/UX primitives (centered overlays, loading bars, notifications).
  - UIA automation helpers (Gemini prompt field discovery, Chrome tab helpers).
  - Media/audio helpers (via `lib\Media.ahk`, `SpotifyWASAPI.ahk`).
- Some “agent debug log” functions are intentionally no-op, indicating a pattern of **runtime-evidence logging that can be enabled/disabled**.

## Key entry points (representative)
- **Gemini UIA helpers**:
  - `FindGeminiPromptField(uia)`: robust multi-strategy prompt field finder using multiple EN/PT names and class heuristics.
  - `FocusGeminiAskFieldForHwnd(geminiHwnd, playChime := false)`: focuses the Gemini prompt field when the Chrome window is already active (avoids stealing focus).
  - `GetChromeActiveTabIndex(uia)`: returns active tab index/count via UIA.
  - Gemini model picker helpers:
    - `FindGeminiModePickerButton(uia)`
    - `GeminiNormalizeModelName(...)`
    - `GeminiNormalizeModelLabel(...)`
    - `GetGeminiActiveModelFromPickerOnly(uia)`
- **Banner/overlay constants** (defined early):
  - `BANNER_ACCENT_SUCCESS`, `BANNER_ACCENT_ERROR`, `BANNER_ACCENT_INTERMEDIATE`, `BANNER_ACCENT_INFO`
- **Media/sound helpers** (via includes):
  - `ScriptSoundPlay`, `ScriptSoundBeep`, `IsSoundEnabled` (and/or related helpers).

## Internal structure
- **Resilience to UI drift**:
  - Repeated fallback passes for UIA queries (exact name + ControlType → broader searches).
  - Heuristics based on class names like `ql-editor`.
- **Non-intrusive focus policy**:
  - Several helpers explicitly avoid `WinActivate` and require the caller to ensure the window is foreground before focusing UIA elements.
- **Localization**:
  - Supports English and Portuguese UI strings for Gemini controls.

## Dependencies
- `env.ahk`
- `UIA-v2\Lib\UIA.ahk`, `UIA-v2\Lib\UIA_Browser.ahk`
- `lib\Media.ahk`
- `SpotifyWASAPI.ahk`
- Plus repo-local assets (e.g., sound files under `sounds\`).

## Notable patterns (as used here)
- **Centralized shared capabilities**: other scripts depend on Utils rather than re-implementing UIA discovery and overlays.
- **Early constants**: banner colors and config defined before other initialization to avoid order issues.
- **“No-op” debug log functions**: keep callsites intact but disable overhead in production.

## DIP investigation output
### Findings summary
(paste here)

### Recommended changes
(paste here)

### Risks / regressions to watch
(paste here)

## DIP investigation inquiry
You are investigating `Utils.ahk`, the shared foundation for most scripts in this repo. Key constraints: correctness under UI changes (Gemini/Chrome), consistent UX overlays, and keeping performance acceptable because many hotkeys depend on these helpers.

Please research and report on:
- **Modern / alternative methods** for the shared capabilities:
  - Replacing UIA browser automation with extensions or Playwright daemon.
  - Replacing custom overlays with Windows-native notification frameworks or WebView2 overlays.
- **Code efficiency optimizations**:
  - Reduce UIA tree traversal cost (scoping roots, caching, avoiding redundant `FindAll`).
  - Reduce overlay rendering overhead and avoid focus issues.
  - Organize shared utilities into smaller modules for faster load times and clearer ownership.
- **Parallelism / async feasibility in AutoHotkey**:
  - Whether expensive UIA discovery should be delegated to a helper process/daemon with caching.
  - Patterns for async operations while preserving deterministic UX.
- **Complex background macros without interrupting typing**:
  - What subset of actions can be done without activation (UIA Invoke, background messaging).
  - Recommended architecture for queued background actions with minimal user interruption.

