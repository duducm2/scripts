# Spotify.ahk — revision notes

## Purpose
Consolidates **Spotify** hotkeys and functions:
- Open/activate Spotify.
- Use `Volume_Up` / `Volume_Down` with modifier chords to control **Spotify volume** or **YouTube volume**.

Key UX goal: adjust Spotify volume **silently** without activating Spotify (WASAPI path).

## Runtime model / methodology
- **AutoHotkey v2**, `#SingleInstance Force`, `#UseHook` to reliably capture volume keys.
- Splits behavior based on modifier state:
  - **Ctrl held**: Spotify volume adjustment.
  - **Alt held**: YouTube tab volume adjustment (activate browser tab with YouTube URL).
  - Otherwise: pass through system volume.
- Uses **UIA-v2 Browser** integration to detect current tab URL for YouTube targeting.
- Feature flag:
  - `AL_USE_WASAPI := true` to prefer per-process volume control.

## Key entry points
- Hotkey `#!+s`: `OpenSpotify()`.
- `*Volume_Down` / `*Volume_Up`: `HandleVolumeDelta(deltaStep)`.
- Spotify helpers:
  - `GetSpotifyShortcutPath()`: resolves `.lnk` path without hardcoding user paths.
  - `GetSpotifyHwnd()`: strict HWND contract (integer or 0).
- YouTube helper:
  - `GetYouTubeTabHwnd()`: iterates a browser group and checks `uia.GetCurrentURL()`.
  - `IsYouTubeDomain(url)`.

## Internal structure
- **Spotify open strategy**:
  - If Spotify already running: activate and wait active.
  - Else: launch via Start Menu shortcut or Store `shell:AppsFolder` fallback.
- **Volume behavior**:
  - Ctrl+Volume: tries WASAPI adjustment by Spotify PID, otherwise activates Spotify and sends `^{Up}`/`^{Down}`.
  - Alt+Volume: activates YouTube tab and sends `{Up}`/`{Down}`.
  - Both paths attempt to restore previous focus/minimized state with timers.

## Dependencies
- `env.ahk`
- `SpotifyWASAPI.ahk` (per-process audio session control)
- `UIA-v2` (`UIA.ahk`, `UIA_Browser.ahk`)

## Notable patterns (as used here)
- **Modifier-dependent routing** with “restore focus” timers to reduce disruption.
- **Browser URL detection via UIA** instead of title heuristics.

## Known gap / risk worth noting
`SpotifyWASAPI.ahk` currently contains functions that return `false` (stubs for setting volume), which may prevent the intended “silent WASAPI volume adjust” behavior from working as designed. The fallback activation path covers functionality but is intrusive.

## DIP investigation output
### Findings summary
(paste here)

### Recommended changes
(paste here)

### Risks / regressions to watch
(paste here)

## DIP investigation inquiry
You are investigating `Spotify.ahk`, which uses modifier chords on Volume keys to control Spotify or YouTube volume, ideally without stealing focus. Key constraints: low latency, no intrusive window activation, and robustness across Spotify install types and browser variants.

Please research and report on:
- **Modern / alternative methods**:
  - Native media control APIs, Spotify APIs, Windows “App volume and device preferences”, and whether they can be controlled programmatically.
  - Browser-based approaches for YouTube volume (extension) vs UIA.
- **Code efficiency optimizations**:
  - More reliable per-process volume control implementation.
  - Faster YouTube tab detection (event hooks vs scanning all browser windows).
- **Parallelism / async feasibility in AutoHotkey**:
  - Background worker for audio session enumeration and caching.
- **Complex background macros without interrupting typing**:
  - Feasibility of adjusting volume without changing the foreground window (ideal).
  - If activation is unavoidable, best practices to restore focus deterministically and avoid keystroke leakage.

