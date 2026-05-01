# SpotifyWASAPI.ahk — revision notes

## Purpose
Implements (or intends to implement) **WASAPI per-process volume control**, used by `Spotify.ahk` to adjust Spotify playback volume without activating the Spotify window.

## Runtime model / methodology
- Library-style AutoHotkey v2 module of COM/WASAPI helpers.
- Uses COM interop patterns:
  - `ComObjQuery` / QueryInterface-like helpers.
  - `ComCall` with explicit vtable indexes.
- Builds the session manager through `MMDeviceEnumerator` and session APIs.

## Key entry points (current)
- Session PID resolution:
  - `GetAudioSessionProcessId(sessionObj)` queries `IAudioSessionControl2` and calls `GetProcessId` (vtable index 14).
- Manager resolver:
  - `GetDefaultSessionManager()` builds `IAudioSessionManager2` via default audio endpoint.
- Lower-level QI:
  - `QueryInterface(obj, iidStr)`
- Session scalar read:
  - `WASAPI_GetSessionScalar(sessionObj, &outScalar, mgr := 0)` tries `ISimpleAudioVolume` first, then a grouping-based fallback via manager.

## Current behavior gap
Several functions are intentionally stubbed and return immediately:
- `WASAPI_SetSessionScalar(...)` returns `false`
- `AdjustProcessVolumeByPid(...)` returns `false`
- `SetProcessPlaybackVolumePercent(...)` returns `false`
- `ApplyAutoHotkeyAudioSessionsVolumePercent(...)` returns `0`

This means the “silent volume adjust” path in `Spotify.ahk` likely cannot work unless those functions were replaced elsewhere or the code path no longer calls them (worth verifying in runtime).

## Dependencies
- Windows COM (`ole32`) and audio endpoint/session COM interfaces.
- Correct IID/CLSID constants (declared at top of file).

## Notable patterns (as used here)
- **Explicit vtable indexing**: faster but fragile; requires careful documentation and Windows version testing.
- **Fallback logic**: attempts to get volume via grouping parameter if direct `ISimpleAudioVolume` QI is unavailable.

## DIP investigation output
### Findings summary
(paste here)

### Recommended changes
(paste here)

### Risks / regressions to watch
(paste here)

## DIP investigation inquiry
You are investigating `SpotifyWASAPI.ahk`, a low-level COM/WASAPI module intended for per-process volume control without foreground activation. Key constraints: correctness across Windows versions, low latency, and not crashing/hanging the AHK process.

Please research and report on:
- **Modern / alternative methods** to control per-app audio:
  - Windows Core Audio APIs via other wrappers, PowerShell/.NET, or existing reliable libraries.
  - Whether newer Windows APIs offer a simpler/safer approach.
- **Code efficiency and robustness**:
  - Recommended approach to implement `AdjustProcessVolumeByPid` reliably (enumerating sessions, caching, revalidation).
  - Error handling patterns for COM failures and session churn.
- **Parallelism / async feasibility in AutoHotkey**:
  - Whether session enumeration should be cached/refreshed by a worker process to keep hotkeys instant.
- **Background macro feasibility without interrupting typing**:
  - Per-process volume control should be possible without activation—confirm constraints and best practices.

