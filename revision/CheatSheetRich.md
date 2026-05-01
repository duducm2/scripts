# CheatSheetRich.ahk — revision notes

## Purpose
Provides a **RichEdit-based cheat sheet renderer** used by other scripts (notably `Shift keys.ahk`) to display shortcut lists with enhanced formatting (mnemonics in bold/larger font) reliably across theming/DPI.

## Runtime model / methodology
- Pure **library-style** module: defines functions and globals; no primary hotkeys.
- Works around RichEdit quirks by:
  - Loading `msftedit.dll` explicitly before creating RichEdit controls.
  - Using `EM_SETTEXTEX` (UTF-16) rather than relying on `WM_SETTEXT`.
  - Applying `CHARFORMAT2W` formatting spans with UTF-16 code unit indexing.
  - Disabling visual styles via `uxtheme\SetWindowTheme` to prevent “invisible” text in dark/themed contexts.

## Key entry points
- `CheatSheet_EnsureRichDll()`: loads `msftedit.dll` once.
- `CheatSheet_RichSetProcessedBody(ctrl, processedText)`: main “render pipeline”
  - Converts “processed” text into plain text
  - Tracks mnemonic spans
  - Writes text via `EM_SETTEXTEX`
  - Applies base + mnemonic formats via `EM_SETCHARFORMAT`
  - Ensures background color and readonly state are correct

## Internal structure
- **UTF-16 correctness**:
  - `CheatSheet_Utf16Units(s)` counts code units so character indices match RichEdit’s internal storage.
- **Selection APIs**:
  - Uses `EM_EXSETSEL` (CHARRANGE) for “select all” reliability.
- **Formatting**:
  - Builds `CHARFORMAT2W` buffers via `CheatSheet_RichCharFormat2(...)`.
- **Parsing**:
  - `CheatSheet_ParseProcessedLine` and helpers split lines into segments, identifying mnemonic spans (e.g., content inside `[...]`).

## Dependencies
- Win32 messages: `EM_SETTEXTEX`, `EM_SETCHARFORMAT`, `EM_SETREADONLY`, `EM_SETBKGNDCOLOR`, etc.
- `uxtheme.dll` for theming suppression.

## Notable patterns (as used here)
- **Low-level Win32 interop** used to produce deterministic UI behavior across hosts.
- **Defensive rendering**: multiple passes to ensure background/readonly don’t reset formatting.

## DIP investigation output
### Findings summary
(paste here)

### Recommended changes
(paste here)

### Risks / regressions to watch
(paste here)

## DIP investigation inquiry
You are investigating `CheatSheetRich.ahk`, a RichEdit rendering helper that needs to remain stable across Windows 10/11 theming, DPI, and different GUI hosts. Reliability is more important than micro-optimizations, but responsiveness matters during frequent updates (e.g., filtering/search).

Please research and report on:
- **Modern / alternative UI methods** to render a rich cheat sheet:
  - WebView2-based overlay, native UI frameworks, or alternative AHK controls.
  - Trade-offs vs RichEdit: latency, styling control, deployment complexity.
- **Code efficiency optimizations**:
  - Faster span formatting approaches (batch operations, minimizing message traffic).
  - Avoiding repeated full-document reformat when only small changes occur (incremental updates).
- **Parallelism / async feasibility in AutoHotkey**:
  - Whether formatting work should be offloaded (multi-process helper) vs kept synchronous.
- **Background macro feasibility without interrupting typing**:
  - How to show/update overlays without stealing focus (NoActivate windows, input routing).

