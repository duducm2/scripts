# Shift keys.ahk — revision notes

## Purpose
Implements a large **global symbol layer / macro layer** centered around **Win+Alt+Shift** (and other modifier chords). It provides:
- System-wide shortcut remaps and macros.
- UI overlays and cheat sheets for discoverability.
- Automation flows that integrate with Chrome/web apps via UIA and with local tooling via IPC.

This file is effectively the “**keymap hub**” of the automation suite.

## Runtime model / methodology
- **AutoHotkey v2**, `#SingleInstance Force`, global `SetTitleMatchMode 2`.
- Heavy includes:
  - `env.ahk`
  - `UIA-v2` libs
  - `Utils.ahk`
  - `aux\ShiftKeysIPC.ahk` (IPC/bootstrap)
  - `CheatSheetRich.ahk` (RichEdit cheat sheet rendering)
- Uses timers for delayed post-start actions (e.g., applying audio session volume targets).
- Uses optional debug logging with file-lock-safe retries (`SafeDebugLog`) guarded by `DEBUG_SHIFTKEYS`.

## Key entry points
- **Auto-execute section**:
  - Initializes debug flags and helper functions.
  - Defines configuration such as `PROMPT_FILE`.
  - Boots IPC: `ShiftKeysIPC_Bootstrap()` (non-blocking).
  - Provides many helper functions for cheat sheet filtering/search and overlay focus management.
- **Hotkeys (representative, from scan)**:
  - `#!+a`, `#!+/`, `#!+1`, `#!+j`, `#!+l` (Win+Alt+Shift layer).
  - Additional chords include:
    - Copy-related hooks: `~^c`, `~PrintScreen`, `~!PrintScreen`, `$#+s`
    - Many Shift-only and Alt/Ctrl variants (large list; see file hotkey table).
  - The script acts as a single place to coordinate “free vs used” key combinations.

## Internal structure
- **Cheat sheet system**:
  - Uses padded bracket mnemonics (`PadShortcut`) and parsing/search helpers:
    - `CheatSheet_LineSearchHaystack`
    - `CheatSheet_LineMatchesQuery`
    - `CheatSheet_BuildFilteredBodyWithSections`
  - Focus management to prevent the body update from stealing focus from search input:
    - `CheatSheet_FocusSearchEdit`
    - `CheatSheet_DeferFocusSearch`
- **Overlay lifecycle**:
  - Escape handling splits into “app” vs “global” overlays:
    - `CheatSheet_OnEscapeApp`
    - `CheatSheet_OnEscapeGlobal`
- **IPC and safety**:
  - Integrates `ShiftKeysIPC` for daemon-like features.
  - Debug writes are gated and resilient to file locking to avoid hotpath slowdowns.

## Dependencies
- `UIA-v2` (web/app UI automation)
- `Utils.ahk` (shared overlays, helpers, media, UIA utilities)
- `aux\ShiftKeysIPC.ahk` (IPC)
- `CheatSheetRich.ahk` (RichEdit display)
- Various repo assets (e.g., prompt files, sounds).

## Notable patterns (as used here)
- **Centralized hotkey ownership**: one script to avoid conflicting global bindings.
- **Discoverability**: cheat sheet overlay is treated as a first-class feature.
- **Defensive debug**: optional, gated, and avoids breaking runtime responsiveness.

## DIP investigation output
### Findings summary
(paste here)

### Recommended changes
(paste here)

### Risks / regressions to watch
(paste here)

## DIP investigation inquiry
You are investigating `Shift keys.ahk`, a large global hotkey/macro layer that spans many workflows (including complex UI automation). Key constraints: latency (hotkeys must feel instant), safety (avoid accidental input leakage into the wrong window), and feasibility of advanced asynchronous/background macros that do not interrupt typing.

Please research and report on:
- **Modern / alternative methods**:
  - Alternatives to AHK for global key layers (e.g., Kanata, Interception, PowerToys, custom keyboard firmware layers) and trade-offs.
  - For web workflows: extensions or daemon-based automation that avoid window activation.
- **Code efficiency optimizations**:
  - Organizing the script into smaller modules to reduce load time and improve maintainability.
  - Techniques to reduce UIA traversal and repeated work in hot paths.
  - Better logging strategies that preserve responsiveness (ring buffers, async flush).
- **Parallelism / async feasibility in AutoHotkey**:
  - Multi-process worker/daemon patterns for heavy work while keeping hotkey handlers minimal.
  - Queues and state machines to serialize UI operations safely.
- **Complex background macros without interrupting typing**:
  - Feasibility of a macro like “Win+Alt+Shift+A moves mouse, pastes into Gemini, and returns focus” running asynchronously while the user continues typing elsewhere.
  - Whether UIA Invoke/background messaging can replace focus-stealing input injection.
  - Practical architecture recommendation for “background macro queue” with a non-intrusive confirmation UX.

