# AGENTS.md

## Cursor Cloud specific instructions

### Overview

This is a **Windows-only AutoHotkey v2 desktop automation toolkit**. It provides keyboard shortcuts and automation scripts for personal productivity on Windows, automating interactions with Spotify, Microsoft Teams, Chrome, Outlook, Gemini AI, Cursor IDE, WhatsApp, and more.

### Platform constraint

All `.ahk` scripts require **Windows + AutoHotkey v2.0+** to run. They use Windows-specific APIs (UI Automation, DllCall, WinActivate, etc.) and **cannot execute on Linux**. There is no cross-platform runtime, no Docker container, and no build step.

### Architecture

| File | Role |
|---|---|
| `Act.ahk` | Entry point — pulls latest code, launches all other scripts |
| `Shift keys.ahk` | Main hotkey hub (15k lines) — global app shortcuts, UIA patterns, Spotify, Outlook |
| `Utils.ahk` | Shared utilities (8k lines) — cursor centering, composite actions, prompts, Pomodoro, Wikipedia |
| `Gemini.ahk` | Gemini AI integration — TTS, pronunciation, copy response, async workflow |
| `GeminiToCursorBridge.ahk` | Gemini-to-Cursor IDE bridge |
| `AppLaunchers.ahk` | Chrome-based app launchers (Gmail, WhatsApp, YouTube, etc.) |
| `WindowManagement.ahk` | Multi-monitor cycling, window move/maximize, cursor halo |
| `Microsoft Teams.ahk` | Teams meeting helpers, mic/camera state |
| `Outlook.ahk` | Outlook automation |
| `Spotify.ahk` | Spotify media control via UIA |
| `Lib/InfiniteDictation.ahk` | Speech-to-text workflow module |

### Key dependencies (not in repo)

- **`env.ahk`** — Gitignored config file defining `IS_WORK_ENVIRONMENT`, user paths, etc. Must be created manually per machine.
- **`UIA-v2/`** — UI Automation v2 library directory. Listed in `.gitignore` with `~` suffix; the directory exists but is empty. Must be populated manually with the [UIA-v2 library](https://github.com/Descolada/UIA-v2) files (`UIA.ahk`, `UIA_Browser.ahk`).

### Include graph

- `Act.ahk` → `#Include env.ahk`
- `Utils.ahk` → `#Include "Lib\InfiniteDictation.ahk"`

### What cloud agents can do

Since AHK scripts cannot run on Linux, cloud agents working on this codebase are limited to:

1. **Code editing** — Modify `.ahk` files, `.ps1` scripts, docs, and data files.
2. **File validation** — Verify all files are present and readable; check include references resolve.
3. **Documentation** — Update `README.md`, `docs/`, prompt templates, and data files.
4. **No lint/test/build/run** — There is no AHK linter that runs on Linux, no automated tests, no build step, and no way to execute the scripts.

### Conventions

- See `README.md` for coding patterns (UIA anchoring, tab navigation, state toggles).
- See `docs/cheat-sheet-standard.md` for hotkey documentation format.
- See `docs/asynchronous_workflow_standards.md` for async workflow architecture.
- Hotkey modifier order: Shift set first, then Ctrl+Alt+Shift (MEH) when Shift set is full.
- Prefer UIA over pixel/image matching; fail safe with native key fallbacks.
