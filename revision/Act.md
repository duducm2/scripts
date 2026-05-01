# Act.ahk — revision notes

## Purpose
Bootstrapper that updates repos (scripts + notes) and launches the main automation suite appropriate for the current environment (personal vs work).

## Runtime model / methodology
- **AutoHotkey v2** script acting primarily as an **orchestrator** (not hotkey-heavy).
- **Environment-aware branching** via `env.ahk` (e.g., `IS_WORK_ENVIRONMENT`).
- Uses **bounded waits** and visible status via `StandardLoadingBar_*` helpers.
- Uses **external processes** for updates (`git fetch`, `git pull`) via `RunWait`.

## Key entry points
- **Auto-execute section**:
  - Includes `env.ahk` and `Utils.ahk`.
  - In work environment, prompts the user with `MsgBox` gate before proceeding.
  - Pulls updates for:
    - `scriptsFolder` repo
    - `notesFolder` repo
  - Launches the primary scripts via `Run GetScriptPath("<name>.ahk")`, including:
    - `Shift keys.ahk`
    - `Gemini.ahk`
    - `AppLaunchers.ahk`
    - `WindowManagement.ahk`
    - `Mousemaster.ahk`
    - plus work-only scripts (`Microsoft Teams.ahk`, `Outlook.ahk`)

## Structure / dependencies
- **Depends on**:
  - `env.ahk` for environment detection and path selection.
  - `Utils.ahk` for `StandardLoadingBar_*`, `GetScriptPath`, and UX helpers.
- **Assumes**:
  - `git` is available in PATH.
  - The repos are clean enough for `git pull` to succeed unattended.

## Notable patterns (as used here)
- **Operational safety**: prompts on work environment before automation begins.
- **Human feedback**: persistent progress overlay during long operations.
- **Separation of concerns**: Act only launches; other scripts own hotkeys and automation logic.

## DIP investigation output
### Findings summary
(paste here)

### Recommended changes
(paste here)

### Risks / regressions to watch
(paste here)

## DIP investigation inquiry
You are investigating `Act.ahk`, which is a startup/orchestration script that pulls updates and launches multiple AHK automation processes. The key constraints are reliability, quick startup, and not breaking the user’s session with unexpected focus steals or long blocking waits.

Please research and report on:
- **Modern / alternative web-based methods** to achieve the same objectives (e.g., replacing some parts of the suite with web/desktop automation stacks outside AHK, or using OS-native scheduled tasks + webhooks to update repositories).
- **Code efficiency optimizations**:
  - Faster, safer update strategy than `git fetch` + `git pull` in the UI thread.
  - Ways to reduce startup latency while preserving correctness and user visibility.
- **Parallelism / async feasibility in AutoHotkey**:
  - Best practices to parallelize updates or launches (multi-process patterns, timers, message passing, COM callbacks).
- **Background macro feasibility without interrupting typing**:
  - How to orchestrate “launch/update” steps while minimizing focus changes.
  - Approaches to keep Act’s operations non-intrusive (e.g., background `git` with log-only overlays, deferring focus-affecting steps).

