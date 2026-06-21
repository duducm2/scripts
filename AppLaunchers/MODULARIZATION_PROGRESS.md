# AppLaunchers.ahk modularization progress

Complete (2026-06-21). Same include-based recipe as
[`WindowManagement/MODULARIZATION_RECIPE.md`](../WindowManagement/MODULARIZATION_RECIPE.md).

## Result

| Metric                                                 | Before       | After                       |
| ------------------------------------------------------ | ------------ | --------------------------- |
| Orchestrator [`AppLaunchers.ahk`](../AppLaunchers.ahk) | ~2,261 lines | **83 lines**                |
| Modules in [`AppLaunchers/`](../AppLaunchers/)         | 0            | **12 files**                |
| `/validate AppLaunchers.ahk`                           | pass         | pass                        |
| Code-line multiset equivalence                         | —            | **identical** (1,650 lines) |

Baseline rollback ref: `39a8210^` (pre-modularization `AppLaunchers.ahk`)

## Still in orchestrator

- Preamble: `#Requires`, `#SingleInstance`, IPC feature flags (`AL_USE_*`)
- External `#include`s: `env.ahk`, `infra/ipc/AppLauncherIPC.ahk`, UIA, `Utils.ahk`
- `OnExit(AL_AppLaunchersExit)` handler (calls module helpers for teardown)
- `AL_DESKTOP_CACHE` UIA cache creation (must run after UIA include)
- Duplicate hotkey disable (`#!+Y`, `#!+X`) owned by Utils/Shift keys
- `/Updated` volume schedule guard
- MODULE MAP comment block and ordered `#include AppLaunchers\...` pointers

## Modules (12)

| Module                       | ~Lines | Feature                                          |
| ---------------------------- | ------ | ------------------------------------------------ |
| `wikipedia_selector.ahk`     | 532    | Wikipedia selector GUI and char handlers         |
| `pomodoro_timer.ahk`         | 458    | Pomodoro timer system with CSV logging           |
| `desktop_explorer.ahk`       | 360    | Shift+Win+E desktop explorer and UIA helpers     |
| `wikipedia_scroll.ahk`       | 300    | Wikipedia scroll position save/load/restore      |
| `wikipedia_focus_guard.ahk`  | 206    | Wikipedia focus monitor and input guard          |
| `wikipedia_entry.ahk`        | 165    | `SelectWikipediaInHandy` and `#!+k` hotkey       |
| `launch_hotkeys.ahk`         | 157    | Chrome, WhatsApp, YouTube, Cursor launch hotkeys |
| `wikipedia_globals.ahk`      | 39     | Wikipedia selector globals and article list      |
| `hotkey_clipangel.ahk`       | 28     | `#!+.` Clip Angel paste and favorite flow        |
| `config_globals.ahk`         | 19     | Phase 3/4 global state for hooks and wiki FSM    |
| `center_mouse.ahk`           | 15     | `CenterMouse` helper on active window            |
| `hotkey_context_browser.ahk` | 13     | `#!+n` context file browser hotkey               |

Load order: config_globals → hotkey_context_browser → desktop_explorer → launch_hotkeys → wikipedia_globals → wikipedia_focus_guard → wikipedia_scroll → wikipedia_selector → wikipedia_entry → pomodoro_timer → center_mouse → hotkey_clipangel.

## Per-step workflow

1. Pick contiguous block; cut verbatim to `AppLaunchers/<name>.ahk`
2. Replace with `#include %A_ScriptDir%\AppLaunchers\<name>.ahk` at same position
3. Validate (syntax only — do not launch `#SingleInstance` scripts)
4. One commit: only `AppLaunchers.ahk` + new module

Extract **bottom-up** so line numbers above the cut stay stable.

## Validate commands (Windows)

```powershell
$wd = "C:\Users\eduev\Meu Drive\17 - Projects\scripts"
$ahk = "$env:LOCALAPPDATA\Programs\AutoHotkey\v2\AutoHotkey64.exe"
$err = New-TemporaryFile
$p = Start-Process -FilePath $ahk `
  -ArgumentList '/ErrorStdOut /validate "AppLaunchers.ahk"' `
  -WorkingDirectory $wd -Wait -PassThru -NoNewWindow `
  -RedirectStandardError $err.FullName
# exit 0 required
```

## Consumer validation checklist

Cross-validated every 4 modules and at completion:

- [x] `AppLaunchers.ahk`
- [x] `Utils.ahk`
- [x] `Shift keys.ahk`
- [x] `Gemini.ahk`

Leaf process: launched by [`Act.ahk`](../Act.ahk) via `Run GetScriptPath("AppLaunchers.ahk")`. Must stay **last** in Quick Update relaunch order (includes `Utils.ahk`; `/Updated` arg triggers success overlay).

## Equivalence check

Compare multiset of trimmed non-comment code lines (orchestrator + all `AppLaunchers/*.ahk`) against
`git show 39a8210^:AppLaunchers.ahk`. Split on `\n`; ignore new `#include AppLaunchers\...` pointer lines.
Capture baseline blob with `Start-Process git show ... -RedirectStandardOutput` and read as UTF-8.

## Commits (13)

1. `39a8210` — scaffold `AppLaunchers/` folder and MODULE MAP
2. `6d5d1a4` … `aa53439` — 12 bottom-up module extractions (hotkey_clipangel through config_globals)
