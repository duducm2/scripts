# WindowManagement.ahk modularization progress

Phase 1 (pilot) + phase 2 (inline blocks) complete (2026-06-21).
Recipe: [`MODULARIZATION_RECIPE.md`](MODULARIZATION_RECIPE.md).

## Result

| Metric                                                         | Pilot only   | After phase 2                                      |
| -------------------------------------------------------------- | ------------ | -------------------------------------------------- |
| Orchestrator [`WindowManagement.ahk`](../WindowManagement.ahk) | ~4,198 lines | **~210 lines** (+ inline summary docs)             |
| Modules in [`WindowManagement/`](../WindowManagement/)         | 6            | **13 files**                                       |
| `/validate`                                                    | pass         | pass                                               |
| Code-line count (orchestrator + all modules)                   | —            | **4,625 lines** (unchanged vs pre-phase-2 fileset) |

Phase-2 baseline rollback ref: `c9d4ba714bef4f8c0915db6f1be466f9cdb62272`

## Modules (14)

| Module                     | ~Lines | Feature                                           |
| -------------------------- | ------ | ------------------------------------------------- |
| `tile_snap.ahk`            | 929    | Background tile, snap half-pair, maximize helpers |
| `minimized_list.ahk`       | 893    | Hidden background window list GUI                 |
| `window_cycle.ahk`         | 836    | Cycle/minimize/close windows on monitor           |
| `cursor_window_select.ahk` | 837    | Cursor window selection in project selector       |
| `cursor_composer.ahk`      | 700+   | Cursor AI composer focus (UIA)                    |
| `project_selector_01.ahk`  | 624    | Project quick selector GUI (`#!+L`), part 1       |
| `background_scan.ahk`      | 472    | Background scan, title excludes, collection       |
| `project_selector_02.ahk`  | 491    | Selection mode and preview handlers, part 2       |
| `move_monitor.ahk`         | 334    | Move to ordered monitor; MEH+B/C Alt+Tab          |
| `window_tools.ahk`         | 258    | Window tools menu (no hotkey; CAW for [1]-[4])    |
| `hotkeys.ahk`              | 84     | Global hotkey label bindings                      |
| `globals.ahk`              | 57     | Globals + startup timers                          |
| `helpers.ahk`              | 57     | Notifications, activation helpers                 |
| `audio_bt_menu.ahk`        | ~480   | Win+Alt+Shift+9 Bluetooth / audio device menu     |

## Still inline in orchestrator

- Preamble, WM daemon flags, `_DebugLog_WM`, external includes (`env`, `Utils`, `WMIPC`, …)
- MODULE MAP and end-of-file summary/optimization comment block (~115 lines)

## Validate command (Windows)

```powershell
$wd = "C:\Users\eduev\Meu Drive\17 - Projects\scripts"
$ahk = "$env:LOCALAPPDATA\Programs\AutoHotkey\v2\AutoHotkey64.exe"
Start-Process -FilePath $ahk -ArgumentList '/ErrorStdOut /validate "WindowManagement.ahk"' `
  -WorkingDirectory $wd -Wait -NoNewWindow
```

## Rollout status (repo-wide)

| Script                 | Status                                   |
| ---------------------- | ---------------------------------------- |
| `WindowManagement.ahk` | Done (13 modules)                        |
| `Shift keys.ahk`       | Done (65 modules; glue phase 2 complete) |
| `Utils.ahk`            | Done (51 modules)                        |
| `Gemini.ahk`           | Done (11 modules)                        |
| `AppLaunchers.ahk`     | Done (12 modules)                        |

Repo-wide include modularization rollout is **complete** (2026-06-21).
