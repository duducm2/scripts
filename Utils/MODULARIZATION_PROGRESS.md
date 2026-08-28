# Utils.ahk modularization progress

Complete (2026-06-21). Same include-based recipe as
[`WindowManagement/MODULARIZATION_RECIPE.md`](../WindowManagement/MODULARIZATION_RECIPE.md).

## Result

| Metric                                   | Before        | After                        |
| ---------------------------------------- | ------------- | ---------------------------- |
| Orchestrator [`Utils.ahk`](../Utils.ahk) | ~18,792 lines | **223 lines**                |
| Modules in [`Utils/`](../Utils/)         | 0             | **51 files**                 |
| `/validate Utils.ahk`                    | pass          | pass                         |
| Code-line multiset equivalence           | —             | **identical** (15,218 lines) |

Baseline rollback ref: `5951bd39f6a834418f2f9d37f671394e12cdcd7d`

## Still in orchestrator

- Preamble: `#Requires`, `#SingleInstance`, external `#include`s, `g_StudyLinkSubmenuGui`, `UIA_ControlType_Button`
- `StartTimer` / `EndTimer` / `WindowExists` (before banner globals)
- Early `BANNER_ACCENT_*` and `GEMINI_PROMPT_*` globals (must load before modules)
- `#include %A_ScriptDir%\lib\CopilotWeb.ahk` (before `d2c_flow_manager` module)
- MODULE MAP comment block

## Modules (51)

| Module                               | ~Lines | Feature                                           |
| ------------------------------------ | ------ | ------------------------------------------------- |
| `context_file_browser.ahk`           | 1159   | Context file browser (Win+Alt+Shift+N)            |
| `standard_loading_bar.ahk`           | 810    | Standard loading bar lifecycle                    |
| `hotstring_selector_core.ahk`        | 784    | Hotstring selector core / `BuildHotstringCharMap` |
| `peek_pdf_study_03.ahk`              | 782    | Peek PDF / QuickLook study (part 3)               |
| `clip_angel_favorite.ahk`            | 772    | Clip Angel favorite flows                         |
| `chrome_detach_02.ahk`               | 743    | Chrome detach context menu (part 2)               |
| `gemini_cursor_transfer.ahk`         | 742    | Gemini-to-Cursor transfer selector                |
| `chrome_detach_03.ahk`               | 735    | Chrome detach entry (part 3)                      |
| `peek_pdf_study_02.ahk`              | 705    | Peek PDF / QuickLook study (part 2)               |
| `d2c_flow_manager.ahk`               | 664    | D2C_FlowManager state machine                     |
| `peek_pdf_study_01.ahk`              | 663    | Peek PDF / QuickLook study (part 1)               |
| `dictation_toggle.ahk`               | 663    | Dictation indicator + `~#!+0`                     |
| `hotstring_selector_gui.ahk`         | 662    | `ShowHotstringSelector` GUI                       |
| `square_selector_mouse_jump_02.ahk`  | 614    | Square selector mouse jump (part 2)               |
| `square_selector_mouse_jump_01.ahk`  | 595    | Square selector mouse jump (part 1)               |
| `hotstring_selector_handlers_02.ahk` | 572    | Utility category handlers                         |
| `hotstring_selector_handlers_01.ahk` | 567    | `HandleHotstringChar` + Gemini paste              |
| `cursor_composer_focus.ahk`          | 564    | Cursor / VS Code chat focus                       |
| `chrome_detach_01.ahk`               | 546    | Chrome detach helpers (part 1)                    |
| `gemini_mode_picker.ahk`             | 439    | Gemini mode picker UIA                            |
| `handy_uia_helpers.ahk`              | 431    | Handy UIA helpers                                 |
| `hotstrings_core.ahk`                | 410    | Hotstrings + `InitHotstringsCheatSheet()`         |
| `focus_mode.ahk`                     | 331    | Focus mode blackout `#!+Y`                        |
| `helpers_timing_gemini.ahk`          | 307    | `FindGeminiPromptField` and timing helpers        |
| `macros_system.ahk`                  | 285    | Macros + `RegisterMacro`                          |
| `language_flag_indicator.ahk`        | 271    | Language flag chips                               |
| `clip_angel_merge.ahk`               | 230    | Clip Angel merge clips                            |
| `outlook_teams_check.ahk`            | 214    | `CheckAndOpenOutlookTeams`                        |
| `hotstring_gemini_banner.ahk`        | 210    | Hotstring Gemini banner + D2C presets             |
| `f11_fullscreen.ahk`                 | 203    | F11 fullscreen helpers                            |
| `dictation_cleanup.ahk`              | 193    | Dictation clipboard cleanup countdown             |
| `toggle_outlook_teams.ahk`           | 187    | `ToggleOutlookAndTeams`                           |
| `dictation_merge.ahk`                | 179    | Dictation merge countdown                         |
| `mouse_jump_arrows.ahk`              | 176    | Mouse jump prediction helpers                     |
| `ai_generation_state.ahk`            | 176    | Global AI generation state macro                  |
| `handy_ai_model_gui.ahk`             | 161    | `ShowAiModelSelector` GUI                         |
| `global_sound_audio.ahk`             | 154    | Global sound toggle                               |
| `cleanclipboard.ahk`                 | 152    | Clean clipboard macro                             |
| `print_screen_escape.ahk`            | 135    | Print Screen chime + global Escape                |
| `desktop_recycle.ahk`                | 129    | Desktop to Recycle Bin                            |
| `utility_shortcuts.ahk`              | 126    | `#!+U`, `#!+W` Macros, and `^!#` triggers         |
| `hotstring_selector_cleanup.ahk`     | 105    | `CleanupHotstringSelector`                        |
| `dictation_legacy.ahk`               | 92     | Deprecated dictation Gemini confirm               |
| `handy_ai_model_config.ahk`          | 76     | Handy AI model map                                |
| `project_data_cursor.ahk`            | 73     | Project data for Cursor selector                  |
| `clip_angel_activate.ahk`            | 55     | Clip Angel activate                               |
| `files_links.ahk`                    | 49     | Quick-open files + `InitQuickOpenFiles()`         |
| `handy_selector_entry.ahk`           | 39     | `SelectAiModelInHandy` entry                      |
| `study_hotkey_x.ahk`                 | 37     | `#!+x` study hotkey                               |
| `jump_mouse_middle.ahk`              | 31     | `#!+Q` center mouse                               |
| `mouse_jump_hotkeys.ahk`             | 31     | `#!+Arrow` five-step                              |
| `handy_selector_hotkey.ahk`          | 14     | `#!+C` handy selector                             |

## Pack-import modules (not in table above — added after modularization)

Canonical agent documentation: [`docs/prompt-data-output-and-finance-packs.md`](../docs/prompt-data-output-and-finance-packs.md)

| Module                                                                                        | Feature                                    |
| --------------------------------------------------------------------------------------------- | ------------------------------------------ |
| `finance_helpers.ahk` / `finance_import.ahk` / `finance_launcher.ahk`                         | Finance daily/monthly pack import          |
| `mnemonic_palace_helpers.ahk` / `mnemonic_palace_import.ahk` / `mnemonic_palace_launcher.ahk` | Memory Palace + plan pack import           |
| `import_mgmt_helpers.ahk` / `import_mgmt_import.ahk` / `import_mgmt_launcher.ahk`             | Import Management / job search pack import |

Note: `context_file_browser.ahk` exceeds the 900-line target; split further only if edits become unwieldy.

## Per-step workflow

1. Pick contiguous block; cut verbatim to `Utils/<name>.ahk`
2. Replace with `#include %A_ScriptDir%\Utils\<name>.ahk` at same position
3. Validate (syntax only — do not launch `#SingleInstance` scripts)
4. One commit: only `Utils.ahk` + new module

Extract **bottom-up** so line numbers above the cut stay stable. Extract auto-execute blocks (`InitHotstringsCheatSheet`, `InitQuickOpenFiles`) **late**.

## Validate commands (Windows)

```powershell
$wd = "C:\Users\eduev\Meu Drive\17 - Projects\scripts"
$ahk = "$env:LOCALAPPDATA\Programs\AutoHotkey\v2\AutoHotkey64.exe"
$err = New-TemporaryFile
$p = Start-Process -FilePath $ahk `
  -ArgumentList '/ErrorStdOut /validate "Utils.ahk"' `
  -WorkingDirectory $wd -Wait -PassThru -NoNewWindow `
  -RedirectStandardError $err.FullName
# exit 0 required
```

## Consumer validation checklist

After every 5 modules (and at completion), validate:

- [x] `Utils.ahk`
- [x] `Shift keys.ahk`
- [x] `WindowManagement.ahk`
- [x] `AppLaunchers.ahk`
- [x] `Gemini.ahk`
- [x] `Act.ahk`
- [x] `Outlook.ahk`
- [x] `Microsoft Teams.ahk`
- [x] `TestStudyLinkApi.ahk`

## Equivalence check

Compare multiset of trimmed non-comment code lines (orchestrator + all `Utils/*.ahk`) against
`git show 5951bd39:Utils.ahk`. Split on `\n`; ignore new `#include Utils\...` pointer lines.
