# Shift keys.ahk modularization progress

Phase 1 (59 modules) + phase 2 (orchestrator glue) complete (2026-06-21).
Recipe: [`WindowManagement/MODULARIZATION_RECIPE.md`](../WindowManagement/MODULARIZATION_RECIPE.md).

## Result

| Metric                                             | Phase 1            | After phase 2 (glue)         |
| -------------------------------------------------- | ------------------ | ---------------------------- |
| Orchestrator [`Shift keys.ahk`](../Shift keys.ahk) | ~976 lines         | **314 lines**                |
| Modules in [`Shift keys/`](../Shift keys/)         | 59                 | **65 files**                 |
| `/validate`                                        | pass               | pass                         |
| Code-line multiset equivalence                     | identical (21,151) | **identical** (21,151 lines) |

Phase-2 baseline rollback ref: `aa53439` (orchestrator + 59 modules before glue extraction)

## Phase 2 modules (6 new)

| Module                           | ~Lines | Feature                                    |
| -------------------------------- | ------ | ------------------------------------------ |
| `predicates_mercado_livre.ahk`   | 210    | `IsMercadoLivreActive`, `ML_*` UIA helpers |
| `predicates_shopee.ahk`          | 175    | `IsShopeeActive`, `Shopee_*` UIA helpers   |
| `uia_tree_inspector_helpers.ahk` | 129    | `UIATreeInspector_*` focus/jiggle helpers  |
| `predicates_file_dialog.ahk`     | 77     | `IsFileDialogActive` predicate             |
| `vscode_commit_message.ahk`      | 72     | `VSCode_TriggerGenerateCommitMessage`      |
| `predicates_chrome_pdf.ahk`      | 44     | `IsChromePdfViewerActive` predicate        |

## Still inline in orchestrator

- Header comment block, preamble, external includes
- `FocusBlackoutWatcher_Start()`, volume `SetTimer`, `ShiftKeysIPC`, `CheatSheetRich`
- MODULE MAP comment block
- `global g_WikipediaScrollHistory := []` (one line between includes)
- `#HotIf` context resets between hotif modules
- VS Code evidence bootstrap at file end: `EVIDENCE_SEARCH_FROM_SHIFT_KEYS`, `#include VSCodeEvidenceSearch.ahk`, `EvidenceSearch_BindHotkey()`

## Phase 1 modules (59)

Largest modules (good targets for focused AI edits):

- `hotif_outlook_reminder.ahk` (~1,442 lines)
- `hotif_editor_02.ahk` (~1,409)
- `hotif_powerbi.ahk` (~1,289)
- `cursor_predicates.ahk` (~1,264)
- `hotif_scroll_ai.ahk` (~1,089)
- `cheat_sheet_registry.ahk` (~1,251) — canonical `cheatSheets` map + `GLOBAL_CHEAT_SHEET_RAW` (replaces `app_hotkeys.ahk` and `cheat_sheet_data.ahk`)
- `gemini_chrome_02.ahk` (~941)
- `outlook_helpers_02.ahk` (~938)

Full list: all `Shift keys/*.ahk` except the six phase-2 modules above.

## Per-step workflow

1. Pick contiguous block; cut verbatim to `Shift keys/<name>.ahk`
2. Replace with `#include %A_ScriptDir%\Shift keys\<name>.ahk` at same position
3. Validate (syntax only — do not launch `#SingleInstance` scripts)
4. One commit: only `Shift keys.ahk` + new module

Extract **bottom-up** so line numbers above the cut stay stable.

## Validate commands (Windows)

```powershell
$wd = "C:\Users\eduev\Meu Drive\17 - Projects\scripts"
$ahk = "$env:LOCALAPPDATA\Programs\AutoHotkey\v2\AutoHotkey64.exe"
$err = New-TemporaryFile
$p = Start-Process -FilePath $ahk `
  -ArgumentList '/ErrorStdOut /validate "Shift keys.ahk"' `
  -WorkingDirectory $wd -Wait -PassThru -NoNewWindow `
  -RedirectStandardError $err.FullName
# exit 0 required
```

## Consumer validation checklist

Cross-validated every 3 modules and at completion:

- [x] `Shift keys.ahk`
- [x] `Utils.ahk`
- [x] `AppLaunchers.ahk`
- [x] `Gemini.ahk`
- [x] `WindowManagement.ahk`

Leaf process: launched by [`Act.ahk`](../Act.ahk) via `Run GetScriptPath("Shift keys.ahk")`.

## Equivalence check

Compare multiset of trimmed non-comment code lines (orchestrator + all `Shift keys/*.ahk`) against
the fileset at `aa53439` (59 modules + inline-glue orchestrator). Ignore `#include Shift keys\...` pointer lines.
Use `Start-Process git show "rev:Shift keys.ahk" -RedirectStandardOutput` (quoted path) + UTF-8 file read for each blob.

## Phase 2 commits (6)

1. `ad1d1b6` — `uia_tree_inspector_helpers.ahk`
2. `748221c` — `predicates_file_dialog.ahk`
3. `d64766e` — `vscode_commit_message.ahk`
4. `bfcdfd6` — `predicates_shopee.ahk`
5. `e4646ab` — `predicates_mercado_livre.ahk`
6. `64c2ed7` — `predicates_chrome_pdf.ahk`
