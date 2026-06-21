# Shift keys.ahk modularization progress

Pilot complete (2026-06-21). Same include-based recipe as
[`WindowManagement/MODULARIZATION_RECIPE.md`](../WindowManagement/MODULARIZATION_RECIPE.md).

## Result

| Metric | Before | After |
|--------|--------|-------|
| Orchestrator [`Shift keys.ahk`](../Shift keys.ahk) | ~26,292 lines | **976 lines** |
| Modules in [`Shift keys/`](../Shift keys/) | 0 | **59 files** |
| `/validate` | pass | pass |
| Code-line multiset equivalence | — | **identical** (21,151 lines) |

## Modules (59)

Largest modules (good targets for focused AI edits):

- `hotif_outlook_reminder.ahk` (~1,442 lines)
- `hotif_editor_02.ahk` (~1,409)
- `hotif_powerbi.ahk` (~1,289)
- `cursor_predicates.ahk` (~1,264)
- `hotif_scroll_ai.ahk` (~1,089)
- `app_hotkeys.ahk` (~1,092)
- `gemini_chrome_02.ahk` (~941)
- `outlook_helpers_02.ahk` (~938)

Full list: all `Shift keys/*.ahk` files (helpers through m365_copilot_temp).

## Still inline in orchestrator (~400 lines)

Predicate/helper glue left between `#include` lines (not worth splitting further unless
you want a sub-200-line orchestrator):

- `IsChromePdfViewerActive`, Mercado Livre / Shopee detection + UIA helpers
- `IsFileDialogActive`, `UIATreeInspector_*` helpers
- `VSCode_TriggerGenerateCommitMessage`
- Preamble: `FocusBlackoutWatcher_Start`, volume `SetTimer`, external includes
- VSCode evidence bootstrap (`VSCodeEvidenceSearch.ahk`)

## Per-step workflow (for next extractions)

1. Pick contiguous block; cut verbatim to `Shift keys/<name>.ahk`
2. Replace with `#include %A_ScriptDir%\Shift keys\<name>.ahk` at same position
3. `AutoHotkey64.exe /ErrorStdOut /validate "Shift keys.ahk"` (use `WorkingDirectory` = scripts folder)
4. One commit: only `Shift keys.ahk` + new module

Extract **bottom-up** so line numbers above the cut stay stable.

## Next rollout

- Optional: peel remaining orchestrator glue into `Shift keys/predicates_*.ahk`
- Then: `Utils.ahk` (~18.8k) — validate every consumer after each module

## Validate command (Windows)

```powershell
$wd = "C:\Users\eduev\Meu Drive\17 - Projects\scripts"
$ahk = "$env:LOCALAPPDATA\Programs\AutoHotkey\v2\AutoHotkey64.exe"
Start-Process -FilePath $ahk -ArgumentList '/ErrorStdOut /validate "Shift keys.ahk"' `
  -WorkingDirectory $wd -Wait -NoNewWindow
```

Do not launch the script for validation (`#SingleInstance Force` hijacks the live process).
