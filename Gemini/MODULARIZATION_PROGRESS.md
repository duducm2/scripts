# Gemini.ahk modularization progress

Complete (2026-06-21). Same include-based recipe as
[`WindowManagement/MODULARIZATION_RECIPE.md`](../WindowManagement/MODULARIZATION_RECIPE.md).

## Result

| Metric                                     | Before       | After                       |
| ------------------------------------------ | ------------ | --------------------------- |
| Orchestrator [`Gemini.ahk`](../Gemini.ahk) | ~2,186 lines | **54 lines**                |
| Modules in [`Gemini/`](../Gemini/)         | 0            | **11 files**                |
| `/validate Gemini.ahk`                     | pass         | pass                        |
| Code-line multiset equivalence             | —            | **identical** (1,809 lines) |

Baseline rollback ref: `5eeb9a31` (pre-modularization `Gemini.ahk`)

## Still in orchestrator

- Preamble: `#Requires`, `#SingleInstance`, UIA includes, `env.ahk`, `Utils.ahk`
- Duplicate hotkey disable (`#!+Y`, `#!+X`) owned by Shift keys
- `aux/WMIPC.ahk`, `aux/GeminiIPC.ahk`
- MODULE MAP comment block and ordered `#include Gemini\...` pointers
- Note: `Gemini_FocusPromptSameAsOpenHotkey` lives in `Utils.ahk` (shared with Shift keys)

## Modules (11)

| Module                       | ~Lines | Feature                                                  |
| ---------------------------- | ------ | -------------------------------------------------------- |
| `gemini_async_readaloud.ahk` | 412    | `GeminiAsyncReadAloud` async read-aloud class            |
| `config_constants.ahk`       | 333    | Threshold constants, feature flags, early helpers        |
| `gemini_uia_core.ahk`        | 242    | `GeminiState`, copy-button UIA, tab banner, model picker |
| `gemini_delayed_submit.ahk`  | 222    | `GeminiDelayedSubmitMonitor` and start/stop helpers      |
| `background_helpers.ahk`     | 200    | `ShowNotification`, background timers, copy chime        |
| `gemini_async_lookup.ahk`    | 199    | `GeminiAsyncLookup` pronunciation async class            |
| `hotkey_read_copy.ahk`       | 208    | `#!+O` / `#!+P` / `#!+7` and copy helper                 |
| `gemini_async_tts.ahk`       | 133    | `GeminiAsyncTTS` class and `GeminiQueueBackgroundTask`   |
| `gemini_open.ahk`            | 135    | `InitializeGeminiFirstTime` and `#!+I` open/focus        |
| `loading_wait.ahk`           | 82     | Small loading indicator and `WaitForButton` helpers      |
| `hotkey_pronunciation.ahk`   | 71     | Language picker and `#!+8` pronunciation hotkey          |

Load order in orchestrator: config → uia_core → background → loading → read_copy → pronunciation → open → async_readaloud → async_lookup → delayed_submit → async_tts.

## Per-step workflow

1. Pick contiguous block; cut verbatim to `Gemini/<name>.ahk`
2. Replace with `#include %A_ScriptDir%\Gemini\<name>.ahk` at same position
3. Validate (syntax only — do not launch `#SingleInstance` scripts)
4. One commit: only `Gemini.ahk` + new module

Extract **bottom-up** so line numbers above the cut stay stable.

## Validate commands (Windows)

```powershell
$wd = "C:\Users\eduev\Meu Drive\17 - Projects\scripts"
$ahk = "$env:LOCALAPPDATA\Programs\AutoHotkey\v2\AutoHotkey64.exe"
$err = New-TemporaryFile
$p = Start-Process -FilePath $ahk `
  -ArgumentList '/ErrorStdOut /validate "Gemini.ahk"' `
  -WorkingDirectory $wd -Wait -PassThru -NoNewWindow `
  -RedirectStandardError $err.FullName
# exit 0 required
```

## Consumer validation checklist

Cross-validated every 4 modules and at completion:

- [x] `Gemini.ahk`
- [x] `Utils.ahk`
- [x] `Shift keys.ahk`
- [x] `AppLaunchers.ahk`

Leaf process: launched by [`Act.ahk`](../Act.ahk) via `Run GetScriptPath("Gemini.ahk")` — not `#include`d by other scripts.

## Equivalence check

Compare multiset of trimmed non-comment code lines (orchestrator + all `Gemini/*.ahk`) against
`git show 5eeb9a31:Gemini.ahk`. Split on `\n`; ignore new `#include Gemini\...` pointer lines.
Capture baseline blob with `Start-Process git show ... -RedirectStandardOutput` and read as UTF-8.

## Commits (12)

1. `9592e25` — scaffold `Gemini/` folder and MODULE MAP
2. `41691dc` … `b3ea2b9` — 11 bottom-up module extractions (async_tts through config_constants)
