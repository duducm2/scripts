# AHK Modularization Recipe (include-based)

Reusable, low-risk recipe for splitting a large AutoHotkey v2 script into small
per-feature module files so low-context AI models can be fed one feature at a
time. Proven on `WindowManagement.ahk` (pilot); replay on `Shift keys.ahk`
(next) then `Utils.ahk` (last, highest blast radius).

## Core idea

`#include` is a compile-time textual insert. Replacing a contiguous block of a
script with `#include path\to\module.ahk` at the **same position** yields a
program that assembles to the exact same code. The main file stays the runnable
entry point / source of truth; only its physical layout changes. No slicer that
runs at runtime, no sync job, no read-only mirror to drift.

## AHK v2 facts that make this safe

- Function and hotkey definitions are registered globally regardless of which
  file/position they appear in. So function-only and hotkey-only blocks can be
  moved into includes freely.
- The auto-execute thread runs top-level statements in textual order and skips
  over definitions. Any block containing top-level side effects (global var
  init, `SetTimer`, `A_TrayMenu.Add`, GUI setup) must keep its `#include` at the
  original position so startup order is preserved.
- Moving (cutting) code avoids "Duplicate function definition" errors. Never
  leave a copy behind in the main file.
- Files here are UTF-8 (no BOM) with CRLF line endings. Preserve both when
  slicing (write with `System.Text.UTF8Encoding($false)` and join with CRLF).

## Per-module loop

1. Pick a contiguous, cohesive block (use the `; ===` / `; ---` banner sections
   as natural seams). Confirm exact start/end lines so you do not split a
   function or hotkey.
2. Cut the block to `Subfolder\<feature>.ahk` (verbatim, with a short header
   comment), and replace it in the main file with a one-line `#include` plus a
   pointer comment, at the same position.
3. Validate (syntax, no run):
   `AutoHotkey64.exe /ErrorStdOut /validate "Main.ahk"` and check exit code via
   `Start-Process -Wait -PassThru` (AHK is a GUI app, so `$LASTEXITCODE` is not
   set and `/ErrorStdOut` avoids a blocking error dialog).
4. Commit just the two touched files (`git add -- Main.ahk Subfolder\feat.ahk`).
   One commit per module = one-line revert if anything regresses.

Extract bottom-up (highest line numbers first) so earlier blocks keep their
original line numbers. Do side-effect blocks (globals/timers) last.

## Whole-fileset equivalence check

After all extractions, prove no code was lost/duplicated/altered by comparing
the multiset of real code lines (trim each line; drop blanks, `;` comments, and
the new `#include Subfolder\...` pointers) of the pre-refactor file against the
union of the orchestrator + all modules. They must be identical.

Gotcha: `git show <ref>:file` emits the blob with LF endings (working tree is
CRLF). Split on `\n` and `Trim()` so the comparison is ending-agnostic. Capture
the blob byte-faithfully with `Start-Process git show ... -RedirectStandardOutput`
(PowerShell pipelines and `cmd >` can corrupt UTF-8 / line endings).

## Pilot result (WindowManagement.ahk)

Orchestrator 5701 -> ~210 lines (plus summary comments). **13 modules** in
`WindowManagement/` (helpers, globals, tile_snap, window_tools, background_scan,
minimized_list, hotkeys, move_monitor, window_cycle, project_selector_01/02,
cursor_composer, cursor_window_select). `/validate` passes. See
`WindowManagement/MODULARIZATION_PROGRESS.md`.

## Rollout notes

- `Shift keys.ahk` (~26k, leaf process): done (59 modules).
- `Utils.ahk` (~18.8k, shared library): done (51 modules); validate every consumer.
- Optional: `Gemini.ahk`, `AppLaunchers.ahk`; Shift keys orchestrator glue (~400 lines).
