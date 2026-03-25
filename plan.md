---
name: Block Handy Escape Key
overview: Update the global Escape hook in Utils.ahk to explicitly suppress the Escape key when Handy's "Recording" or "Recording Overlay" windows are active.
todos:
  - id: update_escape_condition
    content: In `Utils.ahk`, locate the `Escape::` hotkey under the "Block Escape Key from handy.exe" section. Update the `if (g_DictationActive)` condition to `if (g_DictationActive || WinActive("Recording ahk_exe handy.exe") || WinActive("Recording Overlay ahk_exe handy.exe"))`.
    status: pending
---

# Block Handy Escape Key

## Analysis / Context
The user wants to prevent the `Escape` key from accidentally closing the dictation software "Handy". Specifically, the block must be active whenever `handy.exe` has a window titled "Recording" or "Recording Overlay" in focus. Currently, `Utils.ahk` implements a global `Escape::` hook with `#InputLevel 10` that blocks `Escape` based solely on the `g_DictationActive` boolean state flag. To ensure complete robustness against state desynchronization and to meet the exact window-title requirement, we need to explicitly check for these windows.

## Proposed Changes
Update the existing `Escape::` hotkey in `Utils.ahk` to evaluate `WinActive()` for the specified Handy window titles. Since the hotkey already has `#InputLevel 10` and `#UseHook`, this change will automatically intercept and suppress `Escape` events originating from both physical hardware presses and other macros.

## Files to Modify
- `c:\Users\eduev\Meu Drive\17 - Projects\scripts\Utils.ahk` (around lines 2486-2525, inside the `Escape::` hotkey)

## Implementation Strategy
1. **Locate the Target Hook:** Find the `Escape::` definition in `Utils.ahk` under the `Block Escape Key from handy.exe` section.
2. **Expand the Blocking Condition:**
   - Locate the comment `; Block Escape if dictation is active (state-based, no timeout)` and the corresponding `if (g_DictationActive)` block.
   - Modify the `if` statement to include logical OR checks (`||`) for the required window titles: `WinActive("Recording ahk_exe handy.exe")` and `WinActive("Recording Overlay ahk_exe handy.exe")`.
3. **Preserve Fallback Logic:** Ensure the `return` statement inside this block remains untouched, so the script correctly consumes the keypress without passing it down to the system or `handy.exe`.
""