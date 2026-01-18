---
name: Focus Specific Cursor Window Macro
overview: Implement a macro triggered by 'H' in the Hotstring Selector to list open Cursor windows, activate the selected one, and close others.
todos:
  - id: define_globals
    content: Define global variables for Cursor selector (GUI, active flag, window map, hotkey handlers) in `utils.ahk`.
    status: pending
    dependencies: []
  - id: implement_cleanup
    content: Create `CleanupCursorWindowSelector()` function to destroy GUI and disable hotkeys.
    status: pending
    dependencies: [define_globals]
  - id: implement_selection_handler
    content: Create `HandleCursorWindowSelection(targetHwnd, allCursorWindows)` to activate target and close others.
    status: pending
    dependencies: [implement_cleanup]
  - id: implement_show_selector
    content: Create `ShowCursorWindowSelector()` to list Cursor windows, build GUI, and setup hotkeys using `g_HotstringCharSequence`.
    status: pending
    dependencies: [implement_selection_handler]
  - id: register_macro
    content: Register the new macro in `InitMacros()` with trigger 'h'.
    status: pending
    dependencies: [implement_show_selector]
---

# Focus Specific Cursor Window Macro

## Analysis / Context
The user requires a workflow to manage multiple "Cursor" application windows. The goal is to quickly select one specific window to keep focused while automatically closing all other open Cursor windows. This functionality should be accessible via the existing `Alt+Shift+U` (Hotstring Selector) menu, specifically bound to the 'H' key. The interaction model should be modal, keyboard-centric (single keypress), and consistent with existing selector patterns in `utils.ahk`.

## Proposed Changes
1.  **New Macro Registration**: Add a new entry to the `InitMacros()` function in `utils.ahk` that maps the 'h' key to a new function `ShowCursorWindowSelector`.
2.  **Selector Logic**: Implement `ShowCursorWindowSelector` which will:
    *   Retrieve all open windows for `ahk_exe Cursor.exe`.
    *   Display them in a modal GUI (reusing styles from `ShowHotstringSelector` if possible or creating a similar one).
    *   Assign a unique character to each window from `g_HotstringCharSequence`.
3.  **Action Logic**: Implement `HandleCursorWindowSelection` which will:
    *   Receive the selected window handle (HWND) and the list of all Cursor windows.
    *   Activate the selected window.
    *   Close all other windows in the list.
    *   Close the selector GUI.

## Files to Modify
*   `c:\Users\eduev\Meu Drive\12 - Scripts\utils.ahk`

## Implementation Strategy
1.  **Define Globals**: Add `g_CursorSelectorGui`, `g_CursorSelectorActive`, and `g_CursorSelectorHotkeyHandlers` to manage state.
2.  **Cleanup Function**: Implement `CleanupCursorWindowSelector` to destroy the GUI and turn off dynamic hotkeys (Escape and selection keys).
3.  **Selection Handler**: Create `HandleCursorWindowSelection(targetHwnd, allWindows)`:
    *   Iterate through `allWindows`.
    *   If a window's HWND does not match `targetHwnd`, call `WinClose`.
    *   Call `WinActivate` for `targetHwnd`.
    *   Call `CleanupCursorWindowSelector`.
4.  **Show Function**: Create `ShowCursorWindowSelector`:
    *   Call `CleanupHotstringSelector()` (if called from within that menu) or ensure clean state.
    *   Use `WinGetList("ahk_exe Cursor.exe")` to get windows.
    *   If no windows found, notify user.
    *   Iterate windows, assigning characters from `g_HotstringCharSequence`.
    *   Build a GUI displaying "[Char] > Window Title".
    *   Dynamically register hotkeys for the assigned characters pointing to `HandleCursorWindowSelection`.
    *   Register `Escape` to `CleanupCursorWindowSelector`.
5.  **Registration**: Update `InitMacros` to include `RegisterMacro(ShowCursorWindowSelector, "🪟 Focus Specific Cursor Window", "h")`.