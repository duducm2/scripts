---
name: Focus Specific Cursor Window Macro
overview: Implement a feature triggered by 'h' in the Project Selector (Win+Alt+Shift+L) to list open Cursor windows, activate the selected one, and close others.
todos:
  - id: define_globals
    content: Define global variables for Cursor window management (window map, hotkey handlers) in `WindowManagement.ahk`.
    status: pending
    dependencies: []
  - id: implement_selection_handler
    content: Create `HandleCursorWindowSelection(targetHwnd, allCursorWindows)` to activate target and close others.
    status: pending
    dependencies: [define_globals]
  - id: integrate_project_selector
    content: Integrate Cursor window listing into `ShowProjectSelector()` as a special item with character 'h', showing after projects and before preview windows.
    status: pending
    dependencies: [implement_selection_handler]
  - id: remove_macro_registration
    content: Remove the macro registration from `InitMacros()` in `Utils.ahk`.
    status: pending
    dependencies: [integrate_project_selector]
  - id: remove_cursor_selector_standalone
    content: Remove standalone Cursor window selector functions from `Utils.ahk` (ShowCursorWindowSelector, CleanupCursorWindowSelector, etc.).
    status: pending
    dependencies: [remove_macro_registration]
---

# Focus Specific Cursor Window Macro

## Analysis / Context
The user requires a workflow to manage multiple "Cursor" application windows. The goal is to quickly select one specific window to keep focused while automatically closing all other open Cursor windows. This functionality should be accessible via the existing `Win+Alt+Shift+L` (Project Selector) menu, specifically bound to the 'h' key. The interaction model should be modal, keyboard-centric (single keypress), and consistent with existing Project Selector patterns in `WindowManagement.ahk`.

## Modifications from Original Plan

1. **UI Styling:** Remove the custom dark background (`BackColor := "1a1a1a"`) from the modal window. Use the standard default GUI background instead.

2. **Migration:** Move the "Focus Cursor Window" functionality from the `Win+Alt+Shift+U` (General Macros/Hotstring Selector) menu to the `Win+Alt+Shift+L` (Project Selector) menu.

3. **Refactoring:** Adapt the logic to align with the specific architecture of `Win+Alt+Shift+L`, which is dedicated to window opening and management functions rather than generic macros. Ensure the new feature integrates correctly with the Project Selector context.

## Proposed Changes

1. **Integration into Project Selector**: Instead of creating a standalone selector, add the Cursor window listing as a special item in the Project Selector menu, appearing after the project list and before the "[3] Activate Preview Windows" option.

2. **Character Assignment**: Use character 'h' from `g_ProjectCharSequence` for the Cursor window selector trigger.

3. **Selection Logic**: Implement `HandleCursorWindowSelection` which will:
   * Receive the selected window handle (HWND) and the list of all Cursor windows.
   * Activate the selected window.
   * Close all other windows in the list.
   * Close the selector GUI.

4. **Display Logic**: When 'h' is pressed in the Project Selector:
   * Get all Cursor windows using `WinGetList("ahk_exe Cursor.exe")`.
   * If no windows found, show notification and return.
   * If one window, activate it and return.
   * If multiple windows, show a sub-menu (similar to project selector style) with window titles and character assignments.

## Files to Modify

* `c:\Users\eduev\Meu Drive\12 - Scripts\WindowManagement.ahk` - Add Cursor window selection functionality
* `c:\Users\eduev\Meu Drive\12 - Scripts\Utils.ahk` - Remove macro registration and standalone selector functions

## Implementation Strategy

1. **Define Globals in WindowManagement.ahk**: Add `g_CursorWindowMap` and `g_CursorWindowHotkeyHandlers` to manage state.

2. **Selection Handler**: Create `HandleCursorWindowSelection(targetHwnd, allWindows)`:
   * Iterate through `allWindows`.
   * If a window's HWND does not match `targetHwnd`, call `WinClose`.
   * Call `WinActivate` for `targetHwnd`.
   * Call `CleanupProjectSelector`.

3. **Character Handler Factory**: Create `CreateCursorWindowSelectionHandler(char)` that returns a handler function for the assigned character.

4. **Integration into ShowProjectSelector**:
   * After displaying projects, check if 'h' character is available and assign it for "Focus Cursor Window".
   * In the display text, add "[h] Focus Cursor Window" option.
   * When 'h' is pressed, check if multiple Cursor windows exist:
     - If 0 windows: Show notification, cleanup selector
     - If 1 window: Activate it, cleanup selector
     - If 2+ windows: Show a new modal GUI with window list, assign characters from `g_ProjectCharSequence`, enable hotkeys

5. **Sub-menu GUI**: When multiple windows exist, create a new GUI (reusing Project Selector styles, but with default background):
   * Display window titles with character assignments
   * Enable hotkeys for selection
   * On selection, activate target window and close others
   * On Escape, close the sub-menu GUI

6. **Cleanup**: Remove macro registration from `InitMacros()` and remove standalone Cursor selector functions from `Utils.ahk`.
