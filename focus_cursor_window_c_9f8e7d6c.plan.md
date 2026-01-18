---
name: Focus Cursor Window (Close Others)
overview: Implement a "Focus Cursor Window" feature in the Hotstring Selector (Win+Alt+Shift+U) triggered by 'C', which allows selecting one Cursor window to keep open while closing all others, preserving Project Selector shortcut consistency.
todos:
  - id: port_project_data
    content: Port `g_Projects` and `g_ProjectCharSequence` data structures from `WindowManagement.ahk` to `Utils.ahk` (or create a shared include) to enable consistent key mapping.
    status: pending
    dependencies: []
  - id: implement_window_discovery
    content: Create `GetCursorWindowsWithKeys()` in `Utils.ahk` to list open Cursor windows and assign selection keys, prioritizing existing Project Selector mappings.
    status: pending
    dependencies: [port_project_data]
  - id: implement_focus_logic
    content: Create `FocusCursorWindowAndCloseOthers(targetHwnd)` to activate the selected window and close all other Cursor instances.
    status: pending
    dependencies: [implement_window_discovery]
  - id: create_selection_gui
    content: Implement `ShowCursorFocusSelector()` GUI in `Utils.ahk` to display the list and handle user input.
    status: pending
    dependencies: [implement_focus_logic]
  - id: register_macro
    content: Register the new macro in `InitMacros()` in `Utils.ahk` with the specific trigger character 'C'.
    status: pending
    dependencies: [create_selection_gui]
---

# Focus Cursor Window (Close Others)

## Analysis / Context
The user wants a workflow to declutter their workspace by keeping only one specific Cursor window open. This feature should be accessible via the `Win+Alt+Shift+U` (Hotstring/Macro Selector) menu using the mnemonic key 'C' (for "Close" others).

Crucially, the key assignments in this new selector must be consistent with the `Win+Alt+Shift+L` (Project Selector) macro. If an open window corresponds to a known project (e.g., "Notes" is usually '2'), it must use '2' in this list as well.

## Proposed Changes

### 1. Data Porting / Sharing
To ensure consistency with `Win+Alt+Shift+L` (which resides in `WindowManagement.ahk`), `Utils.ahk` needs access to the project definitions and key mapping logic.
*   **Action:** Copy the `g_Projects` array and `g_ProjectCharSequence` (or relevant mapping logic) into `Utils.ahk`.
*   *Note:* While a shared include file would be cleaner architecturally, duplicating the definition in `Utils.ahk` is a self-contained change that avoids breaking dependencies in `WindowManagement.ahk` during this step.

### 2. Window Discovery & Key Assignment Logic
We need a function that:
1.  Enumerates all `ahk_exe Cursor.exe` windows.
2.  For each window, determines if it matches a project in `g_Projects` (matching title against path/name).
3.  **Key Assignment Strategy:**
    *   **Match Found:** Assign the key that the Project Selector would use (based on the project's index/category order).
    *   **No Match:** Assign the next available key from `g_ProjectCharSequence` that hasn't been reserved by a matched project.

### 3. Focus & Cleanup Logic
A handler function `FocusCursorWindowAndCloseOthers(targetHwnd)` that:
1.  Iterates through the list of all detected Cursor windows.
2.  If `hwnd == targetHwnd`, activate it.
3.  If `hwnd != targetHwnd`, close it (`WinClose`).

### 4. UI Implementation
Create a modal GUI `ShowCursorFocusSelector()` similar to the existing selectors but tailored for this action:
*   Title: "Focus Window (Close Others)"
*   List format: `[Key] Window Title`
*   Trigger: Registered in `InitMacros()` as `RegisterMacro(ShowCursorFocusSelector, "Focus Cursor Window", "c")`.

## Files to Modify

*   `c:\Users\eduev\Meu Drive\12 - Scripts\Utils.ahk`

## Implementation Strategy

1.  **Define Projects in Utils:** Add `g_Projects` and helper functions for project matching (similar to `ExtractProjectMatchSegments` in `WindowManagement.ahk`) to `Utils.ahk`.
2.  **Create Logic:** Implement `GetCursorWindowsWithKeys()` to build the list of windows and their assigned keys.
3.  **Create GUI:** Implement the GUI builder and input handler.
4.  **Register:** Add the macro registration in `InitMacros` with the 'C' trigger.
