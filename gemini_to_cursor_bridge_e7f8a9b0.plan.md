---
name: Gemini to Cursor Bridge
overview: Implement a new "Copy from Gemini" mode (Key K) in the Project Selector that copies the last Gemini response and pastes it into the selected Cursor project's AI field.
todos:
  - id: refactor_project_selection_logic
    content: Refactor HandleSelectionModeProjectSelection in WindowManagement.ahk to separate project activation/focus logic for reuse.
    status: pending
    dependencies: []
  - id: define_copy_mode_globals
    content: Define global variables for Copy from Gemini mode (g_CopyFromGeminiModeActive, handlers) in WindowManagement.ahk.
    status: pending
    dependencies: []
  - id: implement_copy_mode_trigger
    content: Implement HandleCopyFromGeminiModeTrigger to activate the mode when 'K' is pressed in the Project Selector.
    status: pending
    dependencies: [define_copy_mode_globals]
  - id: implement_copy_mode_handler
    content: Implement HandleCopyFromGeminiProjectSelection to execute the copy-activate-paste-submit workflow.
    status: pending
    dependencies:
      [refactor_project_selection_logic, implement_copy_mode_trigger]
  - id: update_project_selector_gui
    content: Update ShowProjectSelector GUI to display the new '[K] Copy from Gemini' option.
    status: pending
    dependencies: [implement_copy_mode_trigger]
  - id: register_hotkeys
    content: Register the 'K' hotkey in ShowProjectSelector and the project hotkeys in the new mode.
    status: pending
    dependencies: [implement_copy_mode_handler]
---

# Gemini to Cursor Bridge

## Analysis / Context

The current Project Selector (`Win+Alt+Shift+L`) allows users to open a project and focus the AI text field using Selection Mode (`L`). A common workflow involves copying a response from Gemini and pasting it into a Cursor project. Currently, this requires manual switching or multiple shortcuts. The new "Copy from Gemini" mode (`K`) streamlines this by automating the fetch-and-paste sequence directly from the project selector.

## Proposed Changes

1.  **Refactor Existing Logic:** Extract the core logic for activating a project window and focusing the AI text field from `HandleSelectionModeProjectSelection` into a reusable function (e.g., `ActivateProjectAndFocusAI`).
2.  **Implement "Copy from Gemini" Mode (`K`):**
    - Create a new mode state `g_CopyFromGeminiModeActive`.
    - When `K` is pressed in the Project Selector, enter this mode (similar to `L`).
    - When a project is selected in this mode:
      1.  Trigger `Win+Alt+Shift+P` (Copy last Gemini message).
      2.  Wait briefly for the clipboard to update.
      3.  Call the refactored `ActivateProjectAndFocusAI` to open the project.
      4.  Send `Ctrl+V` to paste.
      5.  Send `Enter` to submit.

## Files to Modify

- `c:\Users\eduev\Meu Drive\12 - Scripts\WindowManagement.ahk`

## Implementation Strategy

1.  **Refactoring:**
    - Locate `HandleSelectionModeProjectSelection` in `WindowManagement.ahk`.
    - Extract the window finding, launching, and AI field focusing logic into `ActivateCursorProject(projectPath)`.
2.  **Mode Implementation:**
    - Define `g_CopyFromGeminiModeActive` and `g_CopyFromGeminiHotkeyHandlers`.
    - Create `HandleCopyFromGeminiModeTrigger` to toggle the mode and set up hotkeys (reusing the project char mapping logic).
    - Create `HandleCopyFromGeminiProjectSelection(index)`:
      - Send `#!+p` to invoke `Gemini.ahk`'s copy function.
      - `Sleep 500` (or check clipboard).
      - Call `ActivateCursorProject`.
      - `Send "^v"`.
      - `Sleep 100`.
      - `Send "{Enter}"`.
3.  **UI & Hotkeys:**
    - Update `ShowProjectSelector` to show `[K] Copy from Gemini`.
    - Register `k` and `K` in `ShowProjectSelector` to call `HandleCopyFromGeminiModeTrigger`.
    - Ensure `CleanupProjectSelector` handles the new mode variables.
