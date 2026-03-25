e---
name: Clip Angel Sequential Paste Automation
overview: Automate the rapid, sequential pasting of multiple clips from Clip Angel by creating a user prompt to capture the item count and executing a precise sequence of initial and follow-up paste commands.
todos:
  - id: create_execution_function
    content: Define a function `ExecuteSequentialPaste(count)` that executes the Initial Paste Command once, followed by a loop executing the Sequential Paste Command `count - 1` times, with appropriate safety delays between keystrokes.
    status: completed
  - id: create_gui_prompt
    content: Define a function `ShowSequentialPasteMenu()` that displays a small, auto-submitting GUI or InputBox to capture the number of clips the user wants to paste. On submit, this function should call `ExecuteSequentialPaste(count)`.
    status: completed
    dependencies: ["create_execution_function"]
  - id: bind_global_hotkey
    content: Register the `#!+j` (Win+Alt+Shift+J) hotkey to call `ShowSequentialPasteMenu()`.
    status: completed
    dependencies: ["create_gui_prompt"]
  - id: update_cheat_sheet
    content: Update the global shortcut cheat sheet variable in `Shift keys.ahk` to replace the "Available" placeholder for `[Win+Alt+Shift+J]` with the new sequential paste description.
    status: completed
---

# Clip Angel Sequential Paste Automation

## Analysis / Context
Currently, the user has an Initial Paste Command mapped to `Win + Alt + Shift + 1` (`#!+1`) which opens Clip Angel and pastes the top list item. Clip Angel also natively supports a "paste and select previous" function mapped to `Ctrl + Alt + B` (`^!b`). To paste an active queue of N clips, the user currently has to do this manually. This plan introduces an automated pipeline triggered by `Win + Alt + Shift + J` (`#!+j`) that asks for the number of clips (N) and orchestrates the correct sequence of keystrokes.

## Proposed Changes
1. **Execution Logic:** Create a loop-based AutoHotkey function that simulates the exact keystrokes required. Since `Win+Alt+Shift+1` is an existing AHK hotkey, the script will simulate its underlying action (`!v` then `^!b`) or use `SendLevel 1` to trigger it. Then, it will loop `N - 1` times sending `^!b`.
2. **User Interface:** Implement a minimal UI popup (similar to the existing AI Model Selector or Emoji Selector in `Shift keys.ahk`) allowing the user to type a single digit (e.g., 2-9) to immediately trigger the paste sequence. 
3. **Hotkey Binding:** Bind `#!+j`.
4. **Documentation:** Update `ShowGlobalShortcutsHelp()` within `Shift keys.ahk`.

## Files to Modify
- `c:\Users\eduev\Meu Drive\17 - Projects\scripts\Shift keys.ahk`
  - Locate the Global Shortcuts Help string (`=== AVAILABLE SECONDARY...`) to update the documentation.
  - Locate the end of the script or the Clip Angel section to append the new GUI and hotkey logic.

## Implementation Strategy

1. **Step 1: Build the Execution Function**
   - Create `ExecuteSequentialPaste(actionCount)` in `Shift keys.ahk`.
   - Validate that `actionCount` is an integer and `> 0`.
   - For the initial paste, replicate the `#!+1` action: `Send "!v"`, `Sleep 50`, `Send "^!b"`.
   - Calculate remaining clips: `remaining := actionCount - 1`.
   - Create a `loop remaining`:
     - Wait for the previous paste to process in the target application (e.g., `Sleep 300`).
     - Send the sequential paste command: `Send "^!b"`.

2. **Step 2: Build the Input GUI**
   - Create `ShowSequentialPasteMenu()` that initializes an `+AlwaysOnTop +ToolWindow` GUI.
   - Add instructional text (e.g., "How many clips to paste sequentially? (1-9)").
   - Add a `Limit1 Number` Edit control with an `OnEvent("Change", ...)` handler that automatically captures the input, destroys the GUI, and passes the number to `ExecuteSequentialPaste(count)`.
   - Fallback: Ensure an "Escape" hotkey is temporarily bound to cancel and close the GUI (like `CancelEmoji` or `CancelAIModel`).

3. **Step 3: Bind the Hotkey**
   - In `Shift keys.ahk`, outside of any app-specific `#HotIf` block (or in the global section), define:
     ```ahk
     #!+j::ShowSequentialPasteMenu()
     ```

4. **Step 4: Update Documentation**
   - In `Shift keys.ahk`, locate `ShowGlobalShortcutsHelp()`.
   - Find the line `[Win+Alt+Shift+J] > Available`.
   - Replace it with `[Win+Alt+Shift+J] > Paste multiple clips sequentially (Clip Angel)`.
