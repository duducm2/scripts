---
name: Refactor Gemini Pronunciation Workflow
overview: Refactor the Win+Alt+Shift+8 hotkey to be asynchronous, allowing the user to continue working while Gemini generates a response, then notifying them with a non-blocking UI banner.
todos:
  - id: define_async_class
    content: Define GeminiAsyncLookup class in Gemini.ahk to encapsulate state (original HWND, timer, retry count).
    status: pending
    dependencies: []
  - id: implement_init_logic
    content: Implement Start() method to capture active HWND, copy text, activate Gemini, submit prompt, and immediately restore focus to the original window.
    status: pending
    dependencies: [define_async_class]
  - id: implement_monitoring
    content: Implement CheckCompletion() method using SetTimer to poll Gemini (background UIA) for the disappearance of 'Stop streaming'.
    status: pending
    dependencies: [implement_init_logic]
  - id: implement_retrieval
    content: Implement RetrieveResponse() method to activate Gemini, trigger #!+p (Copy Last Message), and capture the clipboard content.
    status: pending
    dependencies: [implement_monitoring]
  - id: implement_context_restore_final
    content: Implement RestoreContext() method to reactivate the stored original window HWND after copying.
    status: pending
    dependencies: [implement_retrieval]
  - id: implement_result_gui
    content: Create ShowResultBanner(text) method to display the definition in a styled, centered GUI with 12s timeout and Esc close handler.
    status: pending
    dependencies: [implement_context_restore_final]
  - id: update_hotkey
    content: Update #!+8 hotkey in Gemini.ahk to instantiate and start the GeminiAsyncLookup workflow.
    status: pending
    dependencies: [implement_result_gui]
---

# Refactor Gemini Pronunciation Workflow

## Analysis / Context

The current `Win+Alt+Shift+8` workflow is synchronous and blocking. It forces the user to wait on the Gemini window while the AI generates a response, which disrupts flow. The desired state is an asynchronous "fire-and-forget" workflow where the user can immediately resume working, and the system notifies them via a banner when the definition is ready.

## Proposed Changes

1.  **State Management:** Introduce a `GeminiAsyncLookup` class to handle the lifecycle of the request.
2.  **Asynchronous Execution:**
    - Capture the user's current window handle (`HWND`).
    - Send the prompt to Gemini.
    - Immediately switch focus back to the user's window.
    - Use `SetTimer` to poll Gemini's state (checking for the "Stop streaming" button or completion indicators) without blocking the main thread.
3.  **Result Retrieval:** Once generation is complete, briefly activate Gemini to copy the result using the existing `#!+p` logic, then restore the user's window.
4.  **UI Feedback:** Display the result in a custom, aesthetically pleasing GUI banner that auto-dismisses or closes on Escape.

## Files to Modify

- `c:\Users\eduev\Meu Drive\12 - Scripts\Gemini.ahk`

## Implementation Strategy

1.  **Class Definition:** Create `class GeminiAsyncLookup` at the bottom of `Gemini.ahk`.
2.  **Initialization (`Start`):**
    - `this.OriginalHwnd := WinExist("A")`
    - Perform copy (`^c`).
    - Activate Gemini, paste prompt, send Enter.
    - `WinActivate("ahk_id " this.OriginalHwnd)` (Restore context immediately).
    - `SetTimer(this.CheckTimerObj, 500)`
3.  **Monitoring (`CheckCompletion`):**
    - Check if Gemini is processing (look for "Stop streaming" button via UIA).
    - If "Stop streaming" is gone (and "Copy" button exists), trigger retrieval.
4.  **Retrieval (`RetrieveResponse`):**
    - `WinActivate` Gemini.
    - Send `#!+p` (Copy last message).
    - Wait for Clipboard change.
    - `WinActivate("ahk_id " this.OriginalHwnd)`.
    - Call `ShowResultBanner(A_Clipboard)`.
5.  **Banner (`ShowResultBanner`):**
    - Create `Gui("+AlwaysOnTop -Caption +ToolWindow")`.
    - Add Text control with the content.
    - Center on screen.
    - Set timer for 12s to `Destroy()`.
    - Bind `Esc` to `Destroy()`.
