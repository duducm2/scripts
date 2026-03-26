---
name: Add File To Cursor Chat Quality Check
overview: Add a deterministic post-action quality check for the Alt+A flow in Cursor so the script can confirm the file-add action succeeded and report success or failure reliably.
todos:
  - id: create_failure_detector
    content: Create a new function Cursor_DetectAddFileFailureSignal() in Shift keys.ahk that returns a specific error string if obvious failure states are detected (e.g., missing window, UIA unavailable).
    status: pending
  - id: create_success_verifier
    content: Create a new function Cursor_WaitForAddFileToChatSuccess(timeoutMs) in Shift keys.ahk that polls the UIA tree for the AI composer input (aislash-editor-input) to verify the chat context successfully opened and is ready.
    status: pending
    dependencies: [create_failure_detector]
  - id: update_activation_function
    content: Modify Cursor_ContextMenuActivateHighlightedItem() to evaluate success using the new verifier after the context menu closes, integrating a single controlled retry if verification fails.
    status: pending
    dependencies: [create_success_verifier]
  - id: update_main_hotkey
    content: Update the !a:: hotkey and Cursor_ContextMenuSelectByDownAndVerify() to use the verified success state, propagating specific failure reasons to StandardLoadingBar_Update for accurate user feedback.
    status: pending
    dependencies: [update_activation_function]
---

# Add File To Cursor Chat Quality Check

## Analysis / Context
The `!a` (Alt+A) hotkey in `Shift keys.ahk` adds a selected file from the Explorer to the Cursor Chat context by invoking the context menu. Currently, success is assumed if the context menu item is clicked and the menu subsequently disappears (`Cursor_WaitForContextMenuItemGone`). This is prone to false positives (e.g., the menu dismisses due to lost focus or an errant click, but the file isn't actually added). A deterministic post-action quality check is needed to ensure the UI state correctly reflects a successful file attachment before reporting success to the user.

## Proposed Changes
1.  **Define a Success Signal:** When a file is successfully added to Cursor chat, the AI pane must become active, and the composer input (`aislash-editor-input`) must be present on-screen.
2.  **Failure Signal Detection:** Identify obvious failure states (e.g., target window lost, AI pane completely unavailable) to provide descriptive error messages.
3.  **Enhance Activation Logic:** Update `Cursor_ContextMenuActivateHighlightedItem` to poll for this success signal after the menu closes, utilizing a short retry mechanism if the signal is not immediately detected.
4.  **Feedback Propagation:** Bubble up the verification result to the `!a::` hotkey so the `StandardLoadingBar_Update` explicitly states if the operation was verified or failed (and the specific reason).

## Files to Modify
- `c:\Users\fie7ca\Documents\scripts\Shift keys.ahk`
  - Region around `!a::` hotkey (approx. line 1993)
  - Region around `Cursor_ContextMenuSelectByDownAndVerify` and `Cursor_ContextMenuActivateHighlightedItem` (approx. line 2063)

## Implementation Strategy
1.  **Implement `Cursor_DetectAddFileFailureSignal()`:**
    *   Check if `WinExist("ahk_exe Cursor.exe")` evaluates to false (return `"Target window closed"`).
    *   Check if `UIA.ElementFromHandle` fails to attach (return `"UIA unreachable"`).
    *   Return an empty string `""` if no explicit failure is detected.
2.  **Implement `Cursor_WaitForAddFileToChatSuccess(timeoutMs := 1800)`:**
    *   Use a `while` loop bounded by `A_TickCount < deadline`.
    *   Inside the loop, get the `UIA_Browser` or `UIA.ElementFromHandle` for Cursor.
    *   Search for the composer input: `Type: UIA.Type.Edit`, `ClassName` containing `aislash-editor-input`.
    *   If the composer input is found and `IsOffscreen` is false, consider it a success and return `true`.
    *   Return `false` if the timeout is reached.
3.  **Update `Cursor_ContextMenuActivateHighlightedItem`:**
    *   After the existing `closed := Cursor_WaitForContextMenuItemGone(...)` logic, add the verification step.
    *   If `closed` is true, call `Cursor_WaitForAddFileToChatSuccess(1000)`.
    *   If it returns false, attempt one retry (e.g., attempt to re-send Enter if the menu item element is still valid) and wait again.
    *   Update the function signature/return to provide an object containing both the boolean status and a reason string: `{ok: true/false, reason: ""}`.
4.  **Update `Cursor_ContextMenuSelectByDownAndVerify` & `!a::`:**
    *   Adjust `Cursor_ContextMenuSelectByDownAndVerify` to pass the detailed result object up to `!a::`.
    *   In `!a::`, evaluate the result.
    *   If success: `StandardLoadingBar_Update("✅ File added to Cursor Chat")`.
    *   If failure: retrieve the reason via the returned object (falling back to `Cursor_DetectAddFileFailureSignal()` if empty), and display `StandardLoadingBar_Update("❌ Failed: " . reason)`.
5.  **Add Minimal Diagnostic Logging:**
    *   Insert calls to `SafeDebugLog` strictly in the failure paths of `Cursor_WaitForAddFileToChatSuccess` to trace exactly why verification failed without spamming the success paths.