````markdown
---
name: Debug Cursor Transfer Modal Visibility
overview: Generate 7 hypotheses for why the Cursor transfer modal fails to appear, and add targeted defensive fixes and telemetry to isolate the root cause.
todos:
  - id: filter_garbage_cursor_windows
    content: Update the WinGetList loop in CursorTransfer_ShowWindowSelector to ignore windows with empty titles or 'preview' in the title, preventing hidden IPC windows from populating the modal.
    status: pending
  - id: remove_owner_property
    content: Remove the '+Owner' option from the Gui creation string in CursorTransfer_ShowWindowSelector to prevent the GUI from inheriting hidden status from the script's main window.
    status: pending
    dependencies: [filter_garbage_cursor_windows]
  - id: apply_wildcard_hotkeys
    content: Prepend an asterisk (*) to all Hotkey registrations (e.g., "*1", "*Escape", "*N") in CursorTransfer_ShowWindowSelector to prevent modifier-key bleed, and ensure CursorTransfer_SelectorClose unregisters the same wildcard keys.
    status: pending
    dependencies: [remove_owner_property]
  - id: instrument_gui_lifecycle
    content: Wrap the GUI display, WinActivate, and while-loop block in a try...catch. Inject JSON debug logs for GUI bounds (cx, cy, gw, gh), WinActivate success, while-loop duration, and any caught exceptions.
    status: pending
    dependencies: [apply_wildcard_hotkeys]
---

# Debug Cursor Transfer Modal

## Analysis / Context

The user requested 5-8 new hypotheses for why the Cursor transfer modal (D2C "C" path) fails to appear or closes instantly, along with concrete steps to verify them. Previous debugging fixed a type-casting error, but the `showing GUI` log line was never reached, or the modal remains invisible.

### Hypotheses

1. **Hypothesis 1: Hidden Owner Window (`+Owner`).**
   - _Reason:_ Using `+Owner` without specifying a parent HWND attaches the GUI to the script's hidden main window. Depending on Windows OS heuristics, this can cause the GUI to never render on-screen or instantly drop to the bottom of the Z-order.
   - _Action:_ Remove `+Owner` from the GUI creation options.

2. **Hypothesis 2: Exception Thrown During GUI Construction (Uncaught).**
   - _Reason:_ Although the sorting type-error was fixed, another property might be throwing a runtime exception (e.g., `w.hotkeyChar` being unset, or a variable typo) causing the thread to abort right before `Gui.Show()`.
   - _Action:_ Wrap the entire GUI construction and display block in a `try...catch` and log the error to `debug-7432d8.log`.

3. **Hypothesis 3: Hotkey Prefix Bleed (Physical Key State).**
   - _Reason:_ If the user physically holds down 'C' or a modifier when the GUI appears, hotkeys registered without the wildcard prefix (e.g., `Hotkey("1", ...)` instead of `Hotkey("*1", ...)`) will be ignored by AutoHotkey.
   - _Action:_ Add the `*` wildcard prefix to all hotkey registrations in the selector.

4. **Hypothesis 4: Empty Result / Instant Loop Exit.**
   - _Reason:_ A race condition or physical key bounce might cause one of the newly registered hotkeys (like `N` or `Escape`) to fire instantly before the GUI visually renders, setting the result and terminating the blocking loop in 0 ms.
   - _Action:_ Log the `durationMs` (`A_TickCount - start`) and the `result` immediately after the `while` loop exits.

5. **Hypothesis 5: 0x0 Size / Positioning Math Error.**
   - _Reason:_ `MonitorGetWorkArea` might yield negative or unexpected values if the primary monitor is offset, pushing the GUI entirely off-screen.
   - _Action:_ Log the calculated `cx`, `cy`, `gw`, and `gh` immediately after `GetPos()`.

6. **Hypothesis 6: GUI Activated but Instantly Loses Focus.**
   - _Reason:_ `WinActivate` might silently fail due to background process restrictions, leaving the GUI buried behind the active window.
   - _Action:_ Log whether `WinGetID("A") == g_CursorTransferSelectorGui.Hwnd` immediately after the `WinActivate` call.

7. **Hypothesis 7: `DetectHiddenWindows` Populating Garbage Windows.**
   - _Reason:_ `WinGetList("ahk_exe Cursor.exe")` is called while `DetectHiddenWindows true` is active. This captures invisible IPC or tooltip windows. The GUI renders with empty titles, and the actual project window is pushed off the list of 9.
   - _Action:_ Filter out windows with empty titles or "preview" during the enumeration loop.

## Proposed Changes

We will apply defensive fixes for the visibility, hotkeys, and garbage windows, and inject targeted telemetry logging to capture GUI coordinates, loop duration, and any uncaught exceptions.

## Files to Modify

- `c:\Users\eduev\Meu Drive\17 - Projects\scripts\Utils.ahk` (around `CursorTransfer_ShowWindowSelector`)

## Implementation Strategy

1. **Filter Garbage Windows (`Utils.ahk`)**:
   Inside `CursorTransfer_ShowWindowSelector`, locate the `WinGetList` loop. Change it to skip empty titles and preview windows:
   ```autohotkey
   for hwnd in WinGetList("ahk_exe Cursor.exe") {
       try {
           title := WinGetTitle("ahk_id " hwnd)
           if (title = "" || InStr(StrLower(title), "preview"))
               continue
           list.Push({ hwnd: hwnd, title: title })
           if (list.Length >= 9)
               break
       } catch {
           continue
       }
   }
   ```
````
