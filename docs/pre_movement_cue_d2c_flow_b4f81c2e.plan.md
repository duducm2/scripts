---
name: Integrate Pre-Movement Sound Cue in D2C Flow
overview: Introduce a 3-second delay with an audio warning before all automated window activations and mouse movements within the Dictation to Gemini to Cursor pipeline to prevent synchronization conflicts with manual user input.
todos:
  - id: create_warning_helper
    content: Create the `PlayPreMovementWarning(targetName)` helper function in Utils.ahk to handle sound playback and the 3-second delay.
    status: pending
  - id: warn_gemini_submit
    content: Inject the pre-movement warning in `D2C_FlowManager.ExecuteGeminiSubmit` before activating Gemini for prompt submission.
    status: pending
    dependencies: [create_warning_helper]
  - id: warn_origin_return_submit
    content: Inject the pre-movement warning in `D2C_FlowManager.ExecuteGeminiSubmit` before restoring focus to the original window.
    status: pending
    dependencies: [create_warning_helper]
  - id: warn_gemini_copy
    content: Inject the pre-movement warning in `D2C_FlowManager.DoCopyCore` before activating Gemini to execute the copy action.
    status: pending
    dependencies: [create_warning_helper]
  - id: warn_origin_return_copy
    content: Inject the pre-movement warning in `D2C_FlowManager.DoCopyCore` before restoring focus to the original window post-copy.
    status: pending
    dependencies: [create_warning_helper]
  - id: warn_cursor_transfer
    content: Inject the pre-movement warning in `D2C_FlowManager.PromptForCursorTransfer` before activating the Cursor window for paste.
    status: pending
    dependencies: [create_warning_helper]
  - id: warn_legacy_monitor_copy
    content: Inject the pre-movement warning in `Gemini.ahk` within `GeminiDelayedSubmitMonitor.DoCopyCore` for Gemini activation and focus restoration.
    status: pending
    dependencies: [create_warning_helper]
  - id: warn_legacy_monitor_transfer
    content: Inject the pre-movement warning in `Gemini.ahk` within `GeminiDelayedSubmitMonitor.CopyAndTransferToCursor` before activating Cursor.
    status: pending
    dependencies: [create_warning_helper]
---

# Integrate Pre-Movement Sound Cue in D2C Flow

## Analysis / Context
The "Voice to Gemini to Cursor" (D2C) pipeline automates pasting text across multiple applications, meaning the script frequently takes control of window focus and keyboard/mouse operations. Because AutoHotkey operates synchronously, any manual user interaction (typing or clicking) at the exact moment the script attempts to automate a transition will cause race conditions and synchronization conflicts. 
To prevent this, the user requires an immediate audiovisual warning accompanied by a 3-second delay prior to any window activation or screen transition.

## Proposed Changes
1. **Centralized Helper**: Create a global helper function `PlayPreMovementWarning(targetName)` that standardizes the 3-second delay, plays `pre-movement.mp3`, and displays a non-blocking UI warning overlay.
2. **Flow Manager Integration**: Strategically inject this helper across the `D2C_FlowManager` state machine in `Utils.ahk` and the `GeminiDelayedSubmitMonitor` class in `Gemini.ahk` specifically right before any `WinActivate` sequence or transition logic.

## Files to Modify
- `c:\Users\fie7ca\Documents\scripts\Utils.ahk`
- `c:\Users\fie7ca\Documents\scripts\Gemini.ahk`

## Implementation Strategy

**Step 1: Implement the Helper Function (in `Utils.ahk`)**
Locate the bottom of the utility functions or near `ShowCenteredOverlay_Utils`. Add:
```ahk
PlayPreMovementWarning(targetName) {
    if (IsSoundEnabled()) {
        try SoundPlay(A_ScriptDir . "\sounds\pre-movement.mp3")
    }
    ShowCenteredOverlay_Utils("✋ Hands off! Moving to " . targetName . "...", 3000, BANNER_ACCENT_INTERMEDIATE)
    Sleep 3000
}
```

**Step 2: Update `Utils.ahk` - `D2C_FlowManager`**
- **ExecuteGeminiSubmit**: Before `GeminiNavigateFocusAndPasteFirstSnippet("", false)`, add `PlayPreMovementWarning("Gemini")`. 
- **ExecuteGeminiSubmit**: Inside the block `if (this.OriginHwnd && WinExist("ahk_id " this.OriginHwnd))`, right before `WinActivate`, add `PlayPreMovementWarning("Original Window")`.
- **DoCopyCore**: Inside the block `if (!WinActive("ahk_id " this.GeminiHwnd))`, right before the `try WinActivate(...)`, add `PlayPreMovementWarning("Gemini")`.
- **DoCopyCore**: Inside the block `if (!skipRestoreFocus && ...)`, right before `WinActivate(...)`, add `PlayPreMovementWarning("Original Window")`.
- **PromptForCursorTransfer**: Immediately before calling `CursorTransfer_ActivateFocusPaste(this.CursorHwnd)`, add `PlayPreMovementWarning("Cursor")`.

**Step 3: Update `Gemini.ahk` - `GeminiDelayedSubmitMonitor`**
- **DoCopyCore**: Apply the identical pattern as Step 2. Before `try { WinActivate("ahk_id " this.GeminiHwnd) }` add `PlayPreMovementWarning("Gemini")`. 
- **DoCopyCore**: Before restoring focus (`WinActivate("ahk_id " this.OriginalHwnd)`), add `PlayPreMovementWarning("Original Window")`.
- **CopyAndTransferToCursor**: Immediately before calling `CursorTransfer_ActivateFocusPaste(hwnd)`, add `PlayPreMovementWarning("Cursor")`.
