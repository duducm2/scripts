---
name: Fix D2C Step 2 Execution and State Corruption
overview: Resolve the crashing options ("Copy", "Copy and read", "Send to Cursor") in Step 2 of the dictation flow by replacing invalid dynamic function calls with IPC/hotkeys, and enforce try/finally state resets to prevent flow corruption.
todos:
  - id: enforce_try_finally_resets
    content: Wrap the internal logic of D2C_FlowManager.ExecuteAction and D2C_FlowManager.PromptForCursorTransfer in try...finally blocks that strictly call this.Reset() to guarantee state clearance even if execution fails.
    status: pending
  - id: rewrite_docopycore_logic
    content: Rewrite D2C_FlowManager.DoCopyCore to eliminate all Func("...").Call() commands. Implement readAloud via Send("#!+o"). Implement standard copying via PostMessage(0x8001) IPC to Gemini.ahk, followed by a clipboard change verification and direct SoundPlay.
    status: pending
    dependencies: [enforce_try_finally_resets]
---

# Fix D2C Step 2 Execution and State Corruption

## Analysis / Context

The Dictation-to-Cursor flow correctly reaches Step 2 ("Copy response?"), but executing "Copy", "Copy and read", or "Send to Cursor" immediately halts the system and breaks subsequent runs.

The root cause is inside `D2C_FlowManager.DoCopyCore()`. It attempts to execute functions located in `Gemini.ahk` (like `CopyLastGeminiMessageWithRetry`, `PlayCopyCompletedChime`, and `GeminiTriggerReadAloud`) using dynamic function calls (`Func("...").Call()`). Because `Utils.ahk` operates in an isolated script process (e.g., `WindowManagement.ahk`), these functions do not exist in its memory space. The `Func` evaluation fails, throwing an unhandled `TargetError`.

Because this error forcefully aborts the thread, the state machine's `this.Reset()` method is never executed. The flow remains permanently locked in the `"PromptingAction"` or `"Transferring"` phase, refusing to restart.

## Proposed Changes

1. **Guaranteed State Reset**: Wrap `ExecuteAction` and `PromptForCursorTransfer` logic inside `try...finally { this.Reset() }` blocks. This ensures the system state is strictly cleared regardless of upstream errors.
2. **Cross-Process Execution (IPC & Hotkeys)**: Completely remove `Func().Call()` from `DoCopyCore()`.
   - For "Read Aloud", simulate the global hotkey: `Send("#!+o")`.
   - For "Copy", use the established IPC message `WM_COPY_LAST_GEMINI (0x8001)` to command `Gemini.ahk` to perform the copy silently without stealing focus.
   - For the completion chime, invoke `SoundPlay` directly instead of calling a missing helper function.

## Files to Modify

- `c:\Users\eduev\Meu Drive\17 - Projects\scripts\Utils.ahk`

## Implementation Strategy

1. **Fix `ExecuteAction` (`Utils.ahk`)**:
   Update `D2C_FlowManager.ExecuteAction` to:
   ```autohotkey
   ExecuteAction(readAloud := false, skipRestoreFocus := false) {
       this.CleanupActionPrompt()
       try {
           this.DoCopyCore(readAloud, skipRestoreFocus)
       } finally {
           this.Reset()
       }
   }
   ```

Fix PromptForCursorTransfer (Utils.ahk): Locate PromptForCursorTransfer() and wrap everything after this.CleanupActionPrompt() in a try...finally block that calls this.Reset(). (Remove the existing this.Reset() at the bottom).
PromptForCursorTransfer() {
this.CurrentPhase := "Transferring"
this.CleanupActionPrompt()
try {
this.DoCopyCore(false, true)

        clip := Trim(A_Clipboard)
        if (clip = "" || StrLen(clip) < 10) {
            Sleep 500
            clip := Trim(A_Clipboard)
        }

        ; ... (keep rest of the existing logic: if clip == "", CursorTransfer_ShowWindowSelector, CursorTransfer_ActivateFocusPaste) ...

    } finally {
        this.Reset()
    }

}

Rewrite DoCopyCore (Utils.ahk): Replace the entire DoCopyCore method with the following robust IPC/Hotkey logic:
DoCopyCore(readAloud := false, skipRestoreFocus := false) {
if (this.HasCopiedForThisResponse)
return
this.HasCopiedForThisResponse := true

    if (readAloud) {
        ; Trigger read aloud (which also copies) via global hotkey
        Send("#!+o")
        Sleep(500)
    } else {
        ; Silent copy via IPC to Gemini.ahk
        clipBefore := A_Clipboard

        DetectHiddenWindows(true)
        prevMatch := A_TitleMatchMode
        SetTitleMatchMode(2)
        postTargetHwnd := 0

        for hwnd in WinGetList("ahk_exe AutoHotkey64.exe") {
            try {
                if (InStr(WinGetTitle("ahk_id " hwnd), "Gemini.ahk")) {
                    postTargetHwnd := hwnd
                    break
                }
            } catch {
                continue
            }
        }
        if (!postTargetHwnd) {
            for hwnd in WinGetList("ahk_exe AutoHotkey32.exe") {
                try {
                    if (InStr(WinGetTitle("ahk_id " hwnd), "Gemini.ahk")) {
                        postTargetHwnd := hwnd
                        break
                    }
                } catch {
                    continue
                }
            }
        }

        if (postTargetHwnd) {
            try PostMessage(0x8001, 0, 0, , "ahk_id " postTargetHwnd)
            loop 40 {
                Sleep(50)
                if (A_Clipboard != clipBefore && Trim(A_Clipboard) != "")
                    break
            }
            if (IsSoundEnabled()) {
                try SoundPlay(A_ScriptDir . "\sounds\copy.wav")
            }
        } else {
            ShowCenteredOverlay_Utils("❌ Gemini.ahk not running", 2000, BANNER_ACCENT_ERROR)
        }

        SetTitleMatchMode(prevMatch)
        DetectHiddenWindows(false)
    }

    if (!skipRestoreFocus && this.OriginHwnd && WinExist("ahk_id " this.OriginHwnd) && !WinActive("ahk_id " this.OriginHwnd)) {
        WinActivate("ahk_id " this.OriginHwnd)
        WinWaitActive("ahk_id " this.OriginHwnd, , 0.5)
    }

}

<!--
[PROMPT_SUGGESTION]Apply the state reset and execution fixes to Utils.ahk per the plan.[/PROMPT_SUGGESTION]
-->
