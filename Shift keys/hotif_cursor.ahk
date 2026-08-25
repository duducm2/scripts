; =============================================================================
; Shift keys module: hotif_cursor.ahk
; Cursor IDE hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

#HotIf IsCursorActive() && WinGetClass("A") != "#32770"

; Alt + M : Quick shortcut menu for Cursor
; Captionless dark-theme key menu (legacy Catppuccin style; Handy #!+C now uses ListView).
global g_CursorShortcutMenuGui := false
global g_CursorShortcutMenuActive := false
global g_CursorShortcutMenuEscPollPrev := false
global g_CursorShortcutMenuPrevHwnd := 0

!m::
{
    ShowCursorShortcutMenu()
}

; Alt + A : Add File to AI Context (Add File to Cursor Chat)
!a::
{
    StandardLoadingBar_Show("⏳ Add file to Cursor Chat...", BANNER_ACCENT_INTERMEDIATE, { passive: false,
        centerOnHwnd: 0 })
    try {
        ; Always move to Cursor primary sidebar Explorer, regardless of current focus/caret.
        StandardLoadingBar_Update("⏳ Focusing Explorer...")
        Send "^+e"
        Sleep 350
        okFocus := FocusCursorFilesExplorer()
        if (!okFocus) {
            Sleep 150
            okFocus := FocusCursorFilesExplorer()
        }
        if (!okFocus) {
            Sleep 150
            okFocus := FocusCursorFilesExplorer()
        }
        if (!okFocus) {
            StandardLoadingBar_Update("❌ Failed: Could not focus Explorer sidebar")
            return
        }
        Sleep 120

        expectedFileName := Cursor_GetFocusedExplorerItemName()
        if (expectedFileName = "") {
            StandardLoadingBar_Update("❌ Failed: Could not read selected file name")
            return
        }

        ; Open context menu for the selected file in Explorer, then navigate to the target item.
        StandardLoadingBar_Update("⏳ Opening context menu...")
        result := Cursor_ContextMenuSelectByDownAndVerifyAny(
            ["Add file to Gemini context", "Add File to Cursor Chat"],
            "{AppsKey}",
            34,
            expectedFileName
        )
        if (result.ok) {
            StandardLoadingBar_Update("✅ File added to Cursor Chat")
            return
        }

        ; Deterministic fallback for Cursor menu ID issues:
        ; copy selected file name via rename mode and insert @filename in AI field.
        StandardLoadingBar_Update("⏳ Fallback: @filename from Explorer selection...")
        fallbackResult := Cursor_FallbackAddFileByAtMention(expectedFileName)
        if (fallbackResult.ok) {
            StandardLoadingBar_Update("✅ File added to Cursor Chat")
            return
        }

        failureReason := result.reason
        if (failureReason = "")
            failureReason := fallbackResult.reason
        if (failureReason = "")
            failureReason := Cursor_DetectAddFileFailureSignal(expectedFileName)
        if (failureReason = "")
            failureReason := "Could not verify add-file action"
        StandardLoadingBar_Update("❌ Failed: " . failureReason)
    } finally {
        StandardLoadingBar_Hide(600)
    }
}

; Ctrl + M : Trigger commit message generation (Cursor-specific: Ctrl+Alt+. flow + commit/push banner)
^M:: {
    global gCommitPushTargetWin
    global gCommitPushDecision
    hwnd := WinExist("A")
    if !hwnd
        return
    gCommitPushTargetWin := hwnd
    ; Default behavior is now: commit + push, unless user opts out.
    gCommitPushDecision := "push"

    ; 1. Trigger generation immediately (Ctrl+Alt+A)
    ScriptSoundPlay(A_ScriptDir "\assets\sounds\commit-start.wav")
    Send "^!."
    ShowCommitPushBanner()

    ; 2. Wait 15s; user can interact with any window
    Sleep 14000
    ; Handoff Stop Sign: warn + play pre-movement cue right before we
    ; regain focus on Cursor and finalize the commit submission.
    PlayPreMovementWarning("Cursor")

    ; 3. Focus Cursor IDE (save current foreground to return later)
    prevHwnd := WinExist("A")
    WinActivate("ahk_id " hwnd)
    if !WinWaitActive("ahk_id " hwnd, , 3)
        return

    ; 4. Execute commit and push if necessary (default: push; user can opt out)
    Send "+v"
    didPush := (gCommitPushDecision = "push")
    if (didPush) {
        Sleep 500
        Send "+b"
    }

    ; 5. Wait for git operations to complete, then open Git panel to verify
    if (didPush) {
        Sleep 4000
    } else {
        Sleep 1500
    }
    ; Send "+d"

    ; Decide whether to return to previous window: stay in Cursor if we pushed
    shouldReturn := !didPush
    gCommitPushDecision := ""

    ; 6. Return to previous screen (graceful error if window no longer exists)
    if (shouldReturn && prevHwnd && prevHwnd != hwnd) {
        if (!WinExist("ahk_id " prevHwnd)) {
            TrayTip("Commit Push", "Previous window no longer available; staying in Cursor.", "Iconi")
            SetTimer(() => TrayTip(), -5000)
        } else {
            try {
                WinActivate("ahk_id " prevHwnd)
                if (!WinWaitActive("ahk_id " prevHwnd, , 2)) {
                    TrayTip("Commit Push", "Could not switch back to previous window.", "Iconi")
                    SetTimer(() => TrayTip(), -5000)
                }
            } catch {
                TrayTip("Commit Push", "Previous window no longer available; staying in Cursor.", "Iconi")
                SetTimer(() => TrayTip(), -5000)
            }
        }
    }
}
