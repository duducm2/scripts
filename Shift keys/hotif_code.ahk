; =============================================================================
; Shift keys module: hotif_code.ahk
; VS Code hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

#HotIf IsCodeFamilyActive() && SafeWinGetClass() != "#32770"

; Alt + M : Quick shortcut menu for VS Code (empty placeholder for future actions)
!m:: {
    ShowVSCodeShortcutMenu()
}

; Alt + C : Open VS Code chat "Add Context" picker (native Ctrl+;)
!c:: {
    Send "^;"
}

; Alt+A / Ctrl+Alt+. / Ctrl+M freed (unused — chords available for future use).
; !a:: VSCode_AddFileToAIContext()
; ^!.:: VSCode_TriggerGenerateCommitMessage(...)
; Ctrl+M commit/push flow kept as unbound helper:
VSCode_CommitPushFlow_Unused(*) {

    global gCommitPushTargetWin
    global gCommitPushDecision
    hwnd := WinExist("A")
    if !hwnd
        return

    gCommitPushTargetWin := hwnd
    ; Default behavior: commit + push, unless user presses N
    gCommitPushDecision := "push"

    ; 1. Trigger generation by clicking the native button (supports both plain and shortcut-suffixed names)
    if (!VSCode_TriggerGenerateCommitMessage(hwnd)) {
        ShowCenteredOverlay_Utils("Generate Commit Message button not found.", 2000, BANNER_ACCENT_ERROR)
        return
    }

    ScriptSoundPlay(A_ScriptDir "\assets\sounds\commit-start.wav")
    ShowCommitPushBanner()

    ; 2. Wait 14s for message generation to complete; user can interact with any window
    Sleep 14000

    ; 3. Focus back to VS Code (save current foreground to return later)
    prevHwnd := WinExist("A")
    WinActivate("ahk_id " hwnd)
    if !WinWaitActive("ahk_id " hwnd, , 3)
        return

    ; 4. Execute commit (Shift+V) and push (Shift+B) if user didn't press N
    Send "+v"
    didPush := (gCommitPushDecision = "push")
    if (didPush) {
        Sleep 500
        Send "+b"
    }

    ; 5. Wait for git operations to complete
    if (didPush) {
        Sleep 4000
    } else {
        Sleep 1500
    }

    ; Decide whether to return: stay in VS Code if we pushed so user can review
    shouldReturn := !didPush
    gCommitPushDecision := ""

    ; 6. Return to previous window (if user opted out of push)
    if (shouldReturn && prevHwnd && prevHwnd != hwnd) {
        if (!WinExist("ahk_id " prevHwnd)) {
            TrayTip("Commit Push", "Previous window no longer available; staying in VS Code.", "Iconi")
            SetTimer(() => TrayTip(), -5000)
        } else {
            try {
                WinActivate("ahk_id " prevHwnd)
                if (!WinWaitActive("ahk_id " prevHwnd, , 2)) {
                    TrayTip("Commit Push", "Could not switch back to previous window.", "Iconi")
                    SetTimer(() => TrayTip(), -5000)
                }
            } catch {
                TrayTip("Commit Push", "Could not switch back to previous window.", "Iconi")
                SetTimer(() => TrayTip(), -5000)
            }
        }
    }
}

#HotIf IsAntigravityActive() && SafeWinGetClass() != "#32770"

; Ctrl+Shift+L — select all identical words. Ctrl+L stays agent line-reference (like Cursor Ctrl+I).
$^+l:: {
    Send "^+l"
}

; Ctrl+I — same as native Antigravity Ctrl+L (Cursor Ctrl+I equivalent).
; Leave Ctrl+L unbound so the IDE owns agent line-ref / focus.
$^i:: Send "^l"

#HotIf