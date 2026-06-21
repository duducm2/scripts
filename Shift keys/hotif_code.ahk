; =============================================================================
; Shift keys module: hotif_code.ahk
; VS Code hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

#HotIf IsCodeActive() && WinGetClass("A") != "#32770"

; Alt + M : Quick shortcut menu for VS Code (empty placeholder for future actions)
!m:: {
    ShowVSCodeShortcutMenu()
}

; Alt + C : Open VS Code chat "Add Context" picker (native Ctrl+;)
!c:: {
    Send "^;"
}

; Alt + A : Add file to AI Context (VS Code Copilot chat)
!a:: {
    VSCode_AddFileToAIContext()
}

; Ctrl + Alt + . : Generate commit message in Source Control (explicit remap for reliability)
^!.:: {
    hwnd := WinExist("A")
    if (!hwnd)
        return
    if (!VSCode_TriggerGenerateCommitMessage(hwnd)) {
        ShowCenteredOverlay_Utils("Generate Commit Message button not found.", 2000, BANNER_ACCENT_ERROR)
    }
}

; Ctrl + M : Trigger commit message generation + commit/push flow (VS Code: native button + Shift+V/+B)
; Mirrors Cursor workflow: generate → wait with banner → commit → push (if user allows) → return
^M:: {
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

    ScriptSoundPlay(A_ScriptDir "\sounds\commit-start.wav")
    ShowCommitPushBanner()

    ; 2. Wait 14s for message generation to complete; user can interact with any window
    Sleep 14000

    ; Handoff Stop Sign: warn + play pre-movement cue before returning to VS Code
    PlayPreMovementWarning("VS Code")

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

