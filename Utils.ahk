#Requires AutoHotkey v2.0+
#SingleInstance Force

; -----------------------------------------------------------------------------
; This script consolidates various utility hotkeys.
; -----------------------------------------------------------------------------

; --- Hotkeys & Functions -----------------------------------------------------

; =============================================================================
; Jump Mouse to Middle of Active Window
; Hotkey: Win+Alt+Shift+3
; Original File: Jump mouse on the middle.ahk
; =============================================================================
#!+Q::
{
    hwnd := WinExist("A")
    if !hwnd {
        return ; silently abort if no active window
    }
    rect := Buffer(16, 0)
    if !DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect) {
        MsgBox "GetWindowRect failed"
        return
    }
    left := NumGet(rect, 0, "int")
    top := NumGet(rect, 4, "int")
    right := NumGet(rect, 8, "int")
    bottom := NumGet(rect, 12, "int")
    centerX := left + (right - left) // 2
    centerY := top + (bottom - top) // 2
    DllCall("SetCursorPos", "int", centerX, "int", centerY)
}

; =============================================================================
; Activate Cursor and Send Key Sequence with Options
; Hotkey: Win+Alt+Shift+C
; =============================================================================

; Function to activate Cursor and send key sequence based on user choice
CursorKeySequence() {
    try {
        ; Get user input for action choice
        userChoice := InputBox(
            "Choose Action:`n`n1. Proceed with terminal`n2. Hit Enter`n3. Allow`n`nEnter choice (1-3):",
            "Cursor Action Selection", "w250 h180")
        if userChoice.Result != "OK"
            return

        ; First activate Cursor
        SetTitleMatchMode 2
        if WinExist("ahk_exe Cursor.exe") {
            WinActivate
            ; Wait for Cursor to be active
            WinWaitActive("ahk_exe Cursor.exe", "", 2)
        } else {
            ; Launch Cursor if not running
            target := IS_WORK_ENVIRONMENT ? "C:\\Users\\fie7ca\\AppData\\Local\\Programs\\cursor\\Cursor.exe" :
                "C:\\Users\\eduev\\AppData\\Local\\Programs\\cursor\\Cursor.exe"
            Run target
            WinWaitActive("ahk_exe Cursor.exe", "", 10)
        }

        ; Small delay to ensure Cursor is ready
        Sleep 200

        ; Send the key sequence based on user choice
        switch userChoice.Value {
            case "1":
            {
                ; Option 1: Proceed with terminal (original sequence)
                Send "^+e"   ; Press Ctrl+Shift+E
                Sleep 100
                Send "^i"    ; Ctrl+I
                Sleep 100
                Send "+{Backspace}"  ; Shift+Backspace
            }
            case "2":
            {
                ; Option 2: Hit Enter (modified sequence)
                Send "^+e"   ; Press Ctrl+Shift+E
                Sleep 100
                Send "^i"    ; Ctrl+I
                Sleep 100
                Send "{Enter}"  ; Enter instead of Shift+Backspace
            }
            case "3":
            {
                ; Option 3: Allow (basic sequence + up arrows + enter)
                Send "^+e"   ; Press Ctrl+Shift+E
                Sleep 100
                Send "^i"    ; Ctrl+I
                Sleep 100
                Send "{Up}"  ; up arrow
                Sleep 100
                Send "{Enter}"  ; Enter
            }
            default:
                MsgBox "Invalid selection. Please choose 1-3.", "Cursor Action Selection", "IconX"
                return
        }

    } catch Error as e {
        MsgBox "Error executing Cursor action: " e.Message, "Cursor Action Error", "IconX"
    }
}

#!+C::
{
    CursorKeySequence()
}

; =============================================================================
; Activate Hunt and Peck
; Hotkey: Win+Alt+Shift+X
; Original File: Hunt and Peck.ahk
; Note: A short press activates Hunt and Peck. A long press (>1s) will activate
;       loop mode, which will continuously reactivate Hunt and Peck after each
;       selection until either long pressed again or max iterations reached.
; =============================================================================

; Global variables for Hunt and Peck loop mode
; Path to Hunt and Peck executable – adjust if you install it elsewhere
#Include %A_ScriptDir%\env.ahk
global g_HnPExePath := GetHnPExePath()
global g_HnPLoopActive := false
global g_HnPLoopGui := false
global g_HnPMaxIterations := 25
global g_HnPCurrentIteration := 0
global g_HnPTargetWindow := 0  ; Store the window handle
global g_HnPRetryCount := 0    ; Track retry attempts for Program Manager recovery

; -----------------------------------------------------------------------------
; Helper: Force-close any running Hunt-and-Peck (hap.exe) processes
; -----------------------------------------------------------------------------
CloseHuntAndPeckProcess() {
    ; Attempt to terminate every instance of hap.exe. Ignoring any errors keeps
    ; the call simple and side-effect-free if the process isn’t running.
    try ProcessClose("hap.exe")
}

; Helper: returns true if active window is Program Manager (desktop) or taskbar – i.e. Hunt-and-Peck anchored incorrectly
IsBadHnPAnchor() {
    return WinActive("Program Manager") || WinActive("ahk_class Shell_TrayWnd")
}

; Safely activate the stored target window if it still exists
SafeActivateTarget(hwnd) {
    if (hwnd && WinExist("ahk_id " hwnd))
        WinActivate("ahk_id " hwnd)
}

; Send a quick right-click to the centred mouse position – this shifts focus to the window's
; main area without selecting items.  Any context menu will be dismissed automatically by
; Hunt-and-Peck's overlay / Esc logic.
RightClickFocus() {
    Click "Right"
    if (IS_WORK_ENVIRONMENT) {
        Sleep 10  ; Reduced sleep for work environment
    } else {
        Sleep 10  ; Personal environment now matches work environment
    }
}

; Shows or hides the loop mode indicator
ShowLoopIndicator(show := true) {
    global g_HnPLoopGui

    if (show && !g_HnPLoopGui) {
        g_HnPLoopGui := Gui()
        g_HnPLoopGui.Opt("+AlwaysOnTop -Caption +ToolWindow")
        g_HnPLoopGui.BackColor := "33AA33"
        g_HnPLoopGui.Add("Text", , "Hunt && Peck Loop Mode Active")
        g_HnPLoopGui.Show("NoActivate y0")
    } else if (!show && g_HnPLoopGui) {
        g_HnPLoopGui.Destroy()
        g_HnPLoopGui := false
    }
}

; Activates Hunt and Peck and handles potential Program Manager activation
ActivateHuntAndPeck(isLoopMode := false) {
    global g_HnPTargetWindow, g_HnPRetryCount, g_HnPExePath

    ; NOTE: Do NOT force-kill hap.exe here – it interferes with subsequent launches.

    ; For single activation (not loop mode), always use current window
    if (!isLoopMode) {
        g_HnPTargetWindow := WinExist("A")
    }

    ; If we don't have a valid window, we can't proceed
    if (!g_HnPTargetWindow || !WinExist("ahk_id " g_HnPTargetWindow)) {
        return false
    }

    ; Ensure our target window is active
    SafeActivateTarget(g_HnPTargetWindow)
    Sleep 40

    ; Shift keyboard focus with a harmless right-click
    RightClickFocus()
    if (IS_WORK_ENVIRONMENT) {
        Sleep 10  ; Reduced sleep for work environment
    } else {
        Sleep 10  ; Personal environment now matches work environment
    }

    ; Wait up to 200 ms for the window to actually become active (covers fast Alt-Tab cases)
    if !WinWaitActive("ahk_id " g_HnPTargetWindow, "", 0.2) {
        ; If it still isn't active, give up on this attempt
        return false
    }

    ; ----------------------------------------------------------------------
    ; Prefer the far more reliable CLI (hap.exe /hint). After running, ensure
    ; we did NOT end up focused on Program Manager **or** the taskbar.
    ; ----------------------------------------------------------------------
    boolSuccess := false
    if (FileExist(g_HnPExePath)) {
        try {
            ; Focus already fixed; just launch
            Run g_HnPExePath " /hint", , "Hide"
            Sleep 120
            boolSuccess := !IsBadHnPAnchor()
        }
    }

    ; If the CLI call failed (or exe not found) fall back to the legacy hotkey
    if (!boolSuccess) {
        ; Legacy hotkey path
        RightClickFocus()
        Send "!ç"
        if (IS_WORK_ENVIRONMENT) {
            Sleep 40  ; Reduced sleep for work environment
        } else {
            Sleep 40  ; Personal environment now matches work environment
        }
        ; Re-activate target window to pull overlay back
        SafeActivateTarget(g_HnPTargetWindow)
        Sleep 40

        ; Still anchored to Program Manager / taskbar? => one retry only.
        if (IsBadHnPAnchor()) {
            if (g_HnPRetryCount < 1) {
                g_HnPRetryCount++

                SafeActivateTarget(g_HnPTargetWindow)
                Sleep 40
                RightClickFocus()
                Sleep 80

                if (IsBadHnPAnchor()) {
                    g_HnPRetryCount := 0
                    return false
                }
            } else {
                g_HnPRetryCount := 0
                return false
            }
        }
    }

    ; Success – reset retry counter
    g_HnPRetryCount := 0
    return true
}

; Handles the Hunt and Peck loop mode
HnPLoopMode() {
    global g_HnPLoopActive, g_HnPCurrentIteration, g_HnPTargetWindow

    if (g_HnPLoopActive) {
        ; Stop the loop mode
        g_HnPLoopActive := false
        g_HnPCurrentIteration := 0
        ShowLoopIndicator(false)

        ; Stop the timer immediately
        SetTimer ActivateHnP, 0

        ; Brief delay to ensure any pending Hunt and Peck activation is complete
        Sleep 100

        ; Ensure we're in the target window and clear any Hunt and Peck state
        if (g_HnPTargetWindow && WinExist("ahk_id " g_HnPTargetWindow)) {
            SafeActivateTarget(g_HnPTargetWindow)
            Sleep 30
            Send "{Esc}"

            ; Safeguard: after 1000 ms send Esc to dismiss any late overlay.
            SetTimer(() => Send("{Esc}"), -1000)

            ; Ensure the Hunt-and-Peck process is fully terminated
            CloseHuntAndPeckProcess()
        }

        ; Enhanced cleanup: Check for remaining HAP.EXE instances multiple times within 2 seconds
        ; and close them if found. This ensures all instances are properly terminated.
        SetTimer(() => CloseHuntAndPeckProcess(), -500)   ; First check at 500ms
        SetTimer(() => CloseHuntAndPeckProcess(), -1000)  ; Second check at 1000ms
        SetTimer(() => CloseHuntAndPeckProcess(), -1500)  ; Third check at 1500ms

        g_HnPTargetWindow := 0
        return
    }

    ; Store the current active window
    g_HnPTargetWindow := WinExist("A")
    if (!g_HnPTargetWindow) {
        return  ; Don't show error, just fail silently
    }

    ; Start the loop mode
    g_HnPLoopActive := true
    g_HnPCurrentIteration := 0
    ShowLoopIndicator(true)

    ; Activate Hunt and Peck immediately for the first time
    ActivateHuntAndPeck()

    ; Define customizable interval (ms) for Hunt and Peck loop
    static loopIntervalMs := 3000   ; was 2000 – extended for more selection time

    ; Start the loop with the new interval
    SetTimer ActivateHnP, loopIntervalMs
}

; Timer function to activate Hunt and Peck
ActivateHnP() {
    global g_HnPLoopActive, g_HnPCurrentIteration, g_HnPMaxIterations, g_HnPTargetWindow

    if (!g_HnPLoopActive || g_HnPCurrentIteration >= g_HnPMaxIterations) {
        SetTimer ActivateHnP, 0  ; Stop the timer
        g_HnPLoopActive := false
        g_HnPCurrentIteration := 0
        g_HnPTargetWindow := 0
        ShowLoopIndicator(false)

        ; Ensure any residual Hunt-and-Peck overlay is cleared
        Send "{Esc}"
        SetTimer(() => Send("{Esc}"), -1000)

        ; Terminate any lingering hap.exe process
        CloseHuntAndPeckProcess()

        ; Enhanced cleanup: Check for remaining HAP.EXE instances multiple times within 2 seconds
        ; and close them if found. This ensures all instances are properly terminated.
        SetTimer(() => CloseHuntAndPeckProcess(), -500)   ; First check at 500ms
        SetTimer(() => CloseHuntAndPeckProcess(), -1000)  ; Second check at 1000ms
        SetTimer(() => CloseHuntAndPeckProcess(), -1500)  ; Third check at 1500ms

        return
    }

    g_HnPCurrentIteration++
    if (!ActivateHuntAndPeck(true)) {  ; Pass true to indicate loop mode
        ; If Hunt and Peck activation failed after retry, stop the loop
        SetTimer ActivateHnP, 0
        g_HnPLoopActive := false
        g_HnPCurrentIteration := 0
        g_HnPTargetWindow := 0
        ShowLoopIndicator(false)

        ; Enhanced cleanup: Check for remaining HAP.EXE instances multiple times within 2 seconds
        ; and close them if found. This ensures all instances are properly terminated.
        SetTimer(() => CloseHuntAndPeckProcess(), -500)   ; First check at 500ms
        SetTimer(() => CloseHuntAndPeckProcess(), -1000)  ; Second check at 1000ms
        SetTimer(() => CloseHuntAndPeckProcess(), -1500)  ; Third check at 1500ms
    }
}

; -----------------------------------------------------------------------------
; Helper: Schedule multiple attempts to close Hunt-and-Peck after single activation
;           Provides redundancy in case the first attempt fails or HnP hangs.
; -----------------------------------------------------------------------------
ScheduleHnPCleanup() {
    global g_HnPLoopActive
    ; Execute CloseHuntAndPeckProcess() after progressive delays, but only if we’re not in loop mode
    static delays := [3000, 7000, 12000]  ; milliseconds
    for d in delays
        SetTimer(() => (!g_HnPLoopActive ? CloseHuntAndPeckProcess() : ""), -d)
}

#!+x::
{
    ; Treat press-and-hold >400 ms as loop-mode trigger (keeps parity with keyboard firmware)
    if KeyWait("x", "T0.4") {
        ; Released within 400 ms → single activation
        ActivateHuntAndPeck(false)  ; Pass false to indicate single activation
        ; Schedule redundant cleanup attempts to ensure hap.exe is terminated
        ScheduleHnPCleanup()
    }
    else {
        ; Key was held down ≥400 ms (long press)
        KeyWait("x")  ; Wait for the key to be released
        HnPLoopMode()
    }
}
