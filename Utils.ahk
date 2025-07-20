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
        MsgBox "Still no active window!"
        return
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

; Send a quick right-click to the centred mouse position – this shifts focus to the window’s
; main area without selecting items.  Any context menu will be dismissed automatically by
; Hunt-and-Peck’s overlay / Esc logic.
RightClickFocus() {
    Click "Right"
    Sleep 30
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
    Sleep 20

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
        Sleep 80
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
    }
}

#!+x::
{
    ; Treat press-and-hold >400 ms as loop-mode trigger (keeps parity with keyboard firmware)
    if KeyWait("x", "T0.4") {
        ; Released within 400 ms → single activation
        ActivateHuntAndPeck(false)  ; Pass false to indicate single activation
    }
    else {
        ; Key was held down ≥400 ms (long press)
        KeyWait("x")  ; Wait for the key to be released
        HnPLoopMode()
    }
}
