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
global g_HnPLoopActive := false
global g_HnPLoopGui := false
global g_HnPMaxIterations := 25
global g_HnPCurrentIteration := 0
global g_HnPTargetWindow := 0  ; Store the window handle
global g_HnPRetryCount := 0    ; Track retry attempts for Program Manager recovery

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

; Try to recover the last active window
RecoverActiveWindow() {
    ; First try Alt+Tab to get back to the last window
    Send "!{Tab}"
    Sleep 30  ; Brief delay to let window activate

    ; If we're still on Program Manager, try Escape and Alt+Tab again
    if (WinActive("Program Manager") || WinActive("ahk_exe explorer.exe")) {
        Send "{Esc}"
        Sleep 30
        Send "!{Tab}"
        Sleep 30
    }

    return !WinActive("Program Manager") && !WinActive("ahk_exe explorer.exe")
}

; Activates Hunt and Peck and handles potential Program Manager activation
ActivateHuntAndPeck() {
    global g_HnPTargetWindow, g_HnPRetryCount

    ; Activate the target window before sending Hunt and Peck
    if (g_HnPTargetWindow && WinExist("ahk_id " g_HnPTargetWindow)) {
        WinActivate("ahk_id " g_HnPTargetWindow)
    } else {
        ; If target window is lost, try to recover the last active window
        if (!RecoverActiveWindow()) {
            return false
        }
        ; Update target window to the recovered window
        g_HnPTargetWindow := WinExist("A")
    }

    Sleep 30  ; Brief delay for window activation
    Send "!ç"  ; Activate Hunt and Peck

    ; Quick check for Program Manager activation
    Sleep 50

    ; Check if Program Manager got activated
    if (WinActive("Program Manager") || WinActive("ahk_exe explorer.exe")) {
        if (g_HnPRetryCount < 1) {  ; Only try recovery once
            g_HnPRetryCount++
            ; Try to recover the window and Hunt and Peck
            if (RecoverActiveWindow()) {
                Sleep 30
                Send "!ç"  ; Try Hunt and Peck again
            } else {
                g_HnPRetryCount := 0
                return false
            }
        } else {
            ; If we already tried recovery once, reset and stop
            g_HnPRetryCount := 0
            return false
        }
    } else {
        ; Success - reset retry counter
        g_HnPRetryCount := 0
    }
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
            WinActivate("ahk_id " g_HnPTargetWindow)
            Sleep 30
            Send "{Esc}"
        } else {
            RecoverActiveWindow()
            Send "{Esc}"
        }

        g_HnPTargetWindow := 0
        return
    }

    ; Store the current active window
    g_HnPTargetWindow := WinExist("A")
    if (!g_HnPTargetWindow) {
        if (!RecoverActiveWindow()) {
            MsgBox "Cannot find active window. Cannot start Hunt and Peck loop mode."
            return
        }
        g_HnPTargetWindow := WinExist("A")
    }

    ; Start the loop mode
    g_HnPLoopActive := true
    g_HnPCurrentIteration := 0
    ShowLoopIndicator(true)

    ; Activate Hunt and Peck immediately for the first time
    ActivateHuntAndPeck()

    ; Start the loop - 2 seconds interval for subsequent activations
    SetTimer ActivateHnP, 2000
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
        return
    }

    g_HnPCurrentIteration++
    if (!ActivateHuntAndPeck()) {
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
    if KeyWait("x", "T1") {
        ; Key was released within 1 second (short press)
        ActivateHuntAndPeck()  ; Use the same activation function for consistency
    }
    else {
        ; Key was held down for >1 second (long press)
        KeyWait("x")  ; Wait for the key to be released
        HnPLoopMode()
    }
}
