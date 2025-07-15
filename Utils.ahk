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

; Activates Hunt and Peck and handles potential Program Manager activation
ActivateHuntAndPeck(isLoopMode := false) {
    global g_HnPTargetWindow, g_HnPRetryCount

    ; For single activation (not loop mode), always use current window
    if (!isLoopMode) {
        g_HnPTargetWindow := WinExist("A")
    }

    ; If we don't have a valid window, we can't proceed
    if (!g_HnPTargetWindow || !WinExist("ahk_id " g_HnPTargetWindow)) {
        return false
    }

    ; Ensure our target window is active
    WinActivate("ahk_id " g_HnPTargetWindow)
    Sleep 30  ; Brief delay for window activation

    ; Wait up to 120 ms for the window to actually become active (covers fast Alt-Tab cases)
    if !WinWaitActive("ahk_id " g_HnPTargetWindow, "", 0.12) {
        ; If it still isn't active, give up on this attempt
        return false
    }

    ; Pre-action: double middle-click (press wheel twice)
    Click "Middle"
    Sleep 40
    Click "Middle"

    ; Now activate Hunt and Peck hotkey sequence
    Send "!ç"
    Sleep 50  ; Brief delay to let Hunt and Peck activate

    ; Check if we lost focus to Program Manager
    if (WinActive("Program Manager")) {
        if (g_HnPRetryCount < 1) {  ; Only try recovery once
            g_HnPRetryCount++

            ; Simply reactivate our target window and try again
            WinActivate("ahk_id " g_HnPTargetWindow)
            Sleep 30
            Send "!ç"

            ; Check if we succeeded this time
            Sleep 50
            if (WinActive("Program Manager")) {
                g_HnPRetryCount := 0
                return false
            }
        } else {
            g_HnPRetryCount := 0
            return false
        }
    }

    ; Success - reset retry counter
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
            WinActivate("ahk_id " g_HnPTargetWindow)
            Sleep 30
            Send "{Esc}"
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
    if KeyWait("x", "T1") {
        ; Key was released within 1 second (short press)
        ActivateHuntAndPeck(false)  ; Pass false to indicate single activation
    }
    else {
        ; Key was held down for >1 second (long press)
        KeyWait("x")  ; Wait for the key to be released
        HnPLoopMode()
    }
}
