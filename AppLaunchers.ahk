#Requires AutoHotkey v2.0+
#SingleInstance Force

; -----------------------------------------------------------------------------
; This script consolidates all Application/Website launcher hotkeys.
; -----------------------------------------------------------------------------

; --- Includes ----------------------------------------------------------------
#include %A_ScriptDir%\env.ahk
#include UIA-v2\Lib\UIA.ahk

; --- Hotkeys & Functions -----------------------------------------------------

; =============================================================================
; Open/Activate Cursor with specific window requirements
; Hotkey: Win+Alt+Shift+N
; Original File: OneNote - Open.ahk (Updated for Cursor)
; =============================================================================
#!+n::
{
    ; Look for Cursor windows with specific names: habits, home, punctual, or work
    ; Exclude windows containing "preview"
    targetWindow := ""
    fallbackWindow := ""

    ; Get all Cursor windows (support both Cursor.exe and Code.exe just in case)
    for proc in ["ahk_exe Cursor.exe", "ahk_exe Code.exe"] {
        for hwnd in WinGetList(proc) {
            try {
                winTitle := WinGetTitle("ahk_id " hwnd)
                winTitleLower := StrLower(winTitle)

                ; Skip any preview windows
                if InStr(winTitleLower, "preview")
                    continue

                ; Store the first valid Cursor window as fallback
                if (!fallbackWindow)
                    fallbackWindow := "ahk_id " hwnd

                ; Check if window contains any of the target names
                if (InStr(winTitleLower, "habits")
                || InStr(winTitleLower, "home")
                || InStr(winTitleLower, "punctual")
                || InStr(winTitleLower, "work")) {
                    targetWindow := "ahk_id " hwnd
                    break
                }
            } catch {
                ; Silently skip invalid windows
            }
        }
        if (targetWindow)
            break
    }

    if (targetWindow) {
        ; Found a matching Cursor window
        WinActivate(targetWindow)
        if WinWaitActive(targetWindow, , 2) {
            CenterMouse()
            Sleep(100)  ; Small delay to ensure window is fully active
            Send("^t")  ; Send Ctrl+Shift+O
        }
    } else if (fallbackWindow) {
        ; No specific window found, but found a general Cursor window - use fallback
        WinActivate(fallbackWindow)
        if WinWaitActive(fallbackWindow, , 2) {
            CenterMouse()
            Sleep(100)  ; Small delay to ensure window is fully active
            Send("^t")  ; Send Ctrl+Shift+O
        }
    } else {
        ; No Cursor window found at all - show fallback panel
        ShowCursorFallbackPanel()
    }
}

; =============================================================================
; Show fallback panel when no matching Cursor window is found
; =============================================================================
ShowCursorFallbackPanel() {
    ; Create a notification panel similar to ChatGPT.ahk style
    fallbackGui := Gui()
    fallbackGui.Opt("+AlwaysOnTop -Caption +ToolWindow")
    fallbackGui.BackColor := "3772FF"  ; Blue background like ChatGPT notifications
    fallbackGui.SetFont("s20 cFFFFFF Bold", "Segoe UI")

    ; Create the message
    message := "No Cursor window found with names: habits, home, punctual, or work"
    fallbackGui.Add("Text", "w600 Center", message)

    ; Center on active monitor
    activeWin := WinGetID("A")
    if (activeWin) {
        WinGetPos(&winX, &winY, &winW, &winH, activeWin)
    } else {
        MonitorGetWorkArea(1, &l, &t, &r, &b)
        winX := l
        winY := t
        winW := r - l
        winH := b - t
    }

    fallbackGui.Show("AutoSize Hide")
    guiW := 0, guiH := 0
    fallbackGui.GetPos(, , &guiW, &guiH)

    guiX := winX + (winW - guiW) / 2
    guiY := winY + (winH - guiH) / 2
    fallbackGui.Show("x" . Round(guiX) . " y" . Round(guiY) . " NA")
    WinSetTransparent(178, fallbackGui)

    ; Auto-hide after 3 seconds
    SetTimer(() => fallbackGui.Destroy(), -3000)
}

; =============================================================================
; Open/Activate Desktop in Explorer
; Hotkey: Shift+Win+E
; Original File: Open Desktop.ahk
; =============================================================================
+#e::
{
    SetTitleMatchMode 2
    if WinExist("Área de Trabalho ahk_class CabinetWClass") || WinExist("Desktop ahk_class CabinetWClass") {
        WinActivate
        WinWaitActive("ahk_class CabinetWClass", , 2)  ; Wait up to 2 seconds for activation
        Sleep(100)  ; Small additional delay to ensure window is ready
        CenterMouse()
    } else {
        target := IS_WORK_ENVIRONMENT ? "C:\Users\fie7ca\Desktop" : "C:\Users\eduev\OneDrive\Desktop"
        Run 'explorer.exe "' target '"'
        ; Wait until Explorer window appears AND becomes active
        WinWait("ahk_class CabinetWClass")
        WinWaitActive("ahk_class CabinetWClass", , 2)  ; Wait up to 2 seconds for activation
        Sleep(100)  ; Small additional delay to ensure window is ready
        CenterMouse()
    }
}

; =============================================================================
; Open/Activate Google Chrome
; Hotkey: Win+Alt+Shift+F
; Original File: Open Google.ahk
; =============================================================================
#!+f::
{
    Run "chrome.exe"
    WinWaitActive("ahk_exe chrome.exe")
    CenterMouse()
}

; =============================================================================
; Open/Activate WhatsApp
; Hotkey: Win+Alt+Shift+Z
; Original File: Open WhatsApp.ahk
; =============================================================================
#!+z::
{
    SetTitleMatchMode(2)
    if WinExist("WhatsApp") {
        WinActivate("WhatsApp")
        CenterMouse()
    } else {
        if (IS_WORK_ENVIRONMENT) {
            Run "C:\Users\fie7ca\Documents\Shortcuts\WhatsApp.lnk"
        } else {
            Run "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\apps do Chrome\WhatsApp Web.lnk"
        }
        WinWaitActive("WhatsApp")
        CenterMouse()
    }
}

; =============================================================================
; Open/Activate YouTube
; Hotkey: Win+Alt+Shift+H
; Original File: Youtube - Activate.ahk
; =============================================================================
#!+h::
{
    SetTitleMatchMode 2
    if WinExist("YouTube") {
        WinActivate
        CenterMouse()
    } else {
        Run "chrome.exe --new-window https://www.youtube.com"
        WinWaitActive("YouTube")
        CenterMouse()
    }
}

; =============================================================================
; Open/Activate Gmail
; Hotkey: Win+Alt+Shift+W
; =============================================================================
#!+w::
{
    SetTitleMatchMode 2
    if WinExist("Gmail ahk_exe chrome.exe") {
        WinActivate
        CenterMouse()
    } else {
        target := IS_WORK_ENVIRONMENT ?
            "C:\Users\fie7ca\Documents\Shortcuts\Gmail.lnk" :
                "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Chrome Apps\Gmail.lnk"
        Run target
        WinWaitActive("Gmail ahk_exe chrome.exe")
        CenterMouse()
    }
}

; =============================================================================
; Open/Activate Cursor
; Hotkey: Win+Alt+Shift+,
; =============================================================================
#!+,::
{
    SetTitleMatchMode 2
    if WinExist("ahk_exe Cursor.exe") {
        WinActivate
        CenterMouse()
    } else {
        target := IS_WORK_ENVIRONMENT ? "C:\\Users\\fie7ca\\AppData\\Local\\Programs\\cursor\\Cursor.exe" :
            "C:\\Users\\eduev\\AppData\\Local\\Programs\\cursor\\Cursor.exe"
        Run target
        WinWaitActive("ahk_exe Cursor.exe")
        CenterMouse()
    }
}

; =============================================================================
; Open/Activate Wikipedia
; Hotkey: Win+Alt+Shift+K
; =============================================================================
#!+k::
{
    SetTitleMatchMode 2
    if WinExist("Wikipedia") {
        WinActivate
        CenterMouse()
    } else {
        target := IS_WORK_ENVIRONMENT ? "C:\Users\fie7ca\Documents\Shortcuts\Wikipedia.lnk" :
            "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Wikipedia.lnk"
        Run target
        WinWaitActive("Wikipedia")
        CenterMouse()
    }
}

; =============================================================================
; Open Google Keep Pomodoro list, show overlay, then send keys
; Hotkey: Win+Alt+Shift+9
; =============================================================================
#!+9::
{
    SoundBeep()
    url := "https://keep.google.com/u/0/#LIST/18_9CP4EZyO6i88cnsRR6HgRNlSEohvkKX0PEZXh8mKnj8H_TpfRhG6aUbrCdxQ"
    Run "chrome.exe " url
    WinWait("ahk_exe chrome.exe")
    WinActivate("ahk_exe chrome.exe")

    imagePath := ""
    for name in ["tomato.png", "tomato.jpg", "pomodoro.png", "pomodoro.jpg"] {
        candidate := A_ScriptDir "\pictures\" name
        if FileExist(candidate) {
            imagePath := candidate
            break
        }
    }

    overlay := Gui()
    overlay.Opt("+AlwaysOnTop -Caption +ToolWindow +E0x20")
    if (imagePath != "") {
        overlay.Add("Picture", "w240 h240 Center", imagePath)
    } else {
        overlay.SetFont("s36", "Segoe UI")
        overlay.Add("Text", "Center cRed", "Pomodoro")
        overlay.BackColor := "FFFFFF"
    }
    overlay.Show("AutoSize Center")
    Sleep(5000)

    WinActivate("ahk_exe chrome.exe")
    Sleep(100)
    Send("{Enter}")
    Sleep(250)
    Send("{Up}")
    Sleep(250)
    ClipSaved := ClipboardAll()
    current := FormatTime(, "dd/MM/yyyy HH:mm")
    A_Clipboard := "tomato " current
    ClipWait(1)
    Send("^v")
    Sleep(50)
    A_Clipboard := ClipSaved

    Send("{Escape}")
    Sleep(4000)
    Send("^w")

    ; Open Windows Clock (Alarms & Clock) and start focus session (Pomodoro)
    try {
        Run "ms-clock:"
        ; The Clock window is hosted by ApplicationFrameHost.exe, title typically contains "Relógio" or "Clock"
        SetTitleMatchMode 2
        WinWait("ahk_exe ApplicationFrameHost.exe")
        WinActivate("ahk_exe ApplicationFrameHost.exe")
        Sleep(2000)

        ; Tab through elements until we find the "Iniciar" button, then press Enter
        try {
            ; Start from the beginning of the application
            Send("^{Home}")
            Sleep(300)

            ; Tab through elements until we find "Iniciar" button
            maxTabs := 50  ; Safety limit to prevent infinite loop
            foundButton := false

            loop maxTabs {
                ; Check if current focused element has AutomationId "TimerPlayPauseButton"
                try {
                    focusedElement := UIA.GetFocusedElement()
                    if (focusedElement && focusedElement.AutomationId == "TimerPlayPauseButton") {
                        foundButton := true
                        break
                    }
                } catch {
                    ; Continue if UIA check fails
                }

                ; Tab to next element
                Send("{Tab}")
                Sleep(200)
            }

            ; If we found the button, press Enter to activate it
            if (foundButton) {
                Send("{Enter}")
                Sleep(200)
            }

            ; Close the Clock window (ApplicationFrameHost.exe)
            try {
                WinClose("ahk_exe ApplicationFrameHost.exe")
            }

        } catch {
            ; Fallback: do nothing if UIA not available
        }

    } catch {
        ; Fail silently if UIA not available here
    }

    ; Quick motivation banner
    goOverlay := CreateCenteredBanner_Launchers("GO!", "3772FF", "FFFFFF", 48, 178)
    Sleep(1200)
    try goOverlay.Destroy()

    SoundBeep()
    overlay.Destroy()
}

; =============================================================================
; Helper function to center mouse on the active window
; =============================================================================
CenterMouse() {
    Sleep(200)
    Send("#!+q")
}

; =============================================================================
; Centered banner helper (AppLaunchers aesthetic)
; =============================================================================
CreateCenteredBanner_Launchers(message, bgColor := "be4747", fontColor := "FFFFFF", fontSize := 24, alpha := 178) {
    bGui := Gui()
    bGui.Opt("+AlwaysOnTop -Caption +ToolWindow")
    bGui.BackColor := bgColor
    bGui.SetFont("s" . fontSize . " c" . fontColor . " Bold", "Segoe UI")
    bGui.Add("Text", "w500 Center", message)

    activeWin := WinGetID("A")
    if (activeWin) {
        WinGetPos(&winX, &winY, &winW, &winH, activeWin)
    } else {
        workArea := SysGet.MonitorWorkArea(SysGet.MonitorPrimary)
        winX := workArea.Left, winY := workArea.Top, winW := workArea.Right - workArea.Left, winH := workArea.Bottom -
            workArea.Top
    }

    bGui.Show("AutoSize Hide")
    guiW := 0, guiH := 0
    bGui.GetPos(, , &guiW, &guiH)

    guiX := winX + (winW - guiW) / 2
    guiY := winY + (winH - guiH) / 2
    bGui.Show("x" . Round(guiX) . " y" . Round(guiY) . " NA")
    WinSetTransparent(alpha, bGui)
    return bGui
}
