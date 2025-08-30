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
; Open/Activate OneNote
; Hotkey: Win+Alt+Shift+N
; Original File: OneNote - Open.ahk
; =============================================================================
#!+n::
{
    if WinExist("ahk_exe ONENOTE.EXE") {
        WinActivate
        CenterMouse()
    } else {
        Run "c:\ProgramData\Microsoft\Windows\Start Menu\Programs\OneNote.lnk"
        WinWaitActive("ahk_exe ONENOTE.EXE")
        CenterMouse()
    }
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
        Run "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Wikipedia.lnk"
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
    Sleep(3000)
    Send("^w")

    ; Open Windows Clock (Alarms & Clock) and start focus session (Pomodoro)
    try {
        Run "ms-clock:"
        ; The Clock window is hosted by ApplicationFrameHost.exe, title typically contains "Relógio" or "Clock"
        SetTitleMatchMode 2
        WinWait("ahk_exe ApplicationFrameHost.exe")
        WinActivate("ahk_exe ApplicationFrameHost.exe")
        Sleep(2000)

        Send("^{Home}")
        Sleep(100)
        Send("{Enter}")

        ; Only proceed if previous Send("{Enter}") (after ^{Home}) succeeded
        Sleep(600)
        loop 9 {
            Send("{Tab}")
            Sleep(100)
        }
        Send("{Enter}")

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
