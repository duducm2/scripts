#Requires AutoHotkey v2.0+
#SingleInstance Force

; -----------------------------------------------------------------------------
; This script consolidates all Microsoft Teams related hotkeys and functions.
; -----------------------------------------------------------------------------

; --- Includes ----------------------------------------------------------------
#include UIA-v2\Lib\UIA.ahk

; --- Helper Functions --------------------------------------------------------

ActivateTeamsMeetingWindow() {
    static processes := ["ms-teams.exe", "Teams.exe", "MSTeams.exe"]
    for proc in processes {
        for hwnd in WinGetList("ahk_exe " proc) {
            if IsTeamsMeetingTitle(title := WinGetTitle(hwnd)) {
                WinActivate(hwnd)
                return true
            }
        }
    }
    if hwnd := WinExist("RegEx)^.*\| Microsoft Teams$") {
        WinActivate(hwnd)
        return true
    }
    MsgBox "Não foi possível encontrar uma janela de reunião ativa.", "Microsoft Teams", "Iconi"
    return false
}

ActivateTeamsChatWindow() {
    static processes := ["ms-teams.exe", "Teams.exe", "MSTeams.exe"]
    for proc in processes {
        for hwnd in WinGetList("ahk_exe " proc) {
            if IsTeamsChatTitle(title := WinGetTitle(hwnd)) {
                WinActivate(hwnd)
                return true
            }
        }
    }
    if hwnd := WinExist("RegEx)^Chat \| .* \| Microsoft Teams$") {
        WinActivate(hwnd)
        return true
    }
    ; No message box here - just return false
    return false
}

FindListItemContaining(root, text) {
    items := root.FindAll(UIA.CreateCondition({ ControlType: "ListItem" }))
    for item in items {
        if InStr(item.Name, text)
            return item
    }
    return false
}

WaitListItem(root, partialName, timeout := 3000) {
    start := A_TickCount
    while (A_TickCount - start < timeout) {
        item := FindListItemContaining(root, partialName)
        if item
            return item
        Sleep 100
    }
    return false
}

IsTeamsMeetingTitle(title) {
    if InStr(title, "Chat |") || InStr(title, "Sharing control bar |")
        return false
    if InStr(title, "Microsoft Teams meeting")
        return true
    return RegExMatch(title, "i)^.*\| Microsoft Teams.*$")
}

IsTeamsChatTitle(title) {
    if InStr(title, "Sharing control bar |") || InStr(title, "Microsoft Teams meeting")
        return false
    return InStr(title, "Chat |") && RegExMatch(title, "i)\| Microsoft Teams$")
}

; --- NEW helper --------------------------------------------------------------
ShowCenteredOverlay(hwndTarget, text, duration := 1500) {
    ; Build a semi-transparent dark overlay with big white text
    WinGetPos(&wx, &wy, &ww, &wh, hwndTarget)
    ov := Gui("+AlwaysOnTop -Caption +ToolWindow")
    ov.BackColor := "333333"          ; dark background
    ov.SetFont("s24 cFFFFFF", "Segoe UI")
    msg := ov.Add("Text", "Center", text)
    ov.Show("AutoSize Hide")          ; measure the GUI first
    ov.GetPos(&gx, &gy, &gw, &gh)
    ; center over target window
    cx := wx + (ww - gw)//2
    cy := wy + (wh - gh)//2
    ov.Show("x" . cx . " y" . cy . " NA")
    WinSetTransparent(220, ov)        ; a bit of transparency
    Sleep duration
    ov.Destroy()
}


; --- Hotkeys & Functions -----------------------------------------------------

; =============================================================================
; Meeting: Toggle Mute
; Hotkey: Win+Alt+Shift+5
; Original File: Microsoft Teams - meeting shortcuts.ahk
; =============================================================================
#!+5:: {
    prev := WinGetID("A")                     ; window you were in
    if ActivateTeamsMeetingWindow() {
        Send "^+m"
        Sleep 100
        WinActivate(prev)                     ; switch back
        ShowCenteredOverlay(prev, "MIC MUTED")
    }
}

; =============================================================================
; Meeting: Toggle Camera
; Hotkey: Win+Alt+Shift+4
; Original File: Microsoft Teams - meeting shortcuts.ahk
; =============================================================================
#!+4:: {
    prev := WinGetID("A")
    if ActivateTeamsMeetingWindow() {
        Send "^+o"
        Sleep 100
        WinActivate(prev)
        ShowCenteredOverlay(prev, "CAMERA OFF")
    }
}


; =============================================================================
; Meeting: Toggle Screen Share  (Win Alt Shift T)
; =============================================================================
#!+t:: {
    prev := WinGetID("A")                 ; remember the window you were in
    if !ActivateTeamsMeetingWindow()
        return

    hwndTeams := WinGetID("A")            ; Teams meeting window
    root := UIA.ElementFromHandle(hwndTeams)
    if !root
        return

    ; --- perform the normal sharing workflow ---
    listItem := FindListItemContaining(root, "Opens list of")
    if listItem {
        listItem.Invoke()
    } else {
        shareBtn := root.FindFirst(UIA.CreateCondition({ AutomationId: "share-button" }))
        if !shareBtn
            return
        shareBtn.Invoke()
        Sleep 1000
        if li := WaitListItem(root, "Opens list of")
            li.Invoke()
    }

    ; --- show the overlay on the PREVIOUS window ---
    ShowCenteredOverlay(prev, "SHARING STARTED")
}

; =============================================================================
; Meeting: Exit Meeting
; Hotkey: Win+Alt+Shift+2
; Original File: Microsoft Teams - meeting shortcuts.ahk
; =============================================================================
#!+2:: {
    if !ActivateTeamsMeetingWindow()
        return
    response := MsgBox("Tem certeza de que deseja sair da reunião?", "Sair da reunião?", "YesNo Icon!")
    if response = "Yes"
        Send "^+h"
}

; =============================================================================
; Activate Chat Window
; Hotkey: Win+Alt+Shift+E
; Original File: Microsoft Teams - meeting shortcuts.ahk
; =============================================================================
#!+E:: {
    if !ActivateTeamsChatWindow() {
        RunTeams()
    }
}

RunTeams() {
    ; Example for Microsoft Store Teams
    ; Run("shell:AppsFolder\MicrosoftTeams_8wekyb3d8bbwe!App")
    
    ; Example for desktop Teams
    Run("c:\Users\fie7ca\Documents\Atalhos\Microsoft Teams - Shortcut.lnk")
}

; =============================================================================
; Activate Meeting Window
; Hotkey: Win+Alt+Shift+3
; Original File: Microsoft Teams - meeting shortcuts.ahk
; =============================================================================
#!+3:: {
    if !ActivateTeamsMeetingWindow()
        MsgBox "Não foi possível encontrar uma janela de reunião ativa.", "Microsoft Teams", "Iconi"
}

; =============================================================================
; Start New Conversation
; Hotkey: Win+Alt+Shift+R
; Original File: Microsoft Teams - New conversation.ahk
; =============================================================================
#!+r::
{
    contact := Trim(InputBox("Enter a Teams contact name:", "Jump to Chat").Value)
    if contact = ""
        return
    SetWinDelay 0
    SetKeyDelay 0, 0
    SetControlDelay 0
    teamsWindow := "Microsoft Teams"
    if !WinExist("ahk_exe ms-teams.exe") && !WinExist("ahk_exe Teams.exe") {
        Run "ms-teams:"
        WinWait(teamsWindow, , 15)
    }
    WinActivate(teamsWindow)
    WinWaitActive(teamsWindow, , 5)
    Send "^g"
    Sleep 100
    SendText(contact)
    Sleep 1300
    Send "{Enter}"
    Sleep 300
    Send "^r"
}
