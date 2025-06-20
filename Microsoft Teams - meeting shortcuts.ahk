#Requires AutoHotkey v2.0+

#include UIA-v2\Lib\UIA.ahk

ActivateTeamsMeetingWindow() {
    static processes := ["ms-teams.exe", "Teams.exe", "MSTeams.exe"]

    ; walk through every Teams window we can find
    for proc in processes {
        for hwnd in WinGetList("ahk_exe " proc) {
            if IsTeamsMeetingTitle(title := WinGetTitle(hwnd)) {
                WinActivate(hwnd)
                return true
            }
        }
    }

    ; fallback: any window whose title *matches the same regex*
    if hwnd := WinExist("RegEx)^.*\| Microsoft Teams$") {  ; use “RegEx)” prefix!
        WinActivate(hwnd)
        return true
    }

    MsgBox "Não foi possível encontrar uma janela de reunião ativa.", "Microsoft Teams", "Iconi"
    return false
}

ActivateTeamsChatWindow() {
    static processes := ["ms-teams.exe", "Teams.exe", "MSTeams.exe"]

    ; walk through every Teams window we can find
    for proc in processes {
        for hwnd in WinGetList("ahk_exe " proc) {
            if IsTeamsChatTitle(title := WinGetTitle(hwnd)) {
                WinActivate(hwnd)
                return true
            }
        }
    }

    ; fallback: any window whose title *matches the same regex*
    if hwnd := WinExist("RegEx)^Chat \| .* \| Microsoft Teams$") {
        WinActivate(hwnd)
        return true
    }

    MsgBox "Não foi possível encontrar uma janela de chat do Teams.", "Microsoft Teams", "Iconi"
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

; --- NEW: one place that decides if a Teams window is a meeting window ---
IsTeamsMeetingTitle(title) {
    ; filter out things we explicitly don’t want
    if InStr(title, "Chat |")              ; chat side-panel
    || InStr(title, "Sharing control bar |")
        return false

    ; ① legacy “Microsoft Teams meeting” pop-out
    if InStr(title, "Microsoft Teams meeting")
        return true

    ; ② anything (including multiple pipes) that ENDS with “| Microsoft Teams”
    ;     ^.*\| Microsoft Teams$
    return RegExMatch(title, "i)^.*\| Microsoft Teams$")
}

IsTeamsChatTitle(title) {
    ; ignore sharing control bar and explicit meeting windows
    if InStr(title, "Sharing control bar |")
    || InStr(title, "Microsoft Teams meeting")
        return false

    ; MUST contain "Chat |" and end with "| Microsoft Teams"
    return InStr(title, "Chat |") && RegExMatch(title, "i)\| Microsoft Teams$")
}

#!+5:: {
    if ActivateTeamsMeetingWindow()
        Send "^+m"
}

#!+4:: {
    if ActivateTeamsMeetingWindow()
        Send "^+o"
}

#!+t:: {
    if !ActivateTeamsMeetingWindow()
        return

    hwnd := WinGetID("A")
    root := UIA.ElementFromHandle(hwnd)
    if !root
        return

    listItem := FindListItemContaining(root, "Opens list of")
    if listItem {
        listItem.Invoke()
        ActivateTeamsMeetingWindow()
        return
    }

    shareBtn := root.FindFirst(UIA.CreateCondition({ AutomationId: "share-button" }))
    if !shareBtn
        return
    shareBtn.Invoke()

    Sleep 1000
    listItem := WaitListItem(root, "Opens list of")
    if listItem {
        listItem.Invoke()
    }
    Sleep 200
    ActivateTeamsMeetingWindow()
}

#!+2:: {
    if !ActivateTeamsMeetingWindow()
        return
    response := MsgBox("Tem certeza de que deseja sair da reunião?", "Sair da reunião?", "YesNo Icon!")
    if response = "Yes"
        Send "^+h"
}

; Win + Alt + Shift + J : Select the chats window (e.g. "Chat | GS/BDU UX Project Champions | Microsoft Teams")
#!+E:: {
    if !ActivateTeamsChatWindow()
        MsgBox "Não foi possível encontrar uma janela de chat do Teams.", "Microsoft Teams", "Iconi"
}

; Win + Alt + Shift + K : Select the meeting window (e.g. "Proj. Streamline Weekly Data Review | Microsoft Teams")
#!+3:: {
    if !ActivateTeamsMeetingWindow()
        MsgBox "Não foi possível encontrar uma janela de reunião ativa.", "Microsoft Teams", "Iconi"
}
