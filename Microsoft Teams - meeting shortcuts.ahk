#Requires AutoHotkey v2.0+

#include UIA-v2\Lib\UIA.ahk

ActivateTeamsMeetingWindow() {
    static processes := ["ms-teams.exe", "Teams.exe", "MSTeams.exe"]
    for proc in processes {
        for hwnd in WinGetList("ahk_exe " proc) {
            title := WinGetTitle(hwnd)
            if (
                ( InStr(title, "Microsoft Teams meeting")
               || RegExMatch(title, "^[^|]+\| Microsoft Teams$")
               || RegExMatch(title, "^.*\(.*\) \| Microsoft Teams$") )
                && !InStr(title, "Chat |")
                && !InStr(title, "Sharing control bar |") ) {
                WinActivate(hwnd)
                return true
            }
        }
    }
    if hwnd := WinExist(".*\| Microsoft Teams") {
        WinActivate(hwnd)
        return true
    }
    MsgBox "Não foi possível encontrar uma janela de reunião ativa.", "Microsoft Teams", "Iconi"
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

#!+q:: {
    if ActivateTeamsMeetingWindow()
        Send "^+m"
}

#!+f:: {
    if ActivateTeamsMeetingWindow()
        Send "^+o"
}

#!+e:: {
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

#!+u:: {
    if !ActivateTeamsMeetingWindow()
        return
    response := MsgBox("Tem certeza de que deseja sair da reunião?", "Sair da reunião?", "YesNo Icon!")
    if response = "Yes"
        Send "^+h"
}
