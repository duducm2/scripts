#Requires AutoHotkey v2.0+
#SingleInstance Force

; Assuming UIA-v2 folder is now a subfolder in the same directory as this script.
#include UIA-v2\\Lib\\UIA.ahk
#include UIA-v2\\Lib\\UIA_Browser.ahk

;─────────────────────────────────────────────────────────────────────────────
; Helper: find the first UIA element whose Name matches any string in array
FindButton(cUIA, names, role := "Button", timeoutMs := 0) {
    for name in names {
        el := (timeoutMs = 0)
            ? cUIA.FindElement({ Name: name, Type: role })
            : cUIA.WaitElement({ Name: name, Type: role }, timeoutMs)
        if el
            return el
    }
    return false
}

; Win+Alt+Shift+L to toggle voice mode
#!+l::
{
    ToggleVoiceMode()
}

ToggleVoiceMode(triedFallback := false, forceAction := "") {
    static isVoiceModeActive := false

    ; Define button names
    startVoiceName := "Start voice mode"
    endVoiceName := "End voice mode"

    ; Prepare arrays for FindButton
    startNames_to_find := [startVoiceName]
    endNames_to_find := [endVoiceName]

    ; Activate ChatGPT window
    SetTitleMatchMode 2
    WinActivate "chatgpt"
    WinWaitActive "ahk_exe chrome.exe"
    cUIA := UIA_Browser()
    Sleep 300

    ; Determine action: start or stop voice mode
    action := forceAction ? forceAction : (!isVoiceModeActive ? "start" : "stop")

    if (action = "start") {
        try {
            if btn := FindButton(cUIA, startNames_to_find) {
                btn.Click()
                isVoiceModeActive := true
                return
            } else if !triedFallback {
                ; Try the opposite action (stop) if start failed
                ToggleVoiceMode(true, "stop")
                return
            } else {
                MsgBox "Could not start or stop voice mode. No button found."
            }
        } catch as e {
            if !triedFallback {
                ToggleVoiceMode(true, "stop")
                return
            } else {
                MsgBox "Error starting/stopping voice mode:`n" e.Message
            }
        }
    } else if (action = "stop") {
        try {
            if btn := FindButton(cUIA, endNames_to_find) {
                btn.Click()
                isVoiceModeActive := false
                return
            } else if !triedFallback {
                ; Try the opposite action (start) if stop failed
                ToggleVoiceMode(true, "start")
                return
            } else {
                MsgBox "Could not stop or start voice mode. No button found."
            }
        } catch as e {
            if !triedFallback {
                ToggleVoiceMode(true, "start")
                return
            } else {
                MsgBox "Error stopping/starting voice mode:`n" e.Message
            }
        }
    }
}
