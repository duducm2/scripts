#Requires AutoHotkey v2.0+
#SingleInstance Force

; Assuming UIA-v2 folder is now a subfolder in the same directory as this script.
#include UIA-v2\\Lib\\UIA.ahk
#include UIA-v2\\Lib\\UIA_Browser.ahk
#include %A_ScriptDir%\\env.ahk ; Include environment configuration

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

; Win+Alt+Shift+3 to toggle microphone
#!+0::
{
    ToggleDictation()
}

ToggleDictation(triedFallback := false, forceAction := "") {
    static isDictating := false

    ; Define button names for both languages
    pt_dictateName := "Botão de ditado"
    en_dictateName := "Dictate button"
    pt_submitName := "Enviar ditado"
    en_submitName := "Submit dictation"
    pt_transcribingName := "Interromper ditado"
    en_transcribingName := "Stop dictation"

    ; Select names based on environment (IS_WORK_ENVIRONMENT is true for Work/Portuguese)
    currentDictateName := IS_WORK_ENVIRONMENT ? pt_dictateName : en_dictateName
    currentSubmitName := IS_WORK_ENVIRONMENT ? pt_submitName : en_submitName
    currentTranscribingName := IS_WORK_ENVIRONMENT ? pt_transcribingName : en_transcribingName

    ; Prepare arrays for FindButton
    dictateNames_to_find := [currentDictateName]
    submitOrStopNames_to_find := [currentSubmitName, currentTranscribingName]

    ; Activate ChatGPT window
    SetTitleMatchMode 2
    WinActivate "chatgpt"
    WinWaitActive "ahk_exe chrome.exe"
    cUIA := UIA_Browser()
    Sleep 300

    ; Determine action: start or stop dictation
    action := forceAction ? forceAction : (!isDictating ? "start" : "stop")

    if (action = "start") {
        try {
            if btn := FindButton(cUIA, dictateNames_to_find) {
                btn.Click()
                isDictating := true
                return
            } else if !triedFallback {
                ; Try the opposite action (stop) if start failed
                ToggleDictation(true, "stop")
                return
            } else {
                MsgBox (IS_WORK_ENVIRONMENT ? "Não foi possível iniciar ou parar o ditado. Nenhum botão encontrado." :
                    "Could not start or stop dictation. No button found.")
            }
        } catch as e {
            if !triedFallback {
                ToggleDictation(true, "stop")
                return
            } else {
                MsgBox (IS_WORK_ENVIRONMENT ? "Erro ao iniciar/parar ditado:`n" :
                    "Error starting/stopping dictation:`n") e.Message
            }
        }
    } else if (action = "stop") {
        try {
            if btn := FindButton(cUIA, submitOrStopNames_to_find) {
                btn.Click()                 ; stop recording / start transcription
                isDictating := false
                return
            } else if !triedFallback {
                ; Try the opposite action (start) if stop failed
                ToggleDictation(true, "start")
                return
            } else {
                MsgBox (IS_WORK_ENVIRONMENT ? "Não foi possível parar ou iniciar o ditado. Nenhum botão encontrado." :
                    "Could not stop or start dictation. No button found.")
            }
        } catch as e {
            if !triedFallback {
                ToggleDictation(true, "start")
                return
            } else {
                MsgBox (IS_WORK_ENVIRONMENT ? "Erro ao parar/iniciar ditado:`n" :
                    "Error stopping/starting dictation:`n") e.Message
            }
        }
    }
}
