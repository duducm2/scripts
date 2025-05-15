;─────────────────────────────────────────────────────────────────────────────
;  ChatGPT Dictation → Transcript → Grammar-correction macro
;  Alt + Shift + B  to start / stop dictation
;─────────────────────────────────────────────────────────────────────────────
#Requires AutoHotkey v2.0
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

;─────────────────────────────────────────────────────────────────────────────
F12 & 3:: {                                                 ; Alt + Shift + B
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
    ; When stopping, the button might be the "submit" one or the "stop dictation/transcribing" one
    submitOrStopNames_to_find := [currentSubmitName, currentTranscribingName]

    ; Activate ChatGPT window
    SetTitleMatchMode 2
    WinActivate "chatgpt - transcription"
    WinWaitActive "ahk_exe chrome.exe"
    cUIA := UIA_Browser()
    Sleep 300

    ; UI labels (Portuguese only) ; These specific lines are now replaced by the dynamic logic above
    ; dictateNames := ["Botão de ditado"]
    ; submitNames := ["Enviar ditado"]

    ;───────────────────────────── START DICTATION ───────────────────────────────
    if !isDictating {
        try {
            if btn := FindButton(cUIA, dictateNames_to_find) {
                btn.Click()
                isDictating := true
            } else
                MsgBox (IS_WORK_ENVIRONMENT ? "'" . currentDictateName . "' não encontrado." : "'" . currentDictateName .
                    "' not found.")
        } catch as e {
            MsgBox (IS_WORK_ENVIRONMENT ? "Erro ao iniciar ditado:`n" : "Error starting dictation:`n") e.Message
            isDictating := false
        }
        return
    }

    ;───────────────────────────── STOP DICTATION ────────────────────────────────
    try {
        if btn := FindButton(cUIA, submitOrStopNames_to_find) { ; Use the array with both possibilities
            btn.Click()                 ; stop recording / start transcription
            isDictating := false
        } else {
            MsgBox (IS_WORK_ENVIRONMENT ? "Botão para parar/enviar ditado ('" . currentSubmitName . "' ou '" .
                currentTranscribingName . "') não encontrado." : "Button to stop/submit dictation ('" .
                    currentSubmitName . "' or '" . currentTranscribingName . "') not found.")
        }
    } catch as e {
        MsgBox (IS_WORK_ENVIRONMENT ? "Erro ao finalizar ditado:`n" : "Error stopping dictation:`n") e.Message
        isDictating := false
    }

}
