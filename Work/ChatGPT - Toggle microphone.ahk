;─────────────────────────────────────────────────────────────────────────────
;  ChatGPT Dictation → Transcript → Grammar-correction macro
;  Alt + Shift + B  to start / stop dictation
;─────────────────────────────────────────────────────────────────────────────
#Requires AutoHotkey v2.0
#SingleInstance Force

#include C:\Users\fie7ca\Documents\UIA-v2\Lib\UIA.ahk
#include C:\Users\fie7ca\Documents\UIA-v2\Lib\UIA_Browser.ahk

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
CapsLock & 3:: {                                                 ; Alt + Shift + B
    static isDictating := false

    ; Activate ChatGPT window
    SetTitleMatchMode 2
    WinActivate "chatgpt - transcription"
    WinWaitActive "ahk_exe chrome.exe"
    cUIA := UIA_Browser()
    Sleep 300

    ; UI labels (Portuguese only)
    dictateNames := ["Botão de ditado"]
    submitNames := ["Enviar ditado"]

    ;───────────────────────────── START DICTATION ───────────────────────────────
    if !isDictating {
        try {
            if btn := FindButton(cUIA, dictateNames) {
                btn.Click()
                isDictating := true
            } else
                MsgBox "‘Botão de ditado’ não encontrado."
        } catch as e {
            MsgBox "Erro ao iniciar ditado:`n" e.Message
            isDictating := false
        }
        return
    }

    ;───────────────────────────── STOP DICTATION ────────────────────────────────
    try {
        if btn := FindButton(cUIA, submitNames) {
            btn.Click()                 ; stop recording / start transcription
            isDictating := false
        } else {
            MsgBox "‘Enviar ditado’ não encontrado."
        }
    } catch as e {
        MsgBox "Erro ao finalizar ditado:`n" e.Message
        isDictating := false
    }

    Send("{CapsLock}")
}
