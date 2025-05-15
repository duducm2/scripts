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
CapsLock & b:: {                                                 ; Alt + Shift + B
    static isDictating     := false
    static submitFailCount := 0

    ; Activate ChatGPT window
    SetTitleMatchMode 2
    WinActivate "chatgpt - transcription"
    WinWaitActive "ahk_exe chrome.exe"
    cUIA := UIA_Browser()
    Sleep 300

    ; UI labels (Portuguese only)
    dictateNames    := ["Botão de ditado"]
    submitNames     := ["Enviar ditado"]
    transcribingDicating := ["Interromper ditado"]
    sendPromptNames := ["Enviar prompt"]
    stopNames       := ["Interromper transmissão", "Stop streaming"]  ; fallback EN optional

    promptHeader := "
(
hey ChatGPT, please help me out correcting the grammar of the following sentence,
making it more cohesive, and also using the already names from my colleagues,
professors, and friends that you already have, also to correct it, and also try
to make this more human-like in the end. Please, do not reply anything, just the
plain text corrected.

)"

;───────────────────────────── START DICTATION ───────────────────────────────
    if !isDictating {
        try {
            Send "+{Escape}"            ; focus textbox
            Sleep 80
            Send "^a{Delete}"
            Sleep 80
            if btn := FindButton(cUIA, dictateNames) {
                btn.Click()
                isDictating     := true
                submitFailCount := 0
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
            submitFailCount := 0

            ; wait until “Enviar ditado” disappears  (=> transcription done)
            ok := cUIA.WaitElementNotExist({ Name: transcribingDicating[1], Type: "Button" }, 60000)
            if !ok {
                MsgBox "Tempo-esgotado: transcrição não terminou em 60 s."
                return
            }
            Sleep 150                   ; small buffer

            ; copy raw transcript
            Send "+{Escape}"
            Sleep 150
            Send "^a^c"
            if !ClipWait(0.5) {
                MsgBox "Falha ao copiar transcrição."
                return
            }
            ; trans := Trim(A_Clipboard, "`r`n ")

            ; build full prompt
            ; fullPrompt := promptHeader . trans

            ; fast replace textbox
            ; cbOld := A_Clipboard
            Sleep 300
            Send "{Right}"
            Send "+{Enter}"
            Send "+{Enter}"
            A_Clipboard := promptHeader
            ClipWait 0.3
            Sleep 150
            Send "^v"
            Sleep 150
            Send "{Enter}"              ; send to ChatGPT
            ; Sleep 250
            ; A_Clipboard := cbOld        ; restore clipboard

            ; OPTIONAL: copy ChatGPT answer when done
            if FindButton(cUIA, stopNames, "Button", 5000) {
                cUIA.WaitElementNotExist({ Name: stopNames[1], Type: "Button" }, 60000)
                cUIA.WaitElementNotExist({ Name: stopNames[2], Type: "Button" }, 1)
                Sleep 400
                Send "^+c"
            }
        } else {
            MsgBox "‘Enviar ditado’ não encontrado."
            submitFailCount++
        }
    } catch as e {
        MsgBox "Erro ao finalizar ditado:`n" e.Message
        submitFailCount++
    }

    if submitFailCount {
        MsgBox "Falha ao enviar ditado. Estado reiniciado-pressione o atalho novamente."
        isDictating     := false
        submitFailCount := 0
    }

    Send("{CapsLock}")
}
