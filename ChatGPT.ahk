#Requires AutoHotkey v2.0+
#SingleInstance Force

; -----------------------------------------------------------------------------
; This script consolidates all ChatGPT related hotkeys and functions.
; -----------------------------------------------------------------------------

; --- Includes ----------------------------------------------------------------
#include UIA-v2\Lib\UIA.ahk
#include UIA-v2\Lib\UIA_Browser.ahk
#include %A_ScriptDir%\env.ahk

; --- Helper Functions --------------------------------------------------------

; Find the first UIA element whose Name matches any string in an array
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

; --- Hotkeys & Functions -----------------------------------------------------

; =============================================================================
; Open ChatGPT
; Hotkey: Win+Alt+Shift+I
; Original File: Open ChatGPT.ahk
; =============================================================================
#!+i::
{
    SetTitleMatchMode(2)
    if WinExist("chatgpt") {
        WinActivate("chatgpt")
        Send("{Esc}")
    } else {
        Run "chrome.exe --new-window https://chatgpt.com/"
    }
}

; =============================================================================
; Check Grammar
; Hotkey: Win+Alt+Shift+O
; Original File: ChatGPT - Check for grammar.ahk
; =============================================================================
#!+o::
{
    A_Clipboard := ""  ; Start off empty to allow ClipWait to detect when the text has arrived.
    Send "^c"
    ClipWait
    SetTitleMatchMode 2
    WinActivate "chatgpt"
    Sleep 250
    Send "{Esc}"
    Sleep 250
    Send "+{Esc}"
    searchString :=
        "Below, you will find content. This content can be either a word or a sentence, in Portuguese or English. I would like you to correct the sentence based on grammar, taking cohesion and coherence into consideration. Please, don't prompt any additional comment, neither put your answer into quotation marks. Remember, we are not playing thumbs up or thumbs down now."
    A_Clipboard := searchString . "`n`nContent: " . A_Clipboard
    Sleep 100
    Send("^a")
    Sleep 500
    Send("^v")
    Sleep 500
    Send("{Enter}")
    Sleep 500
}

; =============================================================================
; Get Pronunciation
; Hotkey: Win+Alt+Shift+8
; Original File: ChatGPT - Pronunciation.ahk
; =============================================================================
#!+8::
{
    A_Clipboard := ""
    Send "^c"
    ClipWait
    SetTitleMatchMode 2
    WinActivate "chatgpt"
    Sleep 250
    Send "{Esc}"
    Sleep 250
    Send "+{Esc}"
    searchString :=
        "Below, you will find a word or phrase. I'd like you to answer in five sections: the 1st section you will repeat the word twice. For each time you repeat, use a point to finish the phrase. The 2nd section should have the definition of the word (You should also say each part of speech does the different definitions belong to). The 3d section should have the pronunciation of this word using the Internation Phonetic Alphabet characters (for American English).The 4th section should have the same word applied in a real sentence (put that in quotations, so I can identify that). In the 5th, Write down the translation of the word into Portuguese. Please, do not title any section. Thanks!"
    A_Clipboard := searchString . "`n`nContent: " . A_Clipboard
    Sleep 100
    Send("^a")
    Sleep 500
    Send("^v")
    Sleep 500
    Send("{Enter}")
    Sleep 500
    Send "!{Tab}"
}

; =============================================================================
; Copy Last Message
; Hotkey: Win+Alt+Shift+P
; Original File: ChatGPT - Copy last message.ahk
; =============================================================================
#!+p:: CopyLastPrompt()

CopyLastPrompt() {
    global IS_WORK_ENVIRONMENT
    static pt_copyName := "Copiar"
    static en_copyName := "Copy"
    copyName := IS_WORK_ENVIRONMENT ? pt_copyName : en_copyName
    SetTitleMatchMode 2
    WinActivate "chatgpt"
    WinWaitActive "ahk_exe chrome.exe"
    cUIA := UIA_Browser()
    Sleep 300
    try {
        buttons := cUIA.FindAll({ Name: copyName, Type: "Button" })
        if buttons.Length {
            lastBtn := buttons[buttons.Length]
            lastBtn.Click()
        } else {
            MsgBox(IS_WORK_ENVIRONMENT ? "Nenhum botão 'Copiar' encontrado." : "No 'Copy' button found.")
        }
    } catch as e {
        MsgBox(IS_WORK_ENVIRONMENT ? "Erro ao copiar:`n" e.Message : "Error copying:`n" e.Message)
    }
}

; =============================================================================
; Toggle "Read Aloud"
; Hotkey: Win+Alt+Shift+9
; Original File: ChatGPT - Click last microphone.ahk
; =============================================================================
#!+9::
{
    SetTitleMatchMode 2
    winTitle := "chatgpt"
    WinActivate winTitle
    WinWaitActive "ahk_exe chrome.exe"
    cUIA := UIA_Browser()
    Sleep 300
    readNames := ["Read aloud", "Ler em voz alta"]
    stopNames := ["Stop", "Parar"]
    stopBtns := []
    for name in stopNames
        for btn in cUIA.FindAll({ Name: name, Type: "Button" })
            stopBtns.Push(btn)
    if stopBtns.Length {
        stopBtns[stopBtns.Length].Click()
        return
    }
    readBtns := []
    for name in readNames
        for btn in cUIA.FindAll({ Name: name, Type: "Button" })
            readBtns.Push(btn)
    if readBtns.Length
        readBtns[readBtns.Length].Click()
    else
        MsgBox "Nenhum botão 'Read aloud/Ler em voz alta' ou 'Stop/Parar' encontrado!"
}

; =============================================================================
; Toggle Voice Mode
; Hotkey: Win+Alt+Shift+L
; Original File: ChatGPT - Talk trough voice.ahk
; =============================================================================
#!+l::
{
    ToggleVoiceMode()
}

ToggleVoiceMode(triedFallback := false, forceAction := "") {
    static isVoiceModeActive := false
    startVoiceName := "Start voice mode"
    endVoiceName := "End voice mode"
    startNames_to_find := [startVoiceName]
    endNames_to_find := [endVoiceName]
    SetTitleMatchMode 2
    WinActivate "chatgpt"
    WinWaitActive "ahk_exe chrome.exe"
    cUIA := UIA_Browser()
    Sleep 300
    action := forceAction ? forceAction : (!isVoiceModeActive ? "start" : "stop")
    if (action = "start") {
        try {
            if btn := FindButton(cUIA, startNames_to_find) {
                btn.Click()
                isVoiceModeActive := true
                return
            } else if !triedFallback {
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

; =============================================================================
; Toggle Dictation (No Auto-Send)
; Hotkey: Win+Alt+Shift+0
; Original File: ChatGPT - Toggle microphone.ahk
; =============================================================================
#!+0::
{
    ToggleDictation()
}

ToggleDictation(triedFallback := false, forceAction := "") {
    static isDictating := false
    pt_dictateName := "Botão de ditado"
    en_dictateName := "Dictate button"
    pt_submitName := "Enviar ditado"
    en_submitName := "Submit dictation"
    pt_transcribingName := "Interromper ditado"
    en_transcribingName := "Stop dictation"
    currentDictateName := IS_WORK_ENVIRONMENT ? pt_dictateName : en_dictateName
    currentSubmitName := IS_WORK_ENVIRONMENT ? pt_submitName : en_submitName
    currentTranscribingName := IS_WORK_ENVIRONMENT ? pt_transcribingName : en_transcribingName
    dictateNames_to_find := [currentDictateName]
    submitOrStopNames_to_find := [currentSubmitName, currentTranscribingName]
    SetTitleMatchMode 2
    WinActivate "chatgpt"
    WinWaitActive "ahk_exe chrome.exe"
    cUIA := UIA_Browser()
    Sleep 300
    action := forceAction ? forceAction : (!isDictating ? "start" : "stop")
    if (action = "start") {
        try {
            if btn := FindButton(cUIA, dictateNames_to_find) {
                btn.Click()
                isDictating := true
                return
            } else if !triedFallback {
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
                btn.Click()
                isDictating := false
                return
            } else if !triedFallback {
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

; =============================================================================
; Toggle Dictation (with Auto-Send)
; Hotkey: Win+Alt+Shift+7
; Original File: ChatGPT - Speak.ahk
; =============================================================================
#!+7::
{
    ToggleDictationSpeak()
}

ToggleDictationSpeak(triedFallback := false, forceAction := "") {
    static isDictating := false
    static submitFailCount := 0
    pt_dictateName := "Botão de ditado"
    en_dictateName := "Dictate button"
    pt_submitName := "Enviar ditado"
    en_submitName := "Submit dictation"
    pt_transcribingName := "Interromper ditado"
    en_transcribingName := "Stop dictation"
    pt_sendPromptName := "Enviar prompt"
    en_sendPromptName := "Send prompt"
    pt_stopStreamingName := "Interromper transmissão"
    en_stopStreamingName := "Stop streaming"
    currentDictateName := IS_WORK_ENVIRONMENT ? pt_dictateName : en_dictateName
    currentSubmitName := IS_WORK_ENVIRONMENT ? pt_submitName : en_submitName
    currentTranscribingName := IS_WORK_ENVIRONMENT ? pt_transcribingName : en_transcribingName
    currentSendPromptName := IS_WORK_ENVIRONMENT ? pt_sendPromptName : en_sendPromptName
    currentStopStreamingName := IS_WORK_ENVIRONMENT ? pt_stopStreamingName : en_stopStreamingName
    dictateNames_to_find := [currentDictateName]
    submitOrStopNames_to_find := [currentSubmitName, currentTranscribingName]
    SetTitleMatchMode 2
    WinActivate "chatgpt"
    WinWaitActive "ahk_exe chrome.exe"
    cUIA := UIA_Browser()
    Sleep 300
    action := forceAction ? forceAction : (!isDictating ? "start" : "stop")
    if (action = "start") {
        try {
            dictateBtn := cUIA.FindElement({ Name: currentDictateName, Type: "Button" })
            if dictateBtn {
                dictateBtn.Click()
                isDictating := true
                submitFailCount := 0
                return
            } else if !triedFallback {
                ToggleDictationSpeak(true, "stop")
                return
            } else {
                MsgBox currentDictateName . " not found, and could not stop dictation either."
            }
        } catch Error as e {
            if !triedFallback {
                ToggleDictationSpeak(true, "stop")
                return
            } else {
                MsgBox "Error during pre-dictation or starting dictation: " e.Message
                isDictating := false
            }
        }
    } else if (action = "stop") {
        try {
            submitBtn := cUIA.FindElement({ Name: currentSubmitName, Type: "Button" })
            if !submitBtn {
                submitBtn := cUIA.FindElement({ Name: currentTranscribingName, Type: "Button" })
            }
            if submitBtn {
                submitBtn.Click()
                isDictating := false
                submitFailCount := 0
                try {
                    Sleep 200
                    finalSendBtn := cUIA.WaitElement({ Name: currentSendPromptName, AutomationId: "composer-submit-button" },
                    13000)
                    if finalSendBtn {
                        SendInput "{Enter}"
                    } else {
                        MsgBox "Timeout: '" . currentSendPromptName .
                            "' button did not reappear after submitting dictation."
                    }
                } catch Error as e_wait {
                    MsgBox "Error waiting for/clicking final " . currentSendPromptName . " button: " e_wait.Message
                }
                return
            } else if !triedFallback {
                ToggleDictationSpeak(true, "start")
                return
            } else {
                MsgBox currentSubmitName . " or " . currentTranscribingName .
                    " button not found, and could not start dictation either."
                submitFailCount++
            }
        } catch Error as e {
            if !triedFallback {
                ToggleDictationSpeak(true, "start")
                return
            } else {
                MsgBox "Error finding or clicking Submit/Stop dictation button: " e.Message
                submitFailCount++
            }
        }
        if submitFailCount >= 1 {
            MsgBox "Failed to submit dictation 1 time. Assuming dictation stopped. Press hotkey again to start."
            isDictating := false
            submitFailCount := 0
        }
    }
}
