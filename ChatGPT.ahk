#Requires AutoHotkey v2.0+
#SingleInstance Force

; -----------------------------------------------------------------------------
; This script consolidates all ChatGPT related hotkeys and functions.
; -----------------------------------------------------------------------------

; --- Includes ----------------------------------------------------------------
#include UIA-v2\Lib\UIA.ahk
#include UIA-v2\Lib\UIA_Browser.ahk
#include %A_ScriptDir%\env.ahk

; --- Config ---------------------------------------------------------------
; Path to the file containing the initial prompt ChatGPT should receive.
PROMPT_FILE := A_ScriptDir "\ChatGPT_Prompt.txt"

; --- Persistent Dictation Indicator ----------------------------------------
; Holds the GUI object while dictation is active. Empty string when hidden.
global dictationIndicatorGui := ""

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
; Original File: Open Chat GPT.ahk
; =============================================================================
#!+i::
{
    SetTitleMatchMode(2)
    if WinExist("chatgpt") {
        WinActivate("chatgpt")
        if WinWaitActive("ahk_exe chrome.exe", , 2)
            CenterMouse()
        Send("{Esc}")
    } else {
        Run "chrome.exe --new-window https://chatgpt.com/"
        if WinWaitActive("ahk_exe chrome.exe", , 5) {
            CenterMouse()
            ; --- Read initial prompt from external file & paste it ---
            Sleep 7000  ; give the page time to load fully
            promptText := ""
            try promptText := FileRead(PROMPT_FILE, "UTF-8")
            if (StrLen(promptText) = 0)
                promptText := "hey, what's up?"

            ; Copy–paste to handle Unicode & speed
            oldClip := A_Clipboard
            A_Clipboard := ""           ; clear first so ClipWait can detect change
            A_Clipboard := promptText
            ClipWait 1                   ; wait up to 1 s for clipboard to update
            Send("^v")                  ; paste
            Sleep 100
            Send("{Enter}")             ; submit
            Sleep 100
            A_Clipboard := oldClip       ; restore previous clipboard
            ; --- Open the options menu (three dots) for the first conversation ---
            Sleep 7500  ; wait for sidebar to update with new chat
            if ClickFirstConversationOptions() {
                ; Navigate to "Rename" and apply new name
                Sleep 300
                Send("{Down}")
                Sleep 120
                Send("{Down}")
                Sleep 80
                Send("{Enter}")
                Sleep 300
                Send("chatgpt")
                Sleep 120
                Send("{Enter}")
            }
        }
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
    if WinWaitActive("ahk_exe chrome.exe", , 2)
        CenterMouse()
    Sleep 250
    Send "{Esc}"
    Sleep 250
    Send "+{Esc}"
    searchString :=
        "Correct the following sentence for grammar, cohesion, and coherence. Respond with only the corrected sentence, no explanations."
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
    if WinWaitActive("ahk_exe chrome.exe", , 2)
        CenterMouse()
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
;  Copy Last Prompt  –  ChatGPT (Chrome)
;  Hotkey:  Win + Alt + Shift + P   (#!+p)
; =============================================================================

#!+p:: CopyLastPrompt()

CopyLastPrompt() {
    SetTitleMatchMode 2
    WinActivate "chatgpt"
    if !WinWaitActive("ahk_exe chrome.exe", , 1)
        return
    CenterMouse()

    cUIA := UIA_Browser()
    Sleep 300

    ; — labels ChatGPT shows in EN and PT —
    copyNames := [
        "Copy to clipboard", "Copiar para a área de transferência", "Copy", "Copiar"
    ]

    ; — collect every matching button —
    copyBtns := []
    for name in copyNames
        for btn in cUIA.FindAll({ Name: name, Type: "Button", matchmode: "Substring" })
            copyBtns.Push(btn)

    if !copyBtns.Length {
        MsgBox "⚠️  No copy button found (EN / PT)."
        return
    }

    lastBtn := copyBtns[copyBtns.Length]   ; the bottom-most one

    isCopied := false
    ; — click it & wait for clipboard —
    try {
        lastBtn.ScrollIntoView()
        Sleep 100
        A_Clipboard := ""                  ; clear first
        lastBtn.Click()
        if !ClipWait(1) {                  ; returns 0 on timeout
            MsgBox "Copy failed – clipboard stayed empty."
        } else {
            isCopied := true
        }
    } catch as e {
        MsgBox "Error clicking copy button:`n" e.Message
    }

    ; optional: jump back to previous window
    Send "!{Tab}"

    if (isCopied) {
        Sleep(300) ; give window time to switch
        ShowNotification("Last prompt copied!")
    }
}

; =============================================================================
; Copy Last Message and Read Aloud
; Hotkey: Win+Alt+Shift+Y
; =============================================================================
#!+y::
{
    SetTitleMatchMode 2
    WinActivate "chatgpt"
    if !WinWaitActive("ahk_exe chrome.exe", , 2)
        return
    CenterMouse()

    cUIA := UIA_Browser()
    Sleep 300

    ; — labels ChatGPT shows in EN and PT —
    copyNames := [
        "Copy to clipboard", "Copiar para a área de transferência", "Copy", "Copiar"
    ]

    ; — collect every matching button —
    copyBtns := []
    for name in copyNames
        for btn in cUIA.FindAll({ Name: name, Type: "Button", matchmode: "Substring" })
            copyBtns.Push(btn)

    if !copyBtns.Length {
        MsgBox "⚠️  No copy button found (EN / PT)."
        return
    }

    lastBtn := copyBtns[copyBtns.Length]   ; the bottom-most one

    ; — click copy & wait for clipboard —
    try {
        lastBtn.ScrollIntoView()
        Sleep 100
        A_Clipboard := ""                  ; clear first
        lastBtn.Click()
        if !ClipWait(1) {                  ; returns 0 on timeout
            MsgBox "Copy failed – clipboard stayed empty."
            return
        }
    } catch as e {
        MsgBox "Error clicking copy button:`n" e.Message
        return
    }

    ; Now trigger read aloud
    Sleep 300
    readNames := ["Read aloud", "Ler em voz alta"]
    stopNames := ["Stop", "Parar"]
    buttonClicked := false
    stopBtns := []
    for name in stopNames
        for btn in cUIA.FindAll({ Name: name, Type: "Button" })
            stopBtns.Push(btn)
    if stopBtns.Length {
        stopBtns[stopBtns.Length].Click()
        Sleep 200  ; Small pause before starting new read
    }

    readBtns := []
    for name in readNames
        for btn in cUIA.FindAll({ Name: name, Type: "Button" })
            readBtns.Push(btn)
    if readBtns.Length {
        readBtns[readBtns.Length].Click()
        buttonClicked := true
    } else {
        MsgBox "No 'Read aloud/Ler em voz alta' button found!"
        return
    }

    if (buttonClicked) {
        ShowNotification("Message copied and reading!")
        Send "!{Tab}"  ; Return to previous window
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
    CenterMouse()
    cUIA := UIA_Browser()
    Sleep 300
    readNames := ["Read aloud", "Ler em voz alta"]
    stopNames := ["Stop", "Parar"]
    buttonClicked := false
    stopBtns := []
    for name in stopNames
        for btn in cUIA.FindAll({ Name: name, Type: "Button" })
            stopBtns.Push(btn)
    if stopBtns.Length {
        stopBtns[stopBtns.Length].Click()
        buttonClicked := true
    } else {
        readBtns := []
        for name in readNames
            for btn in cUIA.FindAll({ Name: name, Type: "Button" })
                readBtns.Push(btn)
        if readBtns.Length {
            readBtns[readBtns.Length].Click()
            buttonClicked := true
        } else {
            MsgBox "Nenhum botão 'Read aloud/Ler em voz alta' ou 'Stop/Parar' encontrado!"
        }
    }
    if (buttonClicked) {
        Send "!{Tab}"
    }
}

; =============================================================================
; Toggle Voice Mode
; Hotkey: Win+Alt+Shift+L
; Original File: ChatGPT - Talk trough voice.ahk
; =============================================================================
#!+l::
{
    ToggleVoiceMode()
    Send "!{Tab}"
}

ToggleVoiceMode(triedFallback := false, forceAction := "") {
    static isVoiceModeActive := false
    SetTitleMatchMode 2
    WinActivate "chatgpt"
    WinWaitActive "ahk_exe chrome.exe"
    CenterMouse()
    cUIA := UIA_Browser()
    Sleep 300

    ; Button names in both English and Portuguese
    startNames := ["Start voice mode", "Iniciar modo voz"]
    endNames := ["End voice mode", "Encerrar modo voz"]

    action := forceAction ? forceAction : (!isVoiceModeActive ? "start" : "stop")

    if (action = "start") {
        try {
            if btn := FindButtonByNames(cUIA, startNames) {
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
            if btn := FindButtonByNames(cUIA, endNames) {
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

FindButtonByNames(cUIA, namesArray) {
    for name in namesArray
        for btn in cUIA.FindAll({ Name: name, Type: "Button" })
            return btn
    return ""
}
; =============================================================================
; Toggle Dictation (No Auto-Send)
; Hotkey: Win+Alt+Shift+0
; Original File: ChatGPT - Toggle microphone.ahk
; =============================================================================
#!+0::
{
    ToggleDictation()
    Send "!{Tab}"
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
    CenterMouse()
    cUIA := UIA_Browser()
    Sleep 300
    action := forceAction ? forceAction : (!isDictating ? "start" : "stop")
    if (action = "start") {
        try {
            if btn := FindButton(cUIA, dictateNames_to_find) {
                btn.Click()
                isDictating := true
                ShowDictationIndicator()
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
                HideDictationIndicator()
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
    Send "!{Tab}"
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
    CenterMouse()
    cUIA := UIA_Browser()
    Sleep 300
    action := forceAction ? forceAction : (!isDictating ? "start" : "stop")
    if (action = "start") {
        try {
            dictateBtn := cUIA.FindElement({ Name: currentDictateName, Type: "Button" })
            if dictateBtn {
                dictateBtn.Click()
                isDictating := true
                ShowDictationIndicator()
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
                HideDictationIndicator()
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
            HideDictationIndicator()
            submitFailCount := 0
        }
    }
}

; =============================================================================
; Copy Last Code
; Hotkey: Win+Alt+Shift+U
; =============================================================================
#!+u::
{
    SetTitleMatchMode(2)
    A_Clipboard := "" ; Empty clipboard to check for new content later

    if WinExist("chatgpt") {
        WinActivate("chatgpt")
        if WinWaitActive("ahk_exe chrome.exe", , 2)
            CenterMouse()
    } else {
        Run "chrome.exe --new-window https://chatgpt.com/"
        if !WinWaitActive("ahk_exe chrome.exe", , 3)
            return
        CenterMouse()
    }
    Sleep 300
    Send("^+;")

    Send("!{Tab}") ; Switch back to the previous window
    Sleep(300) ; Wait for the window switch

    if ClipWait(1) {
        ShowNotification("Last code block copied!")
    }
}

; =============================================================================
; Helper function to center mouse on the active window
; =============================================================================
CenterMouse() {
    Sleep(200)
    Send("#!+q")
}

; =============================================================================
; Helper function to show a notification on the active window
; =============================================================================
ShowNotification(message, durationMs := 2000, bgColor := "FFFF00", fontColor := "000000", fontSize := 24) {
    notificationGui := Gui()
    notificationGui.Opt("+AlwaysOnTop -Caption +ToolWindow")
    notificationGui.BackColor := bgColor
    notificationGui.SetFont("s" . fontSize . " c" . fontColor . " Bold", "Segoe UI")
    notificationGui.Add("Text", "w400 Center", message)

    ; To center on the whole screen if no window is active
    activeWin := WinGetID("A")
    if (activeWin) {
        WinGetPos(&winX, &winY, &winW, &winH, activeWin)
    } else {
        workArea := SysGet.MonitorWorkArea(SysGet.MonitorPrimary)
        winX := workArea.Left, winY := workArea.Top, winW := workArea.Right - workArea.Left, winH := workArea.Bottom -
            workArea.Top
    }

    notificationGui.Show("AutoSize Hide")
    guiW := 0, guiH := 0
    notificationGui.GetPos(, , &guiW, &guiH)

    guiX := winX + (winW - guiW) / 2
    guiY := winY + (winH - guiH) / 2
    notificationGui.Show("x" . Round(guiX) . " y" . Round(guiY) . " NA")
    WinSetTransparent(220, notificationGui)

    Sleep(durationMs)
    if IsObject(notificationGui) && notificationGui.Hwnd {
        notificationGui.Destroy()
    }
}

; =============================================================================
; Helper functions to show/hide the persistent dictation indicator
; =============================================================================
ShowDictationIndicator(message := "Dictation ON", bgColor := "FF0000", fontColor := "FFFFFF", fontSize := 28) {
    global dictationIndicatorGui

    ; If already visible, nothing to do
    if (IsObject(dictationIndicatorGui) && dictationIndicatorGui.Hwnd)
        return

    dictationIndicatorGui := Gui()
    dictationIndicatorGui.Opt("+AlwaysOnTop -Caption +ToolWindow")
    dictationIndicatorGui.BackColor := bgColor
    dictationIndicatorGui.SetFont("s" . fontSize . " c" . fontColor . " Bold", "Segoe UI")
    dictationIndicatorGui.Add("Text", "w350 Center", message)

    ; Center over active window or primary monitor
    activeWin := WinGetID("A")
    if (activeWin) {
        WinGetPos(&winX, &winY, &winW, &winH, activeWin)
    } else {
        workArea := SysGet.MonitorWorkArea(SysGet.MonitorPrimary)
        winX := workArea.Left, winY := workArea.Top, winW := workArea.Right - workArea.Left, winH := workArea.Bottom -
            workArea.Top
    }

    dictationIndicatorGui.Show("AutoSize Hide")
    dictationIndicatorGui.GetPos(, , &guiW, &guiH)
    guiX := winX + (winW - guiW) / 2
    guiY := winY + (winH - guiH) / 2
    dictationIndicatorGui.Show("x" . Round(guiX) . " y" . Round(guiY) . " NA")
    WinSetTransparent(220, dictationIndicatorGui)
}

HideDictationIndicator() {
    global dictationIndicatorGui
    if (IsObject(dictationIndicatorGui) && dictationIndicatorGui.Hwnd) {
        dictationIndicatorGui.Destroy()
        dictationIndicatorGui := ""
    }
}

; =============================================================================
; Helper function to click the first conversation's options (three-dots) button
; =============================================================================
ClickFirstConversationOptions(timeoutMs := 5000) {
    cUIA := UIA_Browser()
    
    ; Possible labels (EN / PT-BR) and control types the button can have
    optNames := ["Open conversation options", "Abrir opções de conversa"]
    optTypes := ["Button", "MenuItem"]

    ; Try to find the element within the given timeout
    startTick := A_TickCount
    btn := ""
    for name in optNames {
        for role in optTypes {
            remaining := timeoutMs - (A_TickCount - startTick)
            if (remaining <= 0)
                break 2 ; out of both loops – overall timeout expired
            thisBtn := cUIA.WaitElement({ Name: name, Type: role }, remaining)
            if thisBtn {
                btn := thisBtn
                break 2
            }
        }
    }
    if !btn {
        MsgBox "⚠️  Conversation options button (EN/PT-BR) not found within " timeoutMs " ms."
        return false
    }
    try {
        btn.ScrollIntoView()
        Sleep 100
        btn.Click()
        return true
    } catch Error as e {
        MsgBox "Error clicking options button:`n" e.Message
        return false
    }
}

; =============================================================================
; Send Ctrl+A, Sleep, Ctrl+X to ChatGPT
; Hotkey: Win+Alt+Shift+J
; =============================================================================
#!+j::
{
    SetTitleMatchMode(2)
    WinActivate("chatgpt")
    if WinWaitActive("ahk_exe chrome.exe", , 2) {
        Send("{Esc}")
        Send("F")
        Send("{Backspace}")
        Sleep(300)
        Send("^a")
        Sleep(350)
        Send("^x")
        Send("!{Tab}")
    }
}
