#Requires AutoHotkey v2.0+
#SingleInstance Force

; -----------------------------------------------------------------------------
; This script consolidates all ChatGPT related hotkeys and functions.
; -----------------------------------------------------------------------------

; Global error handler to suppress system error sounds
OnError(ErrorHandler)

ErrorHandler(exception, mode) {
    ; Silently log errors without showing system dialogs or sounds
    ; Return 1 to suppress the error, -1 to show default error dialog
    return 1
}

; Suppress Windows error sounds using multiple methods
SuppressErrorSounds() {
    try {
        ; Method 1: Disable error beeps via registry
        RegWrite(0, "REG_DWORD", "HKCU\Control Panel\Sound", "Beep")

        ; Method 2: Set SPI_SETBEEP to FALSE
        DllCall("SystemParametersInfo", "UInt", 0x0002, "UInt", 0, "Ptr", 0, "UInt", 0)

        ; Method 3: Disable specific sound events
        DllCall("User32.dll\SystemParametersInfo", "UInt", 0x0003, "UInt", 0, "Ptr", 0, "UInt", 0x0001)
    } catch {
        ; Silently ignore if we can't suppress sounds
    }
}

; Restore error sounds
RestoreErrorSounds() {
    try {
        ; Restore error beeps
        RegWrite(1, "REG_DWORD", "HKCU\Control Panel\Sound", "Beep")

        ; Re-enable beep
        DllCall("SystemParametersInfo", "UInt", 0x0002, "UInt", 1, "Ptr", 0, "UInt", 0)
    } catch {
        ; Silently ignore restore errors
    }
}

; Wrapper functions to suppress Windows error sounds
SafeWinActivate(winTitle) {
    try {
        WinActivate(winTitle)
        return true
    } catch {
        return false
    }
}

SafeWinWaitActive(winTitle, winText := "", timeout := 1) {
    try {
        return WinWaitActive(winTitle, winText, timeout)
    } catch {
        return false
    }
}

SafeSend(keys) {
    try {
        Send(keys)
        return true
    } catch {
        return false
    }
}

SafeClick(element) {
    try {
        if IsObject(element)
            element.Click()
        return true
    } catch {
        return false
    }
}

; Temporarily disable Windows system sounds
DisableSystemSounds() {
    try {
        ; Disable system sound scheme temporarily
        RegWrite("(None)", "REG_SZ", "HKCU\AppEvents\Schemes\Apps\.Default\.Default\.Current")
        RegWrite("(None)", "REG_SZ", "HKCU\AppEvents\Schemes\Apps\.Default\SystemExclamation\.Current")
        RegWrite("(None)", "REG_SZ", "HKCU\AppEvents\Schemes\Apps\.Default\SystemHand\.Current")
        RegWrite("(None)", "REG_SZ", "HKCU\AppEvents\Schemes\Apps\.Default\SystemQuestion\.Current")
    } catch {
        ; Silently ignore registry errors
    }
}

; Re-enable Windows system sounds (restore defaults)
EnableSystemSounds() {
    try {
        ; Restore default Windows sound scheme
        RegWrite("C:\Windows\media\Windows Error.wav", "REG_SZ",
            "HKCU\AppEvents\Schemes\Apps\.Default\.Default\.Current")
        RegWrite("C:\Windows\media\Windows Exclamation.wav", "REG_SZ",
            "HKCU\AppEvents\Schemes\Apps\.Default\SystemExclamation\.Current")
        RegWrite("C:\Windows\media\Windows Critical Stop.wav", "REG_SZ",
            "HKCU\AppEvents\Schemes\Apps\.Default\SystemHand\.Current")
        RegWrite("C:\Windows\media\Windows Question.wav", "REG_SZ",
            "HKCU\AppEvents\Schemes\Apps\.Default\SystemQuestion\.Current")
    } catch {
        ; Silently ignore registry errors
    }
}

; --- Includes ----------------------------------------------------------------
#include UIA-v2\Lib\UIA.ahk
#include UIA-v2\Lib\UIA_Browser.ahk
#include %A_ScriptDir%\env.ahk
#include %A_ScriptDir%\ChatGPT_Loading.ahk

; --- Config ---------------------------------------------------------------
; Path to the file containing the initial prompt ChatGPT should receive.
PROMPT_FILE := A_ScriptDir "\ChatGPT_Prompt.txt"

; --- Persistent Dictation Indicator ----------------------------------------
; Holds the GUI object while dictation is active. Empty string when hidden.
global dictationIndicatorGui := ""
; Flag to request a one-time transcription-finished chime for the Win+Alt+Shift+0 flow
global g_transcribeChimePending := false

; --- Persistent Loading Indicator ------------------------------------------
; Holds the GUI object shown while the script is preparing ChatGPT.
global loadingGui := ""

; --- Helper Functions --------------------------------------------------------

; Find the first UIA element whose Name matches any string in an array
FindButton(cUIA, names, role := "Button", timeoutMs := 0) {
    for name in names {
        try {
            el := (timeoutMs = 0)
                ? cUIA.FindElement({ Name: name, Type: role })
                : cUIA.WaitElement({ Name: name, Type: role }, timeoutMs)
            if el
                return el
        } catch {
            ; Silently continue to next name if this one fails
        }
    }
    return false
}

; Collect visible elements by Name only (exact match, case-insensitive)
CollectByNamesExact(root, namesArr) {
    results := []
    for name in namesArr {
        try {
            for el in root.FindAll({ Name: name, mm: 3, cs: false })
                if !(el.IsOffscreen)
                    results.Push(el)
        } catch {
            ; Silently ignore errors
        }
    }
    return results
}

; Collect visible elements by Name only (substring match, case-insensitive)
CollectByNamesContains(root, namesArr) {
    results := []
    for name in namesArr {
        try {
            for el in root.FindAll({ Name: name, mm: 2, cs: false })
                if !(el.IsOffscreen)
                    results.Push(el)
        } catch {
            ; Silently ignore errors
        }
    }
    return results
}

; Find ChatGPT browser window (case-insensitive contains match for "chatgpt")
GetChatGPTWindowHwnd() {
    try {
        for hwnd in WinGetList("ahk_exe chrome.exe") {
            try {
                if InStr(WinGetTitle("ahk_id " hwnd), "chatgpt", false)
                    return hwnd
            } catch {
                ; Silently skip invalid windows
            }
        }
    } catch {
        ; Silently handle WinGetList errors
    }
    return 0
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
    if hwnd := GetChatGPTWindowHwnd() {
        WinActivate "ahk_id " hwnd
        if WinWaitActive("ahk_id " hwnd, , 2) {
            CenterMouse()
        }
        SafeSend("{Esc}")
    } else {
        Run "chrome.exe --new-window https://chatgpt.com/"
        if WinWaitActive("ahk_exe chrome.exe", , 5) {
            RecenterLoadingOverWindow(WinExist("A"))
            CenterMouse()
            ; --- Read initial prompt from external file & paste it ---
            Sleep (IS_WORK_ENVIRONMENT ? 3500 : 7000)  ; give the page time to load fully
            promptText := ""
            try promptText := FileRead(PROMPT_FILE, "UTF-8")
            if (StrLen(promptText) = 0)
                promptText := "hey, what's up?"

            ; Copy–paste to handle Unicode & speed
            oldClip := A_Clipboard
            A_Clipboard := ""
            A_Clipboard := promptText
            ClipWait 1
            SafeSend("^v")
            Sleep (IS_WORK_ENVIRONMENT ? 50 : 100)
            SafeSend("{Enter}")
            Sleep (IS_WORK_ENVIRONMENT ? 50 : 100)
            A_Clipboard := oldClip
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
    if hwnd := GetChatGPTWindowHwnd()
        WinActivate "ahk_id " hwnd
    if WinWaitActive("ahk_exe chrome.exe", , 2)
        CenterMouse()
    Sleep (IS_WORK_ENVIRONMENT ? 125 : 250)
    SafeSend("{Esc}")
    Sleep (IS_WORK_ENVIRONMENT ? 125 : 250)
    SafeSend("+{Esc}")
    searchString :=
        "Correct the following sentence for grammar, cohesion, and coherence. Respond with only the corrected sentence, no explanations."
    A_Clipboard := searchString . "`n`nContent: " . A_Clipboard
    Sleep (IS_WORK_ENVIRONMENT ? 50 : 100)
    Send("^a")
    Sleep (IS_WORK_ENVIRONMENT ? 250 : 500)
    Send("^v")
    Sleep (IS_WORK_ENVIRONMENT ? 250 : 500)
    Send("{Enter}")
    Sleep (IS_WORK_ENVIRONMENT ? 250 : 500)
    ; After sending, show loading for Stop streaming
    Send "!{Tab}" ; Return to previous window
    buttonNames := ["Stop streaming", "Interromper transmissão"]
    WaitForButtonAndShowSmallLoading(buttonNames, "Waiting for response...")

    ; Play completion sound
    PlayGrammarCheckCompletedChime()
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
    if hwnd := GetChatGPTWindowHwnd()
        WinActivate "ahk_id " hwnd
    if WinWaitActive("ahk_exe chrome.exe", , 2)
        CenterMouse()
    Sleep (IS_WORK_ENVIRONMENT ? 125 : 250)
    Send "{Esc}"
    Sleep (IS_WORK_ENVIRONMENT ? 125 : 250)
    Send "+{Esc}"
    searchString :=
        "Below, you will find a word or phrase. I'd like you to answer in five sections: the 1st section you will repeat the word twice. For each time you repeat, use a point to finish the phrase. The 2nd section should have the definition of the word (You should also say each part of speech does the different definitions belong to). The 3d section should have the pronunciation of this word using the Internation Phonetic Alphabet characters (for American English).The 4th section should have the same word applied in a real sentence (put that in quotations, so I can identify that). In the 5th, Write down the translation of the word into Portuguese. Please, do not title any section. Thanks!"
    A_Clipboard := searchString . "`n`nContent: " . A_Clipboard
    Sleep (IS_WORK_ENVIRONMENT ? 50 : 100)
    Send("^a")
    Sleep (IS_WORK_ENVIRONMENT ? 250 : 500)
    Send("^v")
    Sleep (IS_WORK_ENVIRONMENT ? 250 : 500)
    Send("{Enter}")
    Sleep (IS_WORK_ENVIRONMENT ? 250 : 500)
    ; After sending, show loading for Stop streaming
    Send "!{Tab}" ; Return to previous window
    buttonNames := ["Stop streaming", "Interromper transmissão"]
    WaitForButtonAndShowSmallLoading(buttonNames, "Waiting for response...")
}

; =============================================================================
;  Copy Last Prompt  –  ChatGPT (Chrome)
;  Hotkey:  Win + Alt + Shift + P   (#!+p)
; =============================================================================

#!+p:: CopyLastPrompt()

CopyLastPrompt() {
    SetTitleMatchMode 2
    if hwnd := GetChatGPTWindowHwnd()
        WinActivate "ahk_id " hwnd
    if !WinWaitActive("ahk_exe chrome.exe", , 1)
        return
    CenterMouse()

    try {
        cUIA := UIA_Browser()
        Sleep (IS_WORK_ENVIRONMENT ? 150 : 300)
    } catch {
        ShowNotification("Unable to connect to browser", 1500, "DF2935", "FFFFFF")
        return
    }

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
        ShowNotification("⚠️ No copy button found", 1500, "DF2935", "FFFFFF")
        return
    }

    lastBtn := copyBtns[copyBtns.Length]   ; the bottom-most one

    isCopied := false
    ; — click it & wait for clipboard —
    try {
        lastBtn.ScrollIntoView()
        Sleep (IS_WORK_ENVIRONMENT ? 50 : 100)
        A_Clipboard := ""                  ; clear first
        SafeClick(lastBtn)
        if !ClipWait(1) {                  ; returns 0 on timeout
            ShowNotification("Copy failed – clipboard stayed empty", 1500, "DF2935", "FFFFFF")
        } else {
            isCopied := true
        }
    } catch as e {
        ShowNotification("Error clicking copy button", 1500, "DF2935", "FFFFFF")
    }

    ; optional: jump back to previous window
    Send "!{Tab}"

    if (isCopied) {
        Sleep(IS_WORK_ENVIRONMENT ? 150 : 300) ; give window time to switch
        ShowNotification("Last prompt copied!")
    }
}

; =============================================================================
; Open Read Aloud via More actions on latest assistant message
; Hotkey: Win+Alt+Shift+Y
; =============================================================================
#!+y::
{

    ; If Spotify exists, send a media pause key
    if ProcessExist("Spotify.exe")
        Send("{Media_Stop}")

    SetTitleMatchMode 2
    if hwnd := GetChatGPTWindowHwnd()
        WinActivate "ahk_id " hwnd
    if !WinWaitActive("ahk_exe chrome.exe", , 1)
        return
    CenterMouse()

    try {
        cUIA := UIA_Browser()
        Sleep (IS_WORK_ENVIRONMENT ? 150 : 300)
    } catch {
        ShowNotification("Unable to connect to browser", 1500, "DF2935", "FFFFFF")
        return
    }

    ; Name sets (case-insensitive)
    moreNames := ["More actions", "More options", "Mais ações", "Mais opções"]
    readNames := ["Read aloud", "Ler em voz alta"]
    copyNames := ["Copy to clipboard", "Copiar para a área de transferência", "Copy", "Copiar"]

    CollectPreferExact := (root, arr) => (
        elsExact := CollectByNamesExact(root, arr),
        elsExact.Length ? elsExact : CollectByNamesContains(root, arr)
    )

    ; Try to infer latest assistant message using last Copy button occurrence
    copyCandidates := CollectPreferExact(cUIA, copyNames)
    targetContainer := ""
    isCopied := false
    if (copyCandidates.Length) {
        lastCopy := copyCandidates[copyCandidates.Length]
        current := lastCopy
        loop 12 {
            try {
                current := current.Parent
                if !IsObject(current)
                    break
                cand := CollectPreferExact(current, moreNames)
                if (cand.Length) {
                    targetContainer := current
                    break
                }
            } catch {
                break
            }
        }

        ; Click the last Copy button to copy latest assistant message
        try {
            lastCopy.ScrollIntoView()
            Sleep (IS_WORK_ENVIRONMENT ? 50 : 100)
            A_Clipboard := ""
            SafeClick(lastCopy)
            if ClipWait(1)
                isCopied := true
        } catch {
            ; Silently ignore errors
        }
    }

    Sleep(IS_WORK_ENVIRONMENT ? 150 : 300)

    ; Find and click More actions
    moreBtns := []
    if (IsObject(targetContainer))
        moreBtns := CollectPreferExact(targetContainer, moreNames)
    if !(moreBtns.Length)
        moreBtns := CollectPreferExact(cUIA, moreNames) ; fallback to latest message in thread

    if !(moreBtns.Length) {
        ShowNotification("More btns not found", 1500, "3772FF", "FFFFFF")
        return
    }

    try {
        btn := moreBtns[moreBtns.Length]
        btn.ScrollIntoView()
        Sleep (IS_WORK_ENVIRONMENT ? 40 : 80)
        SafeClick(btn)
    } catch {
        ShowNotification("More btns not found", 1500, "3772FF", "FFFFFF")
        return
    }

    ; From the opened menu, click Read aloud
    Sleep (IS_WORK_ENVIRONMENT ? 60 : 120)
    readItems := CollectPreferExact(cUIA, readNames)
    if !(readItems.Length) {
        ShowNotification("Read Aloud not found", 1500, "3772FF", "FFFFFF")
        return
    }

    try {
        readItem := readItems[readItems.Length]
        readItem.ScrollIntoView()
        Sleep (IS_WORK_ENVIRONMENT ? 30 : 60)
        SafeClick(readItem)
    } catch {
        ShowNotification("Read Aloud not found", 1500, "3772FF", "FFFFFF")
        return
    }

    ; Sleep(IS_WORK_ENVIRONMENT ? 150 : 300)

    ; Send("{Escape}") ;

    ; Sleep(IS_WORK_ENVIRONMENT ? 25 : 50)

    ; Send("{Media_Play_Pause}")

    ; Sleep(IS_WORK_ENVIRONMENT ? 25 : 50)

    Sleep(150)

    Send("!{Tab}") ; Send Shift+Tab to move focus backward

    Sleep(IS_WORK_ENVIRONMENT ? 150 : 150)

    ; Show banner if copied successfully (stay in ChatGPT)
    if (isCopied) {
        ShowNotification("Message copied and reading started!")
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
    if hwnd := GetChatGPTWindowHwnd() {
        WinActivate "ahk_id " hwnd
        WinWaitActive "ahk_id " hwnd
    }
    CenterMouse()
    try {
        cUIA := UIA_Browser()
        Sleep (IS_WORK_ENVIRONMENT ? 150 : 300)
    } catch {
        ShowNotification("Unable to connect to browser", 1500, "DF2935", "FFFFFF")
        return
    }

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
                ShowNotification("Could not start/stop voice mode", 1500, "DF2935", "FFFFFF")
            }
        } catch as e {
            if !triedFallback {
                ToggleVoiceMode(true, "stop")
                return
            } else {
                ShowNotification("Error starting/stopping voice mode", 1500, "DF2935", "FFFFFF")
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
                ShowNotification("Could not stop/start voice mode", 1500, "DF2935", "FFFFFF")
            }
        } catch as e {
            if !triedFallback {
                ToggleVoiceMode(true, "start")
                return
            } else {
                ShowNotification("Error stopping/starting voice mode", 1500, "DF2935", "FFFFFF")
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

EnsureMicVolume100() {
    static lastRunTick := 0
    if (A_TickCount - lastRunTick < 5000)
        return
    lastRunTick := A_TickCount
    ps1Path := A_ScriptDir "\Set-MicVolume.ps1"
    cmd := 'powershell.exe -ExecutionPolicy Bypass -File "' ps1Path '" -Level 100'
    try {
        RunWait cmd, , "Hide"
    } catch {
        ; Silently ignore mic volume errors, still debounced
    }
}

; =============================================================================
; Toggle Dictation (No Auto-Send)
; Hotkey: Win+Alt+Shift+0
; Original File: ChatGPT - Toggle microphone.ahk
; =============================================================================
#!+0::
{
    ToggleDictation(false)
}

; =============================================================================
; Toggle Dictation (with Auto-Send)
; Hotkey: Win+Alt+Shift+7
; Original File: ChatGPT - Speak.ahk
; =============================================================================
#!+7::
{
    ToggleDictation(true)
}

ToggleDictation(autoSend) {
    static isDictating := false
    static stopErrorRetryCount := 0
    global g_transcribeChimePending

    ; --- Button Names (EN/PT) ---
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

    startNames := [currentDictateName]
    stopNames := [currentSubmitName, currentTranscribingName]

    ; --- Activate Window & UIA ---
    SetTitleMatchMode 2
    if hwnd := GetChatGPTWindowHwnd()
        WinActivate "ahk_id " hwnd
    if !WinWaitActive("ahk_exe chrome.exe", , 2)
        return
    CenterMouse()
    try {
        cUIA := UIA_Browser()
        Sleep (IS_WORK_ENVIRONMENT ? 150 : 300)
    } catch {
        ShowNotification("Unable to connect to browser", 1500, "DF2935", "FFFFFF")
        return
    }

    action := !isDictating ? "start" : "stop"

    if (action = "start") {
        try {
            if btn := FindButton(cUIA, startNames) {
                EnsureMicVolume100()
                btn.Click()
                isDictating := true
                ; Wait until dictation actually starts: either the start button disappears
                ; or the stop/submit button appears, then play a short start chime
                ; Keep this check brief to avoid blocking UX
                try {
                    detected := false
                    loops := 0
                    while (loops < 20 && !detected) { ; ~3s max (20 * 150ms)
                        ; Check for stop/submit (meaning listening started)
                        stopFound := false
                        for n in stopNames {
                            try {
                                if cUIA.FindElement({ Name: n, Type: "Button" }) {
                                    stopFound := true
                                    break
                                }
                            } catch {
                                ; Silently ignore UIA element errors
                            }
                        }
                        if (stopFound) {
                            detected := true
                            break
                        }
                        ; Or verify the start button is gone
                        startStillThere := false
                        try {
                            if FindButton(cUIA, startNames)
                                startStillThere := true
                        } catch {
                            startStillThere := false
                        }
                        if (!startStillThere) {
                            detected := true
                            break
                        }
                        Sleep (IS_WORK_ENVIRONMENT ? 75 : 150)
                        loops++
                    }
                    if (detected)
                        PlayDictationStartedChime()
                } catch {
                    ; Silently ignore errors
                }
                ; Switch back to the previous window first so the indicator appears there
                Send "!{Tab}"
                Sleep (IS_WORK_ENVIRONMENT ? 150 : 300)    ; ensure the window switch has completed
                ShowDictationIndicator()
                g_transcribeChimePending := false
            } else {
                ShowNotification((IS_WORK_ENVIRONMENT ? "Não foi possível iniciar o ditado" :
                    "Could not start dictation"), 1500, "DF2935", "FFFFFF")
            }
        } catch Error as e {
            ShowNotification((IS_WORK_ENVIRONMENT ? "Erro ao iniciar o ditado" : "Error starting dictation"), 1500,
            "DF2935", "FFFFFF")
        }
    } else if (action = "stop") {
        try {
            if submitBtn := FindButton(cUIA, stopNames) {
                ;EnsureMicVolume100()
                submitBtn.Click()
                isDictating := false
                HideDictationIndicator()
                Send "!{Tab}" ; Return to previous window immediately to allow multitasking

                ; --- Wait for transcription to finish (indicator appears over ChatGPT) ---
                transcribingWaitNames := [currentTranscribingName, currentSubmitName]
                ; Set a one-time flag so the watcher can emit a distinct chime right before closing
                g_transcribeChimePending := true  ; Always play transcription finished chime
                WaitForButtonAndShowSmallLoading(transcribingWaitNames, "Transcribing…")

                if (autoSend) {
                    try {
                        ; The 'Send prompt' button should appear after transcription
                        finalSendBtn := cUIA.WaitElement({ Name: currentSendPromptName, AutomationId: "composer-submit-button" },
                        8000)
                        if finalSendBtn {
                            finalSendBtn.Click() ; Click instead of sending {Enter}
                            Send "!{Tab}" ; Return to previous window immediately to allow multitasking
                            ; --- Show smaller green loading indicator while ChatGPT is responding ---
                            WaitForButtonAndShowSmallLoading([currentStopStreamingName], "AI is responding…", 180000)
                        } else {
                            ShowNotification("Timeout: Send button did not appear", 1500, "DF2935", "FFFFFF")
                        }
                    } catch Error as e_wait {
                        ShowNotification("Error waiting for send button", 1500, "DF2935", "FFFFFF")
                    }
                }
            } else {
                ShowNotification((IS_WORK_ENVIRONMENT ? "Não foi possível parar o ditado" :
                    "Could not stop dictation"), 1500, "DF2935", "FFFFFF")
                isDictating := false ; Reset state if stop button is not found
            }
        } catch Error as e {
            ; Instead of interrupting with a modal dialog, briefly switch back to the user's window
            ; so the banner appears on the monitor with the active window, then show a quick blue banner
            Send "!{Tab}"
            Sleep (IS_WORK_ENVIRONMENT ? 125 : 250)
            ShowNotification(IS_WORK_ENVIRONMENT ? "Reiniciando ditado…" : "Restarting dictation…", 1200, "3772FF",
                "FFFFFF")
            isDictating := false ; Reset state so we can attempt a fresh start

            stopErrorRetryCount++
            if (stopErrorRetryCount <= 3) {
                ; Attempt to restart dictation quickly
                try {
                    SetTitleMatchMode 2
                    if hwnd := GetChatGPTWindowHwnd() {
                        WinActivate "ahk_id " hwnd
                        WinWaitActive "ahk_id " hwnd
                    }
                    CenterMouse()
                    try {
                        cUIA_restart := UIA_Browser()
                        Sleep (IS_WORK_ENVIRONMENT ? 100 : 200)
                    } catch {
                        ; Silently ignore restart UIA errors
                        return
                    }
                    if btnRestart := FindButton(cUIA_restart, startNames, "Button", 3000) {
                        EnsureMicVolume100()
                        btnRestart.Click()
                        isDictating := true
                        Send "!{Tab}"
                        Sleep (IS_WORK_ENVIRONMENT ? 150 : 300)
                        ShowDictationIndicator()
                        g_transcribeChimePending := false
                        stopErrorRetryCount := 0 ; success – reset counter
                    } else {
                        ; Could not find start button – give up silently this time
                        ; (Leave isDictating := false so the user can try again)
                    }
                } catch {
                    ; Silently ignore restart errors – if they persist, the retry cap below will stop it
                }
            } else {
                ; Too many consecutive failures – stop trying
                ShowNotification(IS_WORK_ENVIRONMENT ? "Falhas repetidas no ditado — parando" :
                    "Repeated dictation failures — stopping", 1800, "DF2935", "FFFFFF")
                stopErrorRetryCount := 0
            }
        }
    }
}

; =============================================================================
; Small Loading Indicator Helpers (for Transcription)
; =============================================================================
global smallLoadingGuis := [] ; Changed to array

ShowSmallLoadingIndicator(state := "Loading…", bgColor := "3772FF") {
    global smallLoadingGuis

    ; If GUIs exist, just update the text of the topmost one (the message)
    if (smallLoadingGuis.Length > 0) {
        try {
            ; The text control is expected to be in the first GUI of the stack
            if (smallLoadingGuis[1].Controls.Length > 0)
                smallLoadingGuis[1].Controls[1].Text := state
        } catch {
            ; Silently handle GUI/control errors and recreate
        }
        return
    }

    ; Create a single, high-contrast, centered banner using the unified builder
    textGui := CreateCenteredBanner(state, bgColor, "FFFFFF", 24, 178)
    smallLoadingGuis.Push(textGui)
}

HideSmallLoadingIndicator() {
    global smallLoadingGuis
    if (smallLoadingGuis.Length > 0) {
        for gui in smallLoadingGuis {
            try gui.Destroy()
            catch {
                ; Silently ignore GUI destroy errors
            }
        }
        smallLoadingGuis := [] ; Reset the array
    }
}

WaitForButtonAndShowSmallLoading(buttonNames, stateText := "Loading…", timeout := 15000) {
    try cUIA := UIA_Browser()
    catch {
        ; Silently ignore UIA browser errors
        return
    }
    start := A_TickCount
    deadline := (timeout > 0) ? (start + timeout) : 0
    btn := ""
    indicatorShown := false
    buttonEverFound := false
    buttonDisappeared := false
    while (timeout <= 0 || A_TickCount < deadline) {
        btn := ""
        for n in buttonNames {
            try {
                btn := cUIA.FindElement({ Name: n, Type: "Button" })
            } catch {
                btn := ""
            }
            if btn
                break
        }
        if btn {
            buttonEverFound := true
            if (!indicatorShown) {
                ShowSmallLoadingIndicator(stateText)
                indicatorShown := true
            }
            while btn && (timeout <= 0 || A_TickCount < deadline) {
                Sleep (IS_WORK_ENVIRONMENT ? 125 : 250)
                btn := ""
                for n in buttonNames {
                    try {
                        btn := cUIA.FindElement({ Name: n, Type: "Button" })
                    } catch {
                        btn := ""
                    }
                    if btn
                        break
                }
            }
            if !btn
                buttonDisappeared := true
            break
        }
        Sleep (IS_WORK_ENVIRONMENT ? 125 : 250)
    }
    ; Play completion sound only for actual AI responses when we saw the button and it disappeared
    try {
        if (buttonEverFound && buttonDisappeared && InStr(StrLower(stateText), "transcrib") = 0)
            PlayCompletionChime()
    } catch {
        ; Silently ignore errors
    }
    ; If transcription-finished chime is pending, fire only if we observed the transcribing button disappear
    try {
        global g_transcribeChimePending
        if (g_transcribeChimePending && buttonEverFound && buttonDisappeared) {
            g_transcribeChimePending := false
            PlayTranscriptionFinishedChime()
        }
    } catch {
        ; Silently ignore errors
    }
    HideSmallLoadingIndicator()
}

; =============================================================================
; Copy Last Code
; Hotkey: Win+Alt+Shift+U
; =============================================================================
#!+u::
{
    SetTitleMatchMode(2)
    A_Clipboard := "" ; Empty clipboard to check for new content later

    if hwnd := GetChatGPTWindowHwnd() {
        WinActivate("ahk_id " hwnd)
        if WinWaitActive("ahk_id " hwnd, , 2)
            CenterMouse()
    } else {
        Run "chrome.exe --new-window https://chatgpt.com/"
        if !WinWaitActive("ahk_exe chrome.exe", , 3)
            return
        CenterMouse()
    }
    Sleep (IS_WORK_ENVIRONMENT ? 150 : 300)
    Send("^+;")

    Send("!{Tab}") ; Switch back to the previous window
    Sleep(IS_WORK_ENVIRONMENT ? 150 : 300) ; Wait for the window switch

    if ClipWait(1) {
        ShowNotification("Last code block copied!")
    }
}

; =============================================================================
; Helper function to center mouse on the active window
; =============================================================================
CenterMouse() {
    Sleep(IS_WORK_ENVIRONMENT ? 100 : 200)
    Send("#!+q")
}

; =============================================================================
; Unified banner builder – consistent shape/font/opacity for all banners here
; =============================================================================
CreateCenteredBanner(message, bgColor := "3772FF", fontColor := "FFFFFF", fontSize := 24, alpha := 178) {
    bGui := Gui()
    bGui.Opt("+AlwaysOnTop -Caption +ToolWindow")
    bGui.BackColor := bgColor
    bGui.SetFont("s" . fontSize . " c" . fontColor . " Bold", "Segoe UI")
    bGui.Add("Text", "w500 Center", message)

    activeWin := WinGetID("A")
    if (activeWin) {
        WinGetPos(&winX, &winY, &winW, &winH, activeWin)
    } else {
        workArea := SysGet.MonitorWorkArea(SysGet.MonitorPrimary)
        winX := workArea.Left, winY := workArea.Top, winW := workArea.Right - workArea.Left, winH := workArea.Bottom -
            workArea.Top
    }

    bGui.Show("AutoSize Hide")
    guiW := 0, guiH := 0
    bGui.GetPos(, , &guiW, &guiH)

    guiX := winX + (winW - guiW) / 2
    guiY := winY + (winH - guiH) / 2
    bGui.Show("x" . Round(guiX) . " y" . Round(guiY) . " NA")
    WinSetTransparent(alpha, bGui)
    return bGui
}

; =============================================================================
; Helper function to show a notification on the active window
; =============================================================================
ShowNotification(message, durationMs := 500, bgColor := "FFFF00", fontColor := "000000", fontSize := 24) {
    notificationGui := CreateCenteredBanner(message, bgColor, fontColor, fontSize, 178)
    Sleep(durationMs)
    if IsObject(notificationGui) && notificationGui.Hwnd {
        notificationGui.Destroy()
    }
}

; =============================================================================
; Helper functions to show/hide the persistent dictation indicator
; =============================================================================
ShowDictationIndicator(message := "Dictation ON") {
    ShowSmallLoadingIndicator(message, "FF0000")
}

HideDictationIndicator() {
    HideSmallLoadingIndicator()
}

; =============================================================================
; Helper function to click the first conversation's options (three-dots) button
; =============================================================================
ClickFirstConversationOptions(timeoutMs := 5000) {
    try {
        cUIA := UIA_Browser()
    } catch {
        ShowNotification("Unable to connect to browser", 1500, "DF2935", "FFFFFF")
        return false
    }

    ; --- 1) locate the "Chats / Conversas" sidebar container (first match) ---
    sidebar := ""
    for grp in cUIA.FindAll({ Type: "Group" }) {
        n := Trim(StrLower(grp.Name))
        if (InStr(n, "chats") || InStr(n, "conversas")) {
            sidebar := grp
            break
        }
    }
    if !sidebar {
        ; Fallback: operate on the whole page root if sidebar not detected
        sidebar := cUIA
    }

    ; --- 2) prepare lookup tables ---
    optNames := [
        ; English
        "Open conversation options", "Conversation options", "Open options", "More options",
        ; Portuguese (Brazil)
        "Abrir opções da conversa", "Abrir opções de conversa", "Opções da conversa", "Mais opções"
    ]

    ; Include Link as another possible control type
    optTypes := ["Button", "MenuItem", "MenuButton", "SplitButton", "Custom", "Link"]

    ; --- 3) iterative scan until timeout elapses ---
    startTick := A_TickCount
    btn := ""
    while ((A_TickCount - startTick) < timeoutMs) {
        for name in optNames {
            for role in optTypes {
                try {
                    btn := sidebar.FindElement({ Name: name, Type: role, matchmode: "Substring" })
                } catch {
                    btn := ""
                }
                if btn {
                    break 2  ; leave both loops
                }
            }
        }
        if btn
            break
        Sleep (IS_WORK_ENVIRONMENT ? 60 : 120)
    }

    ; --- 4) broader fallback search (any type, just keywords) ---
    if !btn {
        broadTerms := ["conversation options", "opções de conversa", "opcões", "options", "opções"]
        for term in broadTerms {
            try {
                btn := sidebar.FindElement({ Name: term, matchmode: "Substring" })
            } catch {
                btn := ""
            }
            if btn
                break
        }
    }

    ; --- 5) if still not found, dump candidate list for debugging ---
    if !btn {
        list := ""
        for el in sidebar.FindAll({ Type: optTypes })
            list .= el.Type "  |  " el.Name "`n"
        FileAppend list, "*options-scan.log", "UTF-8"
        ShowNotification("Options control not found", 1500, "DF2935", "FFFFFF")
        return false
    }

    ; --- 6) success – scroll into view and click ---
    try {
        btn.ScrollIntoView()
        Sleep (IS_WORK_ENVIRONMENT ? 50 : 100)
        btn.Click()
        return true
    } catch Error as e {
        ShowNotification("Error clicking options button", 1500, "DF2935", "FFFFFF")
        return false
    }
}

; =============================================================================
; Ensure the ChatGPT sidebar (Chats list) is open – toggles Ctrl+Shift+S if needed
; =============================================================================
EnsureSidebarOpen(timeoutMs := 3000) {
    try {
        cUIA := UIA_Browser()
    } catch {
        return false
    }
    anchorNames := ["Chats", "Conversas"]

    if SidebarVisible(cUIA, anchorNames)
        return true    ; already open

    ; sidebar hidden – toggle with Ctrl+Shift+S
    Send "^+s"
    Sleep (IS_WORK_ENVIRONMENT ? 300 : 600)

    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        if SidebarVisible(cUIA, anchorNames)
            return true
        Sleep (IS_WORK_ENVIRONMENT ? 100 : 200)
    }
    return false   ; still hidden
}

; Helper – detects if the Chats sidebar is visible by looking for the anchor label
SidebarVisible(uiRoot, names) {
    for n in names {
        try {
            if uiRoot.FindElement({ Name: n, Type: "Text" })
                return true
        } catch {
            ; Silently ignore errors when element not found
        }
    }
    return false
}

; =============================================================================
; Send Ctrl+A, Sleep, Ctrl+X to ChatGPT
; Hotkey: Win+Alt+Shift+J
; =============================================================================
#!+j::
{
    SetTitleMatchMode(2)
    if hwnd := GetChatGPTWindowHwnd()
        WinActivate("ahk_id " hwnd)
    if WinWaitActive("ahk_exe chrome.exe", , 2) {
        Send("{Esc}")
        Send("F")
        Send("{Backspace}")
        Sleep(IS_WORK_ENVIRONMENT ? 150 : 300)
        Send("^a")
        Sleep(IS_WORK_ENVIRONMENT ? 175 : 350)
        Send("^x")
        Send("!{Tab}")
    }
}

; =============================================================================
; Loading indicator helpers (yellow, small) – original
; =============================================================================
ShowLoading(message := "Loading…", bgColor := "FFFF00", fontColor := "000000", fontSize := 26) {
    global loadingGui
    ; If already visible, update text
    if (IsObject(loadingGui) && loadingGui.Hwnd) {
        loadingGui.Controls[1].Text := message
        return
    }
    loadingGui := Gui()
    loadingGui.Opt("+AlwaysOnTop -Caption +ToolWindow")
    loadingGui.BackColor := bgColor
    loadingGui.SetFont("s" . fontSize . " c" . fontColor . " Bold", "Segoe UI")
    loadingGui.Add("Text", "w500 Center", message)

    ; Center on primary monitor
    MonitorGetWorkArea(1, &l, &t, &r, &b)
    winW := r - l
    winH := b - t
    loadingGui.Show("AutoSize Hide")
    loadingGui.GetPos(, , &guiW, &guiH)
    guiX := l + (winW - guiW) / 2
    guiY := t + (winH - guiH) / 2
    loadingGui.Show("x" . Round(guiX) . " y" . Round(guiY) . " NA")
    WinSetTransparent(178, loadingGui)
    Sleep (IS_WORK_ENVIRONMENT ? 75 : 150)
}

HideLoading() {
    global loadingGui
    if (IsObject(loadingGui) && loadingGui.Hwnd) {
        loadingGui.Destroy()
        loadingGui := ""
    }
}

RecenterLoadingOverWindow(hwnd) {
    global loadingGui
    if !(IsObject(loadingGui) && loadingGui.Hwnd)
        return
    if !WinExist(hwnd)
        return

    WinGetPos(&wx, &wy, &ww, &wh, hwnd)
    loadingGui.GetPos(, , &gw, &gh)
    gx := wx + (ww - gw) / 2
    gy := wy + (wh - gh) / 2
    loadingGui.Show("x" . Round(gx) . " y" . Round(gy) . " NA")
}

; =============================================================================
; Completion chime (single beep, debounced)
; =============================================================================
PlayCompletionChime() {
    try {
        static lastTick := 0
        if (A_TickCount - lastTick < 1500)
            return
        lastTick := A_TickCount

        played := false
        ; Prefer Windows MessageBeep (reliable through default output)
        try {
            rc := DllCall("User32\\MessageBeep", "UInt", 0xFFFFFFFF)
            if (rc)
                played := true
        } catch {
            ; Silently ignore errors
        }

        ; Fallback to system asterisk sound
        if !played {
            try {
                played := SoundPlay("*64", false)
            } catch {
                ; Silently ignore errors
            }
        }

        ; Last resort, attempt the classic beep
        if !played {
            try SoundBeep(1100, 130)
            catch {
            }
        }
    } catch {
        ; Silently ignore errors
    }
}

PlayTranscriptionFinishedChime() {
    try {
        static lastTick := 0
        if (A_TickCount - lastTick < 2000)
            return
        lastTick := A_TickCount

        played := false
        ; Prefer a distinct Windows MessageBeep variant (warning icon)
        try {
            rc := DllCall("User32\\MessageBeep", "UInt", 0x00000030)
            if (rc)
                played := true
        } catch {
            ; Silently ignore errors
        }

        ; Fallback to system exclamation sound
        if !played {
            try {
                played := SoundPlay("*48", false)
            } catch {
                ; Silently ignore errors
            }
        }

        ; Last resort, a short, higher-pitched beep
        if !played {
            try SoundBeep(1400, 90)
            catch {
            }
        }
    } catch {
        ; Silently ignore errors
    }
}

; =============================================================================
; Dictation started chime (distinct beep when microphone starts listening)
; =============================================================================
PlayDictationStartedChime() {
    try {
        static lastTick := 0
        if (A_TickCount - lastTick < 1000)
            return
        lastTick := A_TickCount

        played := false
        ; Use information icon beep to distinguish from other cues
        try {
            rc := DllCall("User32\\MessageBeep", "UInt", 0x00000040)
            if (rc)
                played := true
        } catch {
            ; Silently ignore errors
        }

        if !played {
            try {
                played := SoundPlay("*16", false) ; system hand/stop sound as fallback
            } catch {
                ; Silently ignore errors
            }
        }

        if !played {
            try SoundBeep(1200, 100)
            catch {
            }
        }
    } catch {
        ; Silently ignore errors
    }
}

; =============================================================================
; Grammar check started chime (distinct beep when grammar check begins)
; =============================================================================
PlayGrammarCheckStartedChime() {
    try {
        static lastTick := 0
        if (A_TickCount - lastTick < 1000)
            return
        lastTick := A_TickCount

        played := false
        ; Use question icon beep to distinguish from other cues
        try {
            rc := DllCall("User32\\MessageBeep", "UInt", 0x00000020)
            if (rc)
                played := true
        } catch {
            ; Silently ignore errors
        }

        if !played {
            try {
                played := SoundPlay("*32", false) ; system question sound as fallback
            } catch {
                ; Silently ignore errors
            }
        }

        if !played {
            try SoundBeep(1000, 120)
            catch {
            }
        }
    } catch {
        ; Silently ignore errors
    }
}

; =============================================================================
; Grammar check completed chime (distinct beep when grammar check finishes)
; =============================================================================
PlayGrammarCheckCompletedChime() {
    try {
        static lastTick := 0
        if (A_TickCount - lastTick < 1500)
            return
        lastTick := A_TickCount

        played := false
        ; Use asterisk icon beep to distinguish from other cues
        try {
            rc := DllCall("User32\\MessageBeep", "UInt", 0x00000040)
            if (rc)
                played := true
        } catch {
            ; Silently ignore errors
        }

        if !played {
            try {
                played := SoundPlay("*64", false) ; system asterisk sound as fallback
            } catch {
                ; Silently ignore errors
            }
        }

        if !played {
            try SoundBeep(1300, 150)
            catch {
            }
        }
    } catch {
        ; Silently ignore errors
    }
}

; =============================================================================
; Voice Mode watcher: beep when "End voice mode" appears/disappears
; =============================================================================
global g_voiceWatcherOn := false

StartVoiceModeWatcher() {
    global g_voiceWatcherOn
    if g_voiceWatcherOn
        return
    g_voiceWatcherOn := true
    SetTimer(CheckVoiceModeButton, 400)
}

StopVoiceModeWatcher() {
    global g_voiceWatcherOn
    if !g_voiceWatcherOn
        return
    g_voiceWatcherOn := false
    SetTimer(CheckVoiceModeButton, 0)
}

CheckVoiceModeButton() {
    static prevPresent := false
    static lastChangeTick := 0

    hwnd := GetChatGPTWindowHwnd()
    present := false

    if (hwnd) {
        try {
            cUIA := UIA_Browser("ahk_id " hwnd)
            names := ["End voice mode", "Encerrar modo voz"]
            for n in names {
                try {
                    if cUIA.FindElement({ Name: n, Type: "Button" }) {
                        present := true
                        break
                    }
                } catch {
                    ; Silently ignore errors
                }
            }
        } catch {
            ; Silently ignore errors
        }
    }

    if (present != prevPresent) {
        if (A_TickCount - lastChangeTick > 800) {
            lastChangeTick := A_TickCount
            if (present) {
                ; started talking
                PlayDictationStartedChime()
            } else {
                ; finished talking
                PlayCompletionChime()
            }
        }
        prevPresent := present
    }
}

; Initialize comprehensive sound suppression and start watcher on script load
SuppressErrorSounds()
DisableSystemSounds()
StartVoiceModeWatcher()

; Re-enable sounds when script exits
OnExit(RestoreSoundsOnExit)

RestoreSoundsOnExit(*) {
    RestoreErrorSounds()
    EnableSystemSounds()
}
