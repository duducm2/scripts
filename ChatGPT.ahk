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

; --- Auto-Restart Dictation Manager ----------------------------------------
; State machine: IDLE | ACTIVE | RESTARTING | ERROR
global g_dictationState := "IDLE"
global g_dictationStartTime := 0
global g_autoRestartEnabled := false
global g_autoRestartTimer := ""
global g_autoSendMode := false
global g_hotkeyLock := false
global g_lastHotkeyTime := 0
global CONST_MAX_DICTATION_MS := 30000  ; 30 seconds
global CONST_HOTKEY_DEBOUNCE_MS := 500

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
        Sleep(150)
    } catch {
        ShowNotification("Unable to connect to browser", 1500, "DF2935", "FFFFFF")
        return
    }

    ; Go to the end of the thread
    Send("^{End}")
    Sleep(150)

    ; Name sets (case-insensitive)
    ; Include Portuguese variants and ASCII fallbacks (without diacritics)
    moreNames := ["More actions", "More options", "Mais ações", "Mais acoes", "Mais opções", "Mais opcoes"]
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
            Sleep(50)
            A_Clipboard := ""
            SafeClick(lastCopy)
            if ClipWait(1)
                isCopied := true
        } catch {
            ; Silently ignore errors
        }
    }

    Sleep(150)

    ; Find and click More actions
    moreBtns := []
    if (IsObject(targetContainer))
        moreBtns := CollectPreferExact(targetContainer, moreNames)
    if !(moreBtns.Length)
        moreBtns := CollectPreferExact(cUIA, moreNames) ; fallback to latest message in thread

    ; Extra robust fallback: search by Button class/automation id hints seen in PT-BR UI
    if !(moreBtns.Length) {
        try {
            for el in cUIA.FindAll({ Type: "Button" }) {
                n := StrLower(Trim(el.Name))
                cls := el.ClassName
                aid := el.AutomationId
                if (InStr(n, "mais a") || InStr(n, "more a")) {
                    moreBtns.Push(el)
                    continue
                }
                if (InStr(cls, "text-token-text-secondary") || InStr(cls, "hover:bg-token-bg-secondary") || InStr(aid,
                    "radix-"))
                    moreBtns.Push(el)
            }
        } catch {
        }
    }

    if !(moreBtns.Length) {
        ShowNotification("More btns not found", 1500, "3772FF", "FFFFFF")
        return
    }

    try {
        btn := moreBtns[moreBtns.Length]
        btn.ScrollIntoView()
        Sleep(40)
        SafeClick(btn)
    } catch {
        ShowNotification("More btns not found", 1500, "3772FF", "FFFFFF")
        return
    }

    ; From the opened menu, click Read aloud
    Sleep(300)
    readItems := CollectPreferExact(cUIA, readNames)
    if !(readItems.Length) {
        ; Fallback: scan visible menuitems/buttons for Portuguese text
        try {
            for el in cUIA.FindAll({ Type: ["MenuItem", "Button"] }) {
                n := StrLower(Trim(el.Name))
                if (InStr(n, "read aloud") || InStr(n, "ler em voz alta"))
                    readItems.Push(el)
            }
        } catch {
        }
        if !(readItems.Length) {
            ShowNotification("Read Aloud not found", 1500, "3772FF", "FFFFFF")
            return
        }
    }

    try {
        readItem := readItems[readItems.Length]
        readItem.ScrollIntoView()
        Sleep(30)
        SafeClick(readItem)
    } catch {
        ShowNotification("Read Aloud not found", 1500, "3772FF", "FFFFFF")
        return
    }

    Sleep(700)

    Send("!{Tab}") ; Send Shift+Tab to move focus backward

    Sleep(150)

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

; Wait until the ChatGPT composer submit button becomes available
; Uses AutomationId: "composer-submit-button"; falls back to Name if needed
WaitForComposerSubmitButton(timeoutMs := 30000) {
    try {
        hwnd := GetChatGPTWindowHwnd()
        if !hwnd
            return false
        cUIA := UIA_Browser("ahk_id " hwnd)
    } catch {
        return false
    }

    names := ["Send prompt", "Enviar prompt", "Send", "Enviar"]
    start := A_TickCount
    deadline := (timeoutMs > 0) ? (start + timeoutMs) : 0
    while (timeoutMs <= 0 || A_TickCount < deadline) {
        btn := ""
        ; Prefer AutomationId match
        try btn := cUIA.FindElement({ AutomationId: "composer-submit-button", Type: "Button" })
        catch {
            btn := ""
        }
        if !btn {
            ; Fallback to name variants
            for n in names {
                try {
                    btn := cUIA.FindElement({ Name: n, Type: "Button" })
                } catch {
                    btn := ""
                }
                if btn
                    break
            }
        }
        if btn {
            ; Ensure it's visible and enabled before proceeding
            ok := true
            try {
                ok := (!btn.IsOffscreen) && btn.IsEnabled
            } catch {
                ok := true
            }
            if ok
                return true
        }
        Sleep (IS_WORK_ENVIRONMENT ? 100 : 200)
    }
    return false
}

EnsureMicVolume100() {
    static lastRunTick := 0
    if (A_TickCount - lastRunTick < 5000)
        return
    lastRunTick := A_TickCount
    ps1Path := A_ScriptDir "\Set-MicVolume.ps1"
    cmd := 'powershell.exe -ExecutionPolicy Bypass -File "' ps1Path '" -Level 100'
    try {
        Run cmd, , "Hide"  ; Run asynchronously - don't wait for it to complete
    } catch {
        ; Silently ignore mic volume errors, still debounced
    }
}

; =============================================================================
; Toggle Dictation (No Auto-Send) - WITH AUTO-RESTART
; Hotkey: Win+Alt+Shift+0
; Original File: ChatGPT - Toggle microphone.ahk
; =============================================================================
#!+0::
{
    HandleDictationToggleWithAutoRestart(false)
}

; =============================================================================
; Toggle Dictation (with Auto-Send) - WITH AUTO-RESTART
; Hotkey: Win+Alt+Shift+7
; Original File: ChatGPT - Speak.ahk
; =============================================================================
#!+7::
{
    HandleDictationToggleWithAutoRestart(true)
}

ToggleDictation(autoSend, source := "manual") {
    static isDictating := false
    global g_transcribeChimePending
    global g_dictationState
    global g_autoRestartEnabled
    global g_autoRestartTimer
    global g_autoSendMode
    global g_dictationStartTime

    ; --- Activate Window ---
    SetTitleMatchMode 2
    if hwnd := GetChatGPTWindowHwnd()
        WinActivate "ahk_id " hwnd
    if !WinWaitActive("ahk_exe chrome.exe", , 2)
        return

    action := !isDictating ? "start" : "stop"
    g_autoSendMode := autoSend

    if (action = "start") {
        try {
            ; Navigate to text field: Esc, type 'd', backspace
            Send "{Esc}"
            Sleep 100
            Send "d"
            Sleep 100
            Send "{Backspace}"
            Sleep 100

            ; Press Tab twice to focus on dictation button
            Send "{Tab 2}"

            ; Press Enter to start dictation (don't wait for mic volume check)
            Send "{Enter}"
            isDictating := true
            g_dictationState := "ACTIVE"

            ; Start auto-restart timer if enabled
            if (g_autoRestartEnabled || source = "auto_restart") {
                StartAutoRestartTimer()
            }

            ; Start dictation chime and volume check in parallel (don't wait)
            PlayDictationStartedChime()
            EnsureMicVolume100()  ; This runs asynchronously - starts in background

            ; Switch back to previous window and show indicator
            Send "!{Tab}"
            Sleep (IS_WORK_ENVIRONMENT ? 150 : 300)
            ShowDictationIndicator()
        } catch Error as e {
            g_dictationState := "ERROR"
            ShowNotification((IS_WORK_ENVIRONMENT ? "Erro ao iniciar o ditado" : "Error starting dictation"), 1500,
            "DF2935", "FFFFFF")
        }
    } else if (action = "stop") {
        try {
            ; Cancel auto-restart timer
            StopAutoRestartTimer()
            
            ; Stop the square indicator timer immediately
            ; This prevents the timer from recreating the square after we hide it
            HideDictationIndicator()

            ; Return to ChatGPT window
            if hwnd := GetChatGPTWindowHwnd()
                WinActivate "ahk_id " hwnd

            ; calibrate here
            Sleep 200

            ; Press Enter to stop/pause dictation
            Send "{Enter}"
            isDictating := false

            ; Don't hide indicator yet - keep it visible until transcription finishes
            ; (Unless it's a manual stop, in which case we already hid it above)
            Send "!{Tab}" ; Return to previous window

            ; Wait for transcription to finish by monitoring for the submit button
            ; The submit button appears when transcription is complete
            transcriptionFinished := WaitForComposerSubmitButton(30000)

            ; Special handling for autosend mode with auto-restart enabled
            ; Instead of sending, restart dictation immediately
            if (autoSend && g_autoRestartEnabled && source != "manual" && transcriptionFinished) {
                ; Don't hide indicator or send - restart instead
                ; Small delay before restart
                Sleep(500)
                ; Restart dictation directly (don't call ExecuteAutoRestartSequence to avoid recursion)
                ; Set state to ACTIVE so timer can start properly
                g_dictationState := "ACTIVE"
                ; Start new dictation (timer will be started by ToggleDictation)
                ToggleDictation(false, "auto_restart")
                return
            }

            ; If we're in a restart sequence and this is an automatic restart (not manual),
            ; don't hide indicator or play chime - let the restart sequence continue
            ; But if this is a manual stop, allow it to proceed normally
            if (g_dictationState = "RESTARTING" && source != "manual") {
                ; Keep indicator visible, don't play chime
                ; Just return - the restart sequence will handle continuation
                return
            }

            ; Normal flow: hide indicator and handle autosend
            ; (For manual stops, we already hid it above, but ensure it's hidden here too as safety)
            if (transcriptionFinished) {
                ; Transcription finished - hide indicator and play chime
                HideDictationIndicator()
                PlayTranscriptionFinishedChime()
            } else {
                ; Timeout - hide indicator anyway
                HideDictationIndicator()
            }

            ; If auto-send is enabled (and not auto-restart mode), wait a bit then send the prompt
            if (autoSend && (!g_autoRestartEnabled || source = "manual")) {
                ; Return to ChatGPT to send
                if hwnd := GetChatGPTWindowHwnd()
                    WinActivate "ahk_id " hwnd

                Sleep 100

                ; Press Enter to send the transcribed text
                Send "{Enter}"
                Send "!{Tab}" ; Return to previous window

                ; Wait for the AI answering phase and chime on completion
                buttonNames := ["Stop streaming", "Interromper transmissão"]
                WaitForButtonAndShowSmallLoading(buttonNames, "Waiting for response...")
            }

            ; Update state
            if (g_dictationState != "RESTARTING") {
                g_dictationState := "IDLE"
            }
        } catch Error as e {
            HideDictationIndicator() ; Ensure indicator is hidden on error
            g_dictationState := "ERROR"
            ShowNotification(IS_WORK_ENVIRONMENT ? "Erro ao parar o ditado" : "Error stopping dictation", 1500,
                "DF2935", "FFFFFF")
        }
    }
}

; =============================================================================
; Auto-Restart Dictation Manager
; =============================================================================

; Handle dictation toggle with auto-restart enabled
HandleDictationToggleWithAutoRestart(autoSend) {
    global g_autoRestartEnabled
    global g_dictationState

    if (g_dictationState = "IDLE") {
        ; Starting dictation - enable auto-restart
        g_autoRestartEnabled := true
        ToggleDictation(autoSend, "manual")
    } else {
        ; Stopping dictation - disable auto-restart
        g_autoRestartEnabled := false
        ToggleDictation(autoSend, "manual")
    }
}

; Start the 30-second auto-restart timer
StartAutoRestartTimer() {
    global g_dictationStartTime
    global g_autoRestartTimer
    global CONST_MAX_DICTATION_MS

    g_dictationStartTime := A_TickCount
    ; One-shot timer (negative value = one-time execution)
    g_autoRestartTimer := SetTimer(OnAutoRestartTimerExpired, -CONST_MAX_DICTATION_MS)
}

; Stop the auto-restart timer
StopAutoRestartTimer() {
    global g_autoRestartTimer
    if (g_autoRestartTimer) {
        SetTimer(g_autoRestartTimer, 0)  ; Disable timer
        g_autoRestartTimer := ""
    }
}

; Timer callback when 30 seconds expires
OnAutoRestartTimerExpired() {
    global g_dictationState
    global g_autoRestartEnabled

    if (g_dictationState = "ACTIVE" && g_autoRestartEnabled) {
        g_dictationState := "RESTARTING"
        ExecuteAutoRestartSequence()
    }
}

; Execute the automatic restart sequence
ExecuteAutoRestartSequence() {
    global g_dictationState
    global g_autoRestartEnabled
    global g_hotkeyLock
    global g_lastHotkeyTime
    global CONST_HOTKEY_DEBOUNCE_MS

    ; Debounce check - prevent rapid-fire restarts
    if ((A_TickCount - g_lastHotkeyTime) < CONST_HOTKEY_DEBOUNCE_MS) {
        ; Too soon, reschedule for later
        StartAutoRestartTimer()
        g_dictationState := "ACTIVE"
        return
    }

    ; Lock check - prevent concurrent execution
    if (g_hotkeyLock) {
        ; Already executing, reschedule
        StartAutoRestartTimer()
        g_dictationState := "ACTIVE"
        return
    }

    g_hotkeyLock := true
    g_lastHotkeyTime := A_TickCount

    ; Play chime to notify user that restart is beginning
    PlayAutoRestartChime()

    try {
        ; Step 1: Stop current dictation (without autosend)
        ; This will trigger the stop action in ToggleDictation
        ; The stop will wait for transcription but keep indicator visible (due to RESTARTING state check)
        ToggleDictation(false, "auto_restart")

        ; Step 2: Wait for transcription to finish (short timeout to avoid long delays)
        ; The WaitForComposerSubmitButton is already called in ToggleDictation stop action
        ; But we need to ensure it completes - use a shorter timeout for restart
        ; Actually, WaitForComposerSubmitButton was already called, so we just wait a bit
        Sleep(300)

        ; Step 3: Small delay to ensure UI is ready
        Sleep(200)

        ; Step 4: Start dictation again (without autosend, auto-restart will be enabled)
        ; Since isDictating is now false, this will start a new dictation
        ; The indicator should still be visible from before
        ToggleDictation(false, "auto_restart")

        ; Step 5: Verify restart succeeded and reset timer
        Sleep(500)
        if (VerifyDictationActive()) {
            g_dictationState := "ACTIVE"
            ; Timer will be started by ToggleDictation start action
        } else {
            g_dictationState := "ERROR"
            g_autoRestartEnabled := false
            HideDictationIndicator()  ; Hide on error
            ShowNotification(IS_WORK_ENVIRONMENT ? "Falha no reinício automático" : "Auto-restart failed", 2000,
                "DF2935", "FFFFFF")
        }
    } catch Error as e {
        g_dictationState := "ERROR"
        g_autoRestartEnabled := false
        HideDictationIndicator()  ; Hide on error
        ShowNotification(IS_WORK_ENVIRONMENT ? "Erro no reinício automático" : "Auto-restart error", 2000, "DF2935",
            "FFFFFF")
    } finally {
        g_hotkeyLock := false
    }
}

; Verify that dictation is actually active
VerifyDictationActive() {
    global smallLoadingGuis
    ; Check if dictation indicator is visible
    return (smallLoadingGuis.Length > 0)
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
; Helper functions to show/hide the persistent dictation indicator with shrinking banner
; =============================================================================
global dictationProgressGui := ""
global dictationProgressTimer := ""
global dictationProgressStartTime := 0
global CONST_DICTATION_PROGRESS_INITIAL_SIZE := 200  ; Start at 200px square
global CONST_DICTATION_PROGRESS_MIN_SIZE := 1  ; Shrink to 1px (1 second remaining)
global CONST_DICTATION_PROGRESS_UPDATE_INTERVAL := 1000  ; Update every 1 second
; Store the centre point where the indicator should be anchored (screen coordinates)
global dictationCenterX := 0
global dictationCenterY := 0

ShowDictationIndicator(message := "Dictation ON") {
    global dictationProgressGui
    global dictationProgressTimer
    global dictationProgressStartTime
    global CONST_MAX_DICTATION_MS
    global CONST_DICTATION_PROGRESS_INITIAL_SIZE
    global dictationCenterX
    global dictationCenterY
    
    ; Clean up any existing indicator
    HideDictationIndicator()
    
    ; Record start time
    dictationProgressStartTime := A_TickCount
    
    ; --- Capture the visual centre using the existing CentreMouse pipeline ---
    ; This leverages the already-working Win+Alt+Shift+Q logic (Utils.ahk) instead
    ; of re-implementing monitor/window maths here.
    try CenterMouse()
    catch {
        ; If the centering hotkey fails for any reason, we still fall back below
    }

    Sleep(150)  ; Give the mouse a moment to reach the window centre

    pt := Buffer(8, 0)
    if (DllCall("GetCursorPos", "ptr", pt)) {
        dictationCenterX := NumGet(pt, 0, "int")
        dictationCenterY := NumGet(pt, 4, "int")
    } else {
        ; Fallback: approximate using active window
        activeWin := 0
        try {
            activeWin := WinGetID("A")
        } catch {
            activeWin := 0
        }
        if (activeWin) {
            try {
                WinGetPos(&wx, &wy, &ww, &wh, activeWin)
                dictationCenterX := wx + (ww / 2)
                dictationCenterY := wy + (wh / 2)
            } catch {
                workArea := SysGet.MonitorWorkArea(SysGet.MonitorPrimary)
                dictationCenterX := workArea.Left + (workArea.Right - workArea.Left) / 2
                dictationCenterY := workArea.Top + (workArea.Bottom - workArea.Top) / 2
            }
        } else {
            workArea := SysGet.MonitorWorkArea(SysGet.MonitorPrimary)
            dictationCenterX := workArea.Left + (workArea.Right - workArea.Left) / 2
            dictationCenterY := workArea.Top + (workArea.Bottom - workArea.Top) / 2
        }
    }

    ; --- 3) Create initial square indicator at full size ---
    CreateDictationSquare(CONST_DICTATION_PROGRESS_INITIAL_SIZE)
    
    ; Start 1-second timer to update square indicator
    dictationProgressTimer := SetTimer(UpdateDictationSquare, CONST_DICTATION_PROGRESS_UPDATE_INTERVAL)
}

; Helper function to get the monitor work area width for the active window
; Uses the same monitor detection logic as WindowManagement.ahk
GetMonitorWidthForActiveWindow() {
    activeWin := 0
    try {
        activeWin := WinGetID("A")
    } catch {
        ; No active window available, use primary monitor work area
        MonitorGetWorkArea 1, &left, &top, &right, &bottom
        return right - left
    }
    
    if (!activeWin) {
        ; No active window, use primary monitor work area
        MonitorGetWorkArea 1, &left, &top, &right, &bottom
        return right - left
    }
    
    ; Get the monitor handle for the active window
    hMon := 0
    try {
        hMon := DllCall("MonitorFromWindow", "ptr", activeWin, "uint", 2, "ptr")
    } catch {
        ; Fallback to primary monitor work area if detection fails
        MonitorGetWorkArea 1, &left, &top, &right, &bottom
        return right - left
    }
    
    ; Find which monitor index matches this handle
    count := MonitorGetCount()
    loop count {
        i := A_Index
        MonitorGet i, &l, &t, &r, &b
        cx := (l + r) // 2
        cy := (t + b) // 2
        point64 := (cy & 0xFFFFFFFF) << 32 | (cx & 0xFFFFFFFF)
        hMonTarget := DllCall("MonitorFromPoint", "int64", point64, "uint", 2, "ptr")
        
        if (hMon = hMonTarget) {
            ; Found the monitor, return its work area width (excludes taskbar)
            MonitorGetWorkArea i, &l, &t, &r, &b
            return r - l
        }
    }
    
    ; Fallback: use primary monitor work area if no match found
    MonitorGetWorkArea 1, &left, &top, &right, &bottom
    return right - left
}

HideDictationIndicator() {
    global dictationProgressGui
    global dictationProgressTimer
    
    ; Stop the timer
    if (dictationProgressTimer) {
        SetTimer(dictationProgressTimer, 0)
        dictationProgressTimer := ""
    }
    
    ; Destroy the GUI
    if (IsObject(dictationProgressGui) && dictationProgressGui.Hwnd) {
        try dictationProgressGui.Destroy()
        catch {
            ; Silently ignore errors
        }
        dictationProgressGui := ""
    }
}

; Helper function to create a shrinking square indicator
; Centers on the mouse position captured via CenterMouse()
CreateDictationSquare(size) {
    global dictationProgressGui
    global dictationCenterX
    global dictationCenterY
    
    ; Destroy existing GUI if it exists
    if (IsObject(dictationProgressGui) && dictationProgressGui.Hwnd) {
        try dictationProgressGui.Destroy()
        catch {
            ; Silently ignore errors
        }
    }
    
    ; Fallback centre if we don't have a captured point
    centerX := dictationCenterX
    centerY := dictationCenterY
    if (centerX = 0 && centerY = 0) {
        ; Approximate using active window or primary monitor
        activeWin := 0
        try {
            activeWin := WinGetID("A")
        } catch {
            activeWin := 0
        }
        if (activeWin) {
            try {
                WinGetPos(&wx, &wy, &ww, &wh, activeWin)
                centerX := wx + (ww / 2)
                centerY := wy + (wh / 2)
            } catch {
            }
        }
        if (centerX = 0 && centerY = 0) {
            workArea := SysGet.MonitorWorkArea(SysGet.MonitorPrimary)
            centerX := workArea.Left + (workArea.Right - workArea.Left) / 2
            centerY := workArea.Top + (workArea.Bottom - workArea.Top) / 2
        }
    }
    
    ; Create square GUI - just a colored square, no text
    dictationProgressGui := Gui()
    dictationProgressGui.Opt("+AlwaysOnTop -Caption +ToolWindow")
    dictationProgressGui.BackColor := "FF0000"  ; Red color
    
    ; Create a square by showing with explicit size
    ; Position centered on the captured point
    guiX := centerX - (size / 2)
    guiY := centerY - (size / 2)
    
    dictationProgressGui.Show("x" . Round(guiX) . " y" . Round(guiY) . " w" . size . " h" . size . " NA")
    WinSetTransparent(200, dictationProgressGui)
}

; Timer function that recreates the square with smaller size each second
UpdateDictationSquare() {
    global dictationProgressGui
    global dictationProgressStartTime
    global dictationProgressTimer
    global CONST_DICTATION_PROGRESS_INITIAL_SIZE
    global CONST_DICTATION_PROGRESS_MIN_SIZE
    global CONST_MAX_DICTATION_MS
    global g_dictationState
    
    ; Check if dictation is still active - if not, stop the timer and hide indicator
    if (g_dictationState != "ACTIVE" && g_dictationState != "RESTARTING") {
        ; Dictation stopped - clean up
        if (dictationProgressTimer) {
            SetTimer(dictationProgressTimer, 0)
            dictationProgressTimer := ""
        }
        HideDictationIndicator()
        return
    }
    
    ; Calculate elapsed time
    elapsedMs := A_TickCount - dictationProgressStartTime
    progress := elapsedMs / CONST_MAX_DICTATION_MS  ; 0.0 to 1.0
    
    ; Calculate new size: starts at INITIAL_SIZE, shrinks to MIN_SIZE (1px = 1 second remaining)
    if (progress >= 1.0) {
        ; Time's up - use minimum size (1px)
        newSize := CONST_DICTATION_PROGRESS_MIN_SIZE
        ; Stop timer when we reach the end
        if (dictationProgressTimer) {
            SetTimer(dictationProgressTimer, 0)
            dictationProgressTimer := ""
        }
    } else {
        ; Calculate size: initialSize - (progress * (initialSize - minSize))
        ; When progress = 29/30, size should be close to 1px (1 second remaining)
        newSize := Round(CONST_DICTATION_PROGRESS_INITIAL_SIZE - (progress * (CONST_DICTATION_PROGRESS_INITIAL_SIZE - CONST_DICTATION_PROGRESS_MIN_SIZE)))
        ; Ensure we don't go below minimum
        if (newSize < CONST_DICTATION_PROGRESS_MIN_SIZE) {
            newSize := CONST_DICTATION_PROGRESS_MIN_SIZE
        }
    }
    
    ; Recreate the square with new size
    try {
        CreateDictationSquare(newSize)
    } catch {
        ; Silently ignore errors
    }
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
        Sleep(100)
        Send("^a")
        Sleep(100)
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
; Auto-restart chime (distinct beep when 30-second timer expires and restart begins)
; =============================================================================
PlayAutoRestartChime() {
    try {
        static lastTick := 0
        if (A_TickCount - lastTick < 1000)
            return
        lastTick := A_TickCount

        played := false
        ; Use question icon beep (0x00000020) to distinguish from other cues
        ; This is a distinct sound that indicates a transition/restart event
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
