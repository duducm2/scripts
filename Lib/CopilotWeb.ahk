; M365 Copilot web automation (Chrome). Included from Gemini.ahk after Utils.ahk.
; Requires: env.ahk, Utils.ahk, UIA_Browser.

COPILOT_WEB_ACTIVATE_WAIT_MS := 2000
COPILOT_WEB_UIA_SETTLE_MS := 120
COPILOT_WEB_SCROLL_SETTLE_MS := 350
COPILOT_WEB_TAB_SETTLE_MS := 150
COPILOT_WEB_ASYNC_POLL_MS := 500
COPILOT_WEB_ASYNC_MAX_RETRIES := 60
COPILOT_WEB_COPY_MAX_RETRIES := 3
COPILOT_WEB_COPY_RETRY_SLEEP_MS := 400
COPILOT_WEB_STREAM_GONE_VERIFY_MS := 200
COPILOT_WEB_STREAM_GONE_LOOPS := 4
COPILOT_WEB_POST_COPY_SYNC_TIMEOUT_MS := 2000
COPILOT_WEB_CLIPBOARD_POLL_MS := 10
COPILOT_WEB_FIRST_LAUNCH_WAIT_MS := 2500

COPILOT_COPY_RESPONSE_NAMES := ["Copy", "Copiar"]
COPILOT_READ_ALOUD_NAMES := ["Read aloud", "Read Aloud", "Ler em voz alta"]
COPILOT_MORE_OPTIONS_NAMES := ["More options", "Show more options", "Mais opções"]
COPILOT_TTS_PAUSE_NAMES := ["Pause", "Pausar"]
COPILOT_TTS_RESUME_NAMES := ["Resume", "Retomar"]

UIA_Copilot_ControlType_Button := 50000
UIA_Copilot_ControlType_MenuItem := 50011

CopilotWeb_Notify(message, durationMs := 800, fontSize := 22) {
    StandardLoadingBar_Show(message, BANNER_ACCENT_INTERMEDIATE, { passive: true, fontSize: fontSize })
    StandardLoadingBar_Hide(durationMs)
}

CopilotWeb_IsCopilotWindow(hwnd) {
    global COPILOT_WEB_TITLE_NEEDLE
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    try {
        if (StrLower(WinGetProcessName("ahk_id " hwnd)) != "chrome.exe")
            return false
        return InStr(WinGetTitle("ahk_id " hwnd), COPILOT_WEB_TITLE_NEEDLE, false) > 0
    } catch {
        return false
    }
}

GetCopilotWebWindowHwnd() {
    hwnd := WinExist("A")
    if (CopilotWeb_IsCopilotWindow(hwnd))
        return hwnd
    try {
        for h in WinGetList("ahk_exe chrome.exe") {
            if (CopilotWeb_IsCopilotWindow(h))
                return h
        }
    } catch {
    }
    return 0
}

CopilotWeb_FindFirstInUia(uia, criteriaList) {
    if (!IsObject(uia))
        return 0
    for criteria in criteriaList {
        try {
            el := uia.FindFirst(criteria)
            if (el)
                return el
        } catch {
        }
    }
    return 0
}

CopilotWeb_FindComposer(uia) {
    return CopilotWeb_FindFirstInUia(uia, [
        { AutomationId: "m365-chat-editor-target-element", ControlType: "Edit" },
        { AutomationId: "m365-chat-editor-target-element" },
        { Name: "Message Copilot", ControlType: "Edit" }
    ])
}

CopilotWeb_FindStopGenerating(uia) {
    if (!IsObject(uia))
        return 0
    el := CopilotWeb_FindFirstInUia(uia, [
        { Name: "Stop generating", ControlType: "Button" },
        { Name: "Stop generating", Type: UIA_Copilot_ControlType_Button }
    ])
    if (el)
        return el
    try {
        return uia.FindFirst({ Name: "Stop generating", matchmode: "Substring", ControlType: "Button" })
    } catch {
    }
    return 0
}

CopilotWeb_FindSendButton(uia) {
    return CopilotWeb_FindFirstInUia(uia, [
        { Name: "Send ", matchmode: "Substring", ControlType: "Button" },
        { ClassName: "fai-SendButton", matchmode: "Substring", ControlType: "Button" }
    ])
}

CopilotWeb_TrySubmit(uia) {
    if (!IsObject(uia))
        return false
    if (CopilotWeb_FindStopGenerating(uia))
        return false
    sendBtn := CopilotWeb_FindSendButton(uia)
    if (sendBtn) {
        try {
            if (sendBtn.GetPropertyValue(UIA.Property.IsInvokePatternAvailable)) {
                sendBtn.InvokePattern.Invoke()
                return true
            }
        } catch {
        }
        try {
            sendBtn.Click()
            return true
        } catch {
        }
    }
    SendInput "{Enter}"
    return true
}

CopilotWeb_PlayFocusedChime(minIntervalMs := 400) {
    static lastChimeTick := 0
    if (!IsSoundEnabled())
        return false
    now := A_TickCount
    if (lastChimeTick && (now - lastChimeTick) < minIntervalMs)
        return false
    lastChimeTick := now
    ScriptSoundPlay(A_ScriptDir . "\sounds\gemini-focused.wav")
    return true
}

CopilotWeb_FocusComposer(uia, playChime := true) {
    el := CopilotWeb_FindComposer(uia)
    if (!el)
        return 0
    try {
        if (el.HasKeyboardFocus) {
            if (playChime)
                CopilotWeb_PlayFocusedChime()
            return el
        }
    } catch {
    }
    try el.ScrollIntoView()
    try el.SetFocus()
    Sleep 40
    try el.Click()
    Sleep 40
    try {
        if (el.HasKeyboardFocus && playChime)
            CopilotWeb_PlayFocusedChime()
    } catch {
    }
    return el
}

CopilotWeb_FocusComposerForHwnd(copilotHwnd, playChime := false) {
    if (!copilotHwnd || !WinActive("ahk_id " copilotHwnd))
        return false
    try {
        uia := UIA_Browser("ahk_id " copilotHwnd)
        return !!CopilotWeb_FocusComposer(uia, playChime)
    } catch {
    }
    return false
}

CopilotWeb_WaitForComposerDiscoverable(uia, timeoutMs := 500) {
    start := A_TickCount
    while (A_TickCount - start < timeoutMs) {
        if (CopilotWeb_FindComposer(uia))
            return true
        Sleep 25
    }
    return !!CopilotWeb_FindComposer(uia)
}

CopilotWeb_ActivateWindow(hwnd, timeoutMs := COPILOT_WEB_ACTIVATE_WAIT_MS) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    try {
        WinActivate("ahk_id " hwnd)
    } catch {
        return false
    }
    return WinWaitActive("ahk_id " hwnd, , timeoutMs // 1000)
}

CopilotWeb_OpenOrFocus() {
    global COPILOT_WEB_URL
    SetTitleMatchMode(2)
    hwnd := GetCopilotWebWindowHwnd()
    if (hwnd) {
        alreadyActive := false
        try alreadyActive := WinActive("ahk_id " hwnd)
        if (!alreadyActive) {
            if !CopilotWeb_ActivateWindow(hwnd)
                return false
        }
        try uia := UIA_Browser("ahk_id " hwnd)
        catch {
            ShowCenteredOverlay_Utils("❌ Error: Could not attach to Copilot window.", 2000, BANNER_ACCENT_ERROR)
            return false
        }
        if (!alreadyActive)
            CopilotWeb_WaitForComposerDiscoverable(uia)
        CopilotWeb_FocusComposer(uia, true)
        return true
    }
    try {
        StandardLoadingBar_Show("📤 Opening Copilot...", BANNER_ACCENT_INTERMEDIATE)
        Run 'chrome.exe --new-window "' COPILOT_WEB_URL '"'
        if !WinWaitActive("ahk_exe chrome.exe", , 8) {
            StandardLoadingBar_Hide(0)
            return false
        }
        Sleep COPILOT_WEB_FIRST_LAUNCH_WAIT_MS
        hwnd := GetCopilotWebWindowHwnd()
        if (!hwnd)
            hwnd := WinExist("A")
        if (!hwnd) {
            StandardLoadingBar_Hide(0)
            return false
        }
        try uia := UIA_Browser("ahk_id " hwnd)
        catch {
            StandardLoadingBar_Hide(0)
            return false
        }
        CopilotWeb_WaitForComposerDiscoverable(uia, 4000)
        CopilotWeb_FocusComposer(uia, true)
        StandardLoadingBar_Hide(0)
        return true
    } catch {
        StandardLoadingBar_Hide(0)
        return false
    }
}

CopilotWeb_ComposerGetText(copilotHwnd := 0) {
    try {
        uia := copilotHwnd ? UIA_Browser("ahk_id " copilotHwnd) : UIA_Browser()
        pf := CopilotWeb_FindComposer(uia)
        if (!pf)
            return ""
        try {
            text := Trim(pf.Value)
            if (text != "")
                return text
        } catch {
        }
        try {
            text := Trim(pf.TextPattern.DocumentRange.GetText(-1))
            if (text != "")
                return text
        } catch {
        }
    } catch {
    }
    return ""
}

; Navigate to Copilot, focus composer, paste snippet (optional text or Clip Angel first snippet).
CopilotWeb_NavigateFocusAndPaste(optionalPromptText := "", autoSubmit := false) {
    global COPILOT_WEB_URL
    SetTitleMatchMode(2)
    copilotHwnd := GetCopilotWebWindowHwnd()
    if (!copilotHwnd) {
        Run 'chrome.exe --new-window "' COPILOT_WEB_URL '"'
        if !WinWaitActive("ahk_exe chrome.exe", , 5)
            return 0
        Sleep COPILOT_WEB_FIRST_LAUNCH_WAIT_MS
        copilotHwnd := GetCopilotWebWindowHwnd()
        if (!copilotHwnd)
            copilotHwnd := WinExist("A")
    }
    if (!copilotHwnd)
        return 0
    WinActivate("ahk_id " copilotHwnd)
    WinWaitActive("ahk_id " copilotHwnd, , 2)
    Sleep COPILOT_WEB_TAB_SETTLE_MS
    try {
        uia := UIA_Browser("ahk_id " copilotHwnd)
        Sleep 80
        CopilotWeb_FocusComposer(uia, false)
    } catch {
        return 0
    }
    WinActivate("ahk_id " copilotHwnd)
    WinWaitActive("ahk_id " copilotHwnd, , 2)
    Sleep 150
    if (optionalPromptText != "")
        InsertText(optionalPromptText)
    else {
        Send "!v"
        Sleep 50
        Send "^!b"
    }
    Sleep 250
    CopilotWeb_PlayFocusedChime()
    if (autoSubmit) {
        Sleep 1000
        try {
            uia := UIA_Browser("ahk_id " copilotHwnd)
            CopilotWeb_TrySubmit(uia)
        } catch {
            Send "{Enter}"
        }
    }
    return copilotHwnd
}

CopilotWeb_IsCopyResponseButton(name) {
    if (!name)
        return false
    for n in COPILOT_COPY_RESPONSE_NAMES {
        if (name = n)
            return true
    }
    if (InStr(name, "Copy prompt", false) || InStr(name, "Copiar prompt", false))
        return false
    return false
}

CopilotWeb_GetCopyButtonsArray(uia) {
    out := []
    if (!IsObject(uia))
        return out
    try {
        allButtons := uia.FindAll({ Type: "Button" })
        for button in allButtons {
            if (CopilotWeb_IsCopyResponseButton(button.Name))
                out.Push(button)
        }
    } catch {
    }
    return out
}

CopilotWeb_GetLastCopyButton(uia) {
    arr := CopilotWeb_GetCopyButtonsArray(uia)
    if (arr.Length = 0)
        return 0
    lastEl := 0
    lastTop := ""
    for btn in arr {
        try {
            br := btn.BoundingRectangle
        } catch {
            continue
        }
        if (!IsObject(br))
            continue
        if ((br.r - br.l) <= 0 || (br.b - br.t) <= 0)
            continue
        if (lastEl = 0 || br.t >= lastTop) {
            lastEl := btn
            lastTop := br.t
        }
    }
    return lastEl ? lastEl : arr[arr.Length]
}

CopilotWeb_GetLowestButtonByNames(uia, names) {
    if (!IsObject(uia))
        return 0
    candidates := []
    try {
        allButtons := uia.FindAll({ Type: "Button" })
        for button in allButtons {
            for n in names {
                if (button.Name = n || InStr(button.Name, n, false) = 1) {
                    candidates.Push(button)
                    break
                }
            }
        }
    } catch {
    }
    if (candidates.Length = 0)
        return 0
    lastEl := 0
    lastTop := ""
    for btn in candidates {
        try {
            br := btn.BoundingRectangle
        } catch {
            continue
        }
        if (!IsObject(br))
            continue
        if (lastEl = 0 || br.t >= lastTop) {
            lastEl := btn
            lastTop := br.t
        }
    }
    return lastEl
}

CopilotWeb_FindPauseResumeButton(uia, which) {
    names := (which = "Resume") ? COPILOT_TTS_RESUME_NAMES : COPILOT_TTS_PAUSE_NAMES
    return CopilotWeb_GetLowestButtonByNames(uia, names)
}

CopilotWeb_CopyLastMessageToClipboard(options := "", copilotHwnd := 0) {
    restoreWindow := (options = "" || !options.HasProp("restoreWindow")) ? true : options.restoreWindow
    playChimeAndNotify := (options = "" || !options.HasProp("playChimeAndNotify")) ? true : options.playChimeAndNotify
    alreadyActive := (options != "" && options.HasProp("alreadyActive")) ? options.alreadyActive : false
    try {
        SetTitleMatchMode(2)
        if (!copilotHwnd)
            copilotHwnd := GetCopilotWebWindowHwnd()
        if (!copilotHwnd)
            return false
        if (!alreadyActive) {
            if !CopilotWeb_ActivateWindow(copilotHwnd)
                return false
            Sleep COPILOT_WEB_TAB_SETTLE_MS
        }
        uia := alreadyActive ? UIA_Browser() : UIA_Browser("ahk_id " copilotHwnd)
        Sleep COPILOT_WEB_UIA_SETTLE_MS
        Send "^{End}"
        Sleep COPILOT_WEB_SCROLL_SETTLE_MS
        copyBtn := CopilotWeb_GetLastCopyButton(uia)
        if (!copyBtn)
            return false
        A_Clipboard := ""
        copyBtn.Click()
        if !ClipWait(2)
            return false
        if (playChimeAndNotify) {
            try ScriptSoundPlay(A_ScriptDir . "\sounds\copy.wav")
            CopilotWeb_Notify("Copied!", 800, 24)
        }
        if (restoreWindow)
            Send "!{Tab}"
        else
            CopilotWeb_FocusComposerForHwnd(copilotHwnd, false)
        return true
    } catch {
        return false
    }
}

CopilotWeb_CopyLastMessageWithRetry(options := "", copilotHwnd := 0, maxRetries := COPILOT_WEB_COPY_MAX_RETRIES) {
    baseDelay := COPILOT_WEB_COPY_RETRY_SLEEP_MS
    loop maxRetries {
        if (CopilotWeb_CopyLastMessageToClipboard(options, copilotHwnd))
            return true
        if (A_Index < maxRetries)
            Sleep baseDelay * (1 << (A_Index - 1))
    }
    return false
}

CopilotBackgroundSetTimer(task, callback, periodMs := COPILOT_WEB_ASYNC_POLL_MS) {
    CopilotBackgroundStopTimer(task)
    task.TimerCallback := callback
    SetTimer(task.TimerCallback, periodMs)
}

CopilotBackgroundStopTimer(task) {
    cb := ""
    try cb := task.TimerCallback
    catch
        cb := ""
    if (cb)
        SetTimer(cb, 0)
    try task.TimerCallback := ""
}

CopilotWeb_VerifyStreamingStopped(copilotHwnd) {
    loop COPILOT_WEB_STREAM_GONE_LOOPS {
        Sleep COPILOT_WEB_STREAM_GONE_VERIFY_MS
        try {
            uia := UIA_Browser("ahk_id " copilotHwnd)
            if (!CopilotWeb_FindStopGenerating(uia))
                continue
            return false
        } catch {
            return true
        }
    }
    return true
}

CopilotWeb_MonitorStreamingTransition(task, onCompleteCallback) {
    task.RetryCount++
    if (task.RetryCount > task.MaxRetries) {
        CopilotBackgroundStopTimer(task)
        return "timeout"
    }
    if (!task.CopilotHwnd || !WinExist("ahk_id " task.CopilotHwnd)) {
        CopilotBackgroundStopTimer(task)
        return "unavailable"
    }
    try {
        uia := UIA_Browser("ahk_id " task.CopilotHwnd)
    } catch {
        return "unavailable"
    }
    if (CopilotWeb_FindStopGenerating(uia)) {
        task.ButtonEverFound := true
        return "streaming"
    }
    if (!task.ButtonEverFound)
        return "waiting"
    if (!CopilotWeb_VerifyStreamingStopped(task.CopilotHwnd))
        return "streaming"
    CopilotBackgroundStopTimer(task)
    onCompleteCallback.Call()
    return "completed"
}

CopilotWeb_TriggerReadAloud(copyFirst := true, options := "") {
    return (CopilotAsyncReadAloud(copyFirst, options)).Start()
}

; --- Async read aloud (#!+o, D2C R) ---
class CopilotAsyncReadAloud {
    __New(copyFirst := true, options := "") {
        this.CopyFirst := copyFirst
        this.OriginalHwnd := (options != "" && options.HasProp("originalHwnd")) ? options.originalHwnd : 0
        this.CopilotHwnd := (options != "" && options.HasProp("copilotHwnd")) ? options.copilotHwnd : 0
        this.AlreadyActive := (options != "" && options.HasProp("alreadyActive")) ? options.alreadyActive : false
        this.Uia := 0
        this.CopyRetryCount := 0
    }

    Start() {
        if (!this.OriginalHwnd)
            try this.OriginalHwnd := WinExist("A")
        SetTitleMatchMode(2)
        if (!this.CopilotHwnd)
            this.CopilotHwnd := GetCopilotWebWindowHwnd()
        if (!this.CopilotHwnd) {
            CopilotWeb_Notify("Read aloud failed – Copilot is not open", 1800, 22)
            return false
        }
        SetTimer(this.Run.Bind(this), -1)
        return true
    }

    Run(*) {
        try {
            if (!WinActive("ahk_id " this.CopilotHwnd)) {
                StandardLoadingBar_Show("⏳ Switching to Copilot…", BANNER_ACCENT_INTERMEDIATE)
                if !CopilotWeb_ActivateWindow(this.CopilotHwnd) {
                    StandardLoadingBar_Hide(0)
                    this.Fail()
                    return
                }
            }
            try this.Uia := UIA_Browser("ahk_id " this.CopilotHwnd)
            catch {
                StandardLoadingBar_Hide(0)
                this.Fail()
                return
            }
            Sleep COPILOT_WEB_UIA_SETTLE_MS
            StandardLoadingBar_Hide(0)

            pauseBtn := CopilotWeb_FindPauseResumeButton(this.Uia, "Pause")
            if (pauseBtn) {
                try pauseBtn.Click()
                CopilotWeb_Notify("Paused", 800, 24)
                this.RestoreFocus()
                return
            }
            resumeBtn := CopilotWeb_FindPauseResumeButton(this.Uia, "Resume")
            if (resumeBtn) {
                try resumeBtn.Click()
                CopilotWeb_Notify("Resumed", 800, 24)
                this.RestoreFocus()
                return
            }

            Send "^{End}"
            Sleep COPILOT_WEB_SCROLL_SETTLE_MS
            if (this.CopyFirst) {
                copyOpt := { restoreWindow: false, playChimeAndNotify: false, alreadyActive: true }
                if (!CopilotWeb_CopyLastMessageToClipboard(copyOpt, this.CopilotHwnd))
                    this.CopyRetryCount := COPILOT_WEB_COPY_MAX_RETRIES
            }

            readBtn := CopilotWeb_GetLowestButtonByNames(this.Uia, COPILOT_READ_ALOUD_NAMES)
            if (!readBtn) {
                moreBtn := CopilotWeb_GetLowestButtonByNames(this.Uia, COPILOT_MORE_OPTIONS_NAMES)
                if (moreBtn) {
                    try moreBtn.Click()
                    Sleep 200
                    try this.Uia := UIA_Browser("ahk_id " this.CopilotHwnd)
                    for n in COPILOT_READ_ALOUD_NAMES {
                        try {
                            item := this.Uia.FindFirst({ Name: n, Type: UIA_Copilot_ControlType_MenuItem })
                            if (item) {
                                item.Click()
                                readBtn := item
                                break
                            }
                        } catch {
                        }
                    }
                }
            } else {
                try readBtn.Click()
            }

            if (!readBtn) {
                ShowCenteredOverlay_Utils("Read aloud not available for Copilot web", 2500, BANNER_ACCENT_ERROR)
                this.RestoreFocus()
                return
            }

            CopilotWeb_Notify(this.CopyFirst ? "Copied & Reading aloud" : "Reading aloud", 800, 24)
            this.RestoreFocus()
        } catch {
            this.Fail()
        }
    }

    RestoreFocus() {
        orig := this.OriginalHwnd
        if (orig && orig != this.CopilotHwnd && WinExist("ahk_id " orig)) {
            try {
                WinActivate("ahk_id " orig)
                WinWaitActive("ahk_id " orig, , 0.5)
            } catch {
            }
        } else if (this.CopilotHwnd && WinActive("ahk_id " this.CopilotHwnd))
            CopilotWeb_FocusComposerForHwnd(this.CopilotHwnd, false)
    }

    Fail() {
        StandardLoadingBar_Hide(0)
        CopilotWeb_Notify("Read aloud failed – Copilot UI not ready", 2000, 22)
        this.RestoreFocus()
    }
}

; --- TTS from selection (#!+7) ---
class CopilotAsyncTTS {
    static TTSPrompt :=
        "Repeat the following text exactly as it is. Do not add any introduction, explanation, or markdown formatting. Just output the text itself:`n`n"
    static PostStreamingDelayMs := 600

    __New() {
        this.OriginalHwnd := 0
        this.CopilotHwnd := 0
        this.RetryCount := 0
        this.MaxRetries := COPILOT_WEB_ASYNC_MAX_RETRIES
        this.ButtonEverFound := false
    }

    Start() {
        this.OriginalHwnd := WinExist("A")
        if !this.OriginalHwnd
            return
        StandardLoadingBar_Show("⏳ Loading…", BANNER_ACCENT_INTERMEDIATE)
        A_Clipboard := ""
        Send "^c"
        if !ClipWait(2) {
            StandardLoadingBar_Hide(0)
            return
        }
        SetTitleMatchMode(2)
        this.CopilotHwnd := GetCopilotWebWindowHwnd()
        if !this.CopilotHwnd {
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ Copilot is not open", 2000, BANNER_ACCENT_ERROR)
            return
        }
        if !CopilotWeb_ActivateWindow(this.CopilotHwnd) {
            StandardLoadingBar_Hide(0)
            return
        }
        try uia := UIA_Browser("ahk_id " this.CopilotHwnd)
        catch {
            StandardLoadingBar_Hide(0)
            return
        }
        Sleep 300
        if (!CopilotWeb_FocusComposer(uia, false)) {
            StandardLoadingBar_Hide(0)
            return
        }
        A_Clipboard := CopilotAsyncTTS.TTSPrompt . A_Clipboard
        Sleep 100
        Send("^a")
        Sleep 500
        Send("^v")
        Sleep 500
        Send("{Enter}")
        Sleep 300
        origHwnd := this.OriginalHwnd
        try {
            if WinExist("ahk_id " origHwnd) {
                WinActivate("ahk_id " origHwnd)
                WinWaitActive("ahk_id " origHwnd, , 1)
            }
        } catch {
        }
        this.RetryCount := 0
        CopilotBackgroundSetTimer(this, this.CheckCompletion.Bind(this), COPILOT_WEB_ASYNC_POLL_MS)
    }

    CheckCompletion() {
        state := CopilotWeb_MonitorStreamingTransition(this, this.OnStreamingCompleted.Bind(this))
        if (state = "timeout")
            StandardLoadingBar_Hide(0)
    }

    OnStreamingCompleted() {
        StandardLoadingBar_Hide(0)
        try ScriptSoundPlay(A_ScriptDir . "\sounds\gemini-completion.wav")
        catch {
        }
        Sleep(CopilotAsyncTTS.PostStreamingDelayMs)
        CopilotWeb_TriggerReadAloud(false, { originalHwnd: this.OriginalHwnd, copilotHwnd: this.CopilotHwnd })
    }
}

; --- Pronunciation lookup (#!+8) ---
class CopilotAsyncLookup {
    __New(lang, selectedText := "", preCopiedText := "") {
        this.Lang := lang
        this.PreCopiedText := preCopiedText
        this.OriginalHwnd := 0
        this.CopilotHwnd := 0
        this.RetryCount := 0
        this.MaxRetries := COPILOT_WEB_ASYNC_MAX_RETRIES
        this.ButtonEverFound := false
        if (selectedText != "")
            this.PreCopiedText := selectedText
    }

    Start() {
        this.OriginalHwnd := WinExist("A")
        if !this.OriginalHwnd
            return
        StandardLoadingBar_Show("⏳ Loading…", BANNER_ACCENT_INTERMEDIATE)
        clipOk := false
        if (this.PreCopiedText != "") {
            A_Clipboard := this.PreCopiedText
            clipOk := true
        } else {
            A_Clipboard := ""
            clipOk := TryCopySelectionToClipboard_QuickLookAware()
        }
        if !clipOk {
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ Copy failed (no clipboard text).", 2400, BANNER_ACCENT_ERROR)
            return
        }
        SetTitleMatchMode(2)
        this.CopilotHwnd := GetCopilotWebWindowHwnd()
        if !this.CopilotHwnd {
            StandardLoadingBar_Hide(0)
            return
        }
        if !CopilotWeb_ActivateWindow(this.CopilotHwnd) {
            StandardLoadingBar_Hide(0)
            return
        }
        try uia := UIA_Browser("ahk_id " this.CopilotHwnd)
        catch {
            StandardLoadingBar_Hide(0)
            return
        }
        Sleep 300
        if (!CopilotWeb_FocusComposer(uia, false)) {
            StandardLoadingBar_Hide(0)
            return
        }
        promptName := this.Lang ? "pronunciation-lookup-" . this.Lang : "pronunciation-lookup"
        lookupLangTitle := Map("pt", "Português brasileiro", "en", "English", "de", "Deutsch")
        lead := ""
        if (this.Lang != "" && lookupLangTitle.Has(this.Lang))
            lead :=
                "Answer preamble (mandatory): The very first line of your reply must be exactly this language label (so the reader sees which lookup mode was used):`n"
                . lookupLangTitle[this.Lang]
                .
                "`nThe second line must be blank. After that, follow every instruction below—including the seven sections—with no section titles or markdown headings for those sections (this preamble is the only allowed title line).`n`n"
        searchString := lead . RTrim(GetPromptText(promptName), "`r`n")
        A_Clipboard := searchString . "`n`nContent: " . A_Clipboard
        Sleep 100
        Send("^a")
        Sleep 500
        Send("^v")
        Sleep 500
        Send("{Enter}")
        Sleep 300
        origHwnd := this.OriginalHwnd
        try {
            if WinExist("ahk_id " origHwnd) {
                WinActivate("ahk_id " origHwnd)
                WinWaitActive("ahk_id " origHwnd, , 1)
            }
        } catch {
        }
        this.RetryCount := 0
        CopilotBackgroundSetTimer(this, this.CheckCompletion.Bind(this), COPILOT_WEB_ASYNC_POLL_MS)
    }

    CheckCompletion() {
        state := CopilotWeb_MonitorStreamingTransition(this, this.OnStreamingCompleted.Bind(this))
        if (state = "timeout")
            StandardLoadingBar_Hide(0)
    }

    OnStreamingCompleted() {
        try ScriptSoundPlay(A_ScriptDir . "\sounds\gemini-completion.wav")
        catch {
        }
        this.RetrieveResponse()
    }

    RetrieveResponse() {
        try WinActivate("ahk_id " this.CopilotHwnd)
        catch {
            StandardLoadingBar_Hide(0)
            return
        }
        if !WinWaitActive("ahk_exe chrome.exe", , 2) {
            StandardLoadingBar_Hide(0)
            return
        }
        copyOpt := { restoreWindow: false, playChimeAndNotify: false, alreadyActive: true }
        seqBefore := Clipboard_GetSequenceNumber()
        if !CopilotWeb_CopyLastMessageWithRetry(copyOpt, this.CopilotHwnd) {
            StandardLoadingBar_Hide(0)
            return
        }
        syncElapsed := 0
        while (syncElapsed < COPILOT_WEB_POST_COPY_SYNC_TIMEOUT_MS) {
            if (Clipboard_GetSequenceNumber() != seqBefore)
                break
            Sleep COPILOT_WEB_CLIPBOARD_POLL_MS
            syncElapsed += COPILOT_WEB_CLIPBOARD_POLL_MS
        }
        bannerText := A_Clipboard
        origHwnd := this.OriginalHwnd
        if (origHwnd && WinExist("ahk_id " origHwnd)) {
            try {
                WinActivate("ahk_id " origHwnd)
                if (!WinActive("ahk_id " origHwnd))
                    WinWaitActive("ahk_id " origHwnd, , 0.5)
            } catch {
            }
        }
        StandardLoadingBar_Hide(0)
        if (!bannerText || StrLen(Trim(bannerText)) = 0)
            return
        state := "ℹ " . bannerText
        closeNoOp(*) {
        }
        closeKeys := Map("Enter", closeNoOp, "Escape", closeNoOp, "E", closeNoOp)
        StandardLoadingBar_ShowWithKeys(state, closeKeys, 0, 0, "",
            BANNER_ACCENT_INTERMEDIATE, 600, 17, "", false,
            "[Enter] [E] [Esc] Close", true)
    }
}

CopilotHotkey_ShowPronunciationLanguagePicker(selectedText) {
    onSelect(lang) {
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        if (lang != "")
            (CopilotAsyncLookup(lang, selectedText)).Start()
    }
    onTimeout() {
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        StandardLoadingBar_Show("⏳ Detecting language…", BANNER_ACCENT_INTERMEDIATE, { textWidth: 450, fontSize: 17 })
        lang := DetectLang_AhkFallback(selectedText)
        if !(lang = "pt" || lang = "en" || lang = "de")
            lang := "en"
        (CopilotAsyncLookup(lang, selectedText)).Start()
    }
    keyCallbacks := Map(
        "1", (*) => onSelect("pt"),
        "2", (*) => onSelect("en"),
        "3", (*) => onSelect("de"),
        "*Escape", (*) => onSelect("")
    )
    StandardLoadingBar_ShowWithKeys("❓ Auto-detect in 2s — press to override", keyCallbacks, 2000, 0, onTimeout,
        BANNER_ACCENT_INTERMEDIATE,
        450, 17, "", false, "[1] Portuguese  [2] English  [3] German  [Esc] Cancel", false, true)
}
