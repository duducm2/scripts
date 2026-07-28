; =============================================================================
; Gemini module: gemini_async_tts.ahk
; GeminiAsyncTTS class and GeminiQueueBackgroundTask
; Extracted verbatim from Gemini.ahk; loaded via #include into the
; Gemini.ahk process, which remains the entry point / source of truth.
; =============================================================================

; =============================================================================
; GeminiAsyncTTS – copy selection, send "repeat exactly" to Gemini, then trigger read aloud
; (formerly Win+Alt+Shift+7; chord freed — call class directly if rebinding)
; =============================================================================
class GeminiAsyncTTS {
    static TTSPrompt :=
        "Repeat the following text exactly as it is. Do not add any introduction, explanation, or markdown formatting. Just output the text itself:`n`n"
    static PostStreamingDelayMs := 600

    __New() {
        this.OriginalHwnd := 0
        this.GeminiHwnd := 0
        this.TimerCallback := ""
        this.RetryCount := 0
        this.MaxRetries := 60   ; 60 * 500ms = 30s timeout
        this.ButtonEverFound := false
        this.CopyCountAtSubmit := 0
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
        this.GeminiHwnd := GetGeminiWindowHwnd()
        if !this.GeminiHwnd {
            StandardLoadingBar_Hide(0)
            return
        }
        try {
            WinActivate("ahk_id " this.GeminiHwnd)
        } catch {
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        if !WinWaitActive("ahk_exe chrome.exe", , 2) {
            StandardLoadingBar_Hide(0)
            return
        }
        ; For TTS from selection (#!+7), always use the trash tab (second Gemini tab) when sending the prompt.
        ; Chrome convention: Ctrl+2 selects the second tab in the window.
        Send("^2")
        Sleep 150
        ShowGeminiTabBanner(2, this.GeminiHwnd)
        uia := UIA_Browser()
        Sleep 300
        promptField := Gemini_FocusPromptSameAsOpenHotkey(uia)
        if (!promptField) {
            StandardLoadingBar_Hide(0)
            return
        }
        ; Paste prompt + selected text and submit
        A_Clipboard := GeminiAsyncTTS.TTSPrompt . A_Clipboard
        Sleep 100
        Send("^a")
        Sleep 500
        Send("^v")
        Sleep 500
        ; Record Copy button count before submit (used when multiple-message validation is needed).
        Send("^{End}")
        Sleep 350
        this.CopyCountAtSubmit := GetGeminiCopyButtonCount(uia)
        Send("{Enter}")
        Sleep 300
        ; Return focus to original window (activate only; do not WinRestore or we lose maximized state)
        origHwnd := this.OriginalHwnd
        try {
            if WinExist("ahk_id " origHwnd) {
                WinActivate("ahk_id " origHwnd)
                WinWaitActive("ahk_id " origHwnd, , 1)
            }
        } catch {
            if (WinExist("ahk_id " origHwnd))
                WinActivate("ahk_id " origHwnd)
        }
        this.RetryCount := 0
        GeminiBackgroundSetTimer(this, this.CheckCompletion.Bind(this), GEMINI_ASYNC_POLL_MS)
    }

    CheckCompletion() {
        state := GeminiMonitorStreamingTransition(this, this.OnStreamingCompleted.Bind(this))
        if (state = "timeout") {
            StandardLoadingBar_Hide(0)
            return
        }
    }

    OnStreamingCompleted() {
        StandardLoadingBar_Hide(0)
        ; Completion detection matches GeminiAsyncLookup (#!+8): Layer 1 only (Stop button gone). No extra Layer 2 so we don't miss completion.
        try {
            ScriptSoundPlay(A_ScriptDir . "\assets\sounds\gemini-completion.wav")
        } catch {
            PlayCopyCompletedChime()
        }
        ; Allow DOM to finish rendering, then hand off read aloud without keeping focus on Gemini.
        Sleep(GeminiAsyncTTS.PostStreamingDelayMs)
        GeminiTriggerReadAloud(false, true, { originalHwnd: this.OriginalHwnd, geminiHwnd: this.GeminiHwnd })
    }
}

; Optional Python sidecar contract. AHK remains the default orchestration layer; if enabled later,
; the helper should accept a lightweight task envelope such as:
; { kind, geminiHwnd, originalHwnd, copyFirst, useTrashTab, requestedAt }
GeminiQueueBackgroundTask(taskKind, payload := "") {
    if (!GEMINI_USE_PYTHON_IPC)
        return false
    if (!GeminiIPC_EnsureReady(400))
        return false
    reqPayload := payload ? payload : Map()
    if (!(reqPayload is Map))
        reqPayload := Map()
    reqPayload["requestedAt"] := A_TickCount
    resp := GeminiIPC_QueueTask(taskKind, reqPayload)
    if (!GeminiIPC_ResponseOk(resp))
        return false
    return GeminiIPC_ResponseResultMap(resp)
}
