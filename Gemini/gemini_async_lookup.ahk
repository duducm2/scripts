; =============================================================================
; Gemini module: gemini_async_lookup.ahk
; GeminiAsyncLookup pronunciation async class
; Extracted verbatim from Gemini.ahk; loaded via #include into the
; Gemini.ahk process, which remains the entry point / source of truth.
; =============================================================================

; =============================================================================
; GeminiAsyncLookup – async pronunciation lookup (Win+Alt+Shift+8)
; User keeps focus; timer polls for completion; result shown in banner.
; =============================================================================
class GeminiAsyncLookup {
    __New(lang := "", preCopiedText := "") {
        this.Lang := lang
        this.PreCopiedText := preCopiedText
        this.OriginalHwnd := 0
        this.GeminiHwnd := 0
        this.TimerCallback := ""
        this.RetryCount := 0
        this.MaxRetries := 60   ; 60 * 500ms = 30s timeout
        this.ButtonEverFound := false
    }

    Start() {
        this.OriginalHwnd := WinExist("A")
        if !this.OriginalHwnd {
            return
        }
        ; Show loading banner immediately, centered on the monitor where this window is (with warning)
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
            try StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ Copy failed (no clipboard text). QuickLook selection may not support Ctrl+C.",
                2400,
                BANNER_ACCENT_ERROR)
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
        ; For pronunciation lookup (#!+8), always use the trash tab (second Gemini tab).
        ; Chrome convention: Ctrl+2 selects the second tab in the window.
        Send("^2")
        Sleep 150
        ShowGeminiTabBanner(2, this.GeminiHwnd)
        StandardLoadingBar_Update("🔄 Switching to 3.1 Flash-Lite…", BANNER_ACCENT_INTERMEDIATE)
        if (!GeminiSetModelForActiveTabWhenReady("3.1 Flash-Lite", this.GeminiHwnd))
            GeminiSetModelForActiveTabWhenReady("3.1 Flash-Lite", this.GeminiHwnd)
        StandardLoadingBar_Update("⏳ Loading…", BANNER_ACCENT_INTERMEDIATE)
        uia := UIA_Browser()
        Sleep 300
        promptField := Gemini_FocusPromptSameAsOpenHotkey(uia)
        if (!promptField) {
            StandardLoadingBar_Hide(0)
            return
        }
        try {
            promptField.SetFocus()
            Sleep 80
        } catch {
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
        ; Go back to the window where you triggered the hotkey so you can keep working (activate only; do not WinRestore or we lose maximized state)
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
        ; Use the same sound as Shift keys.ahk for consistency
        try {
            ScriptSoundPlay(A_ScriptDir . "\sounds\gemini-completion.wav")
        } catch {
            PlayCopyCompletedChime()
        }
        this.RetrieveResponse()
    }

    RetrieveResponse() {
        ; Activate Gemini once, then copy with retry (exponential backoff) without switching back until done.
        ; Invalidate last-Copy-button cache so we discover the newly completed response (avoid penultimate message).
        GeminiState.Invalidate()
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
        copyOpt := { restoreWindow: false, playChimeAndNotify: false, alreadyActive: true }
        seqBefore := GetClipboardSequenceNumber()
        if !CopyLastGeminiMessageWithRetry(copyOpt, this.GeminiHwnd) {
            StandardLoadingBar_Hide(0)
            return
        }
        ; Wait for clipboard change via sequence number (O(1) per check); proceed as soon as it changes instead of fixed delay.
        syncElapsed := 0
        while (syncElapsed < GEMINI_POST_COPY_SYNC_TIMEOUT_MS) {
            if (GetClipboardSequenceNumber() != seqBefore)
                break
            Sleep GEMINI_CLIPBOARD_POLL_MS
            syncElapsed += GEMINI_CLIPBOARD_POLL_MS
        }
        ; Use clipboard content only after change detected (or timeout) so the banner shows current content.
        bannerText := A_Clipboard
        if (this.OriginalHwnd = this.GeminiHwnd)
            FocusGeminiAskFieldForHwnd(this.GeminiHwnd, false)
        ; Original may have been closed while we waited on Gemini — never WinActivate a stale HWND.
        origHwnd := this.OriginalHwnd
        if (origHwnd && WinExist("ahk_id " origHwnd)) {
            try {
                WinActivate("ahk_id " origHwnd)
                if (!WinActive("ahk_id " origHwnd))
                    WinWaitActive("ahk_id " origHwnd, , 0.5)
            } catch {
                if (WinExist("ahk_id " origHwnd))
                    try WinActivate("ahk_id " origHwnd)
            }
        }
        StandardLoadingBar_Hide(0)
        this.ShowResultBanner(bannerText)
    }

    ShowResultBanner(text) {
        if (!text || StrLen(Trim(text)) = 0)
            return
        state := "ℹ " . text
        closeNoOp(*) {
        }
        closeKeys := Map("Enter", closeNoOp, "Escape", closeNoOp, "E", closeNoOp)
        ; timeoutMs 0 = no auto-dismiss; user closes with Enter, E, or Escape (Utils.ahk StandardLoadingBar_ShowWithKeys).
        StandardLoadingBar_ShowWithKeys(state, closeKeys, 0, 0, "",
            BANNER_ACCENT_INTERMEDIATE, 600,
            17, "", false,
            "[Enter] [E] [Esc] Close",
            true)
    }
}
