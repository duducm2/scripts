; =============================================================================
; Gemini module: gemini_delayed_submit.ahk
; GeminiDelayedSubmitMonitor and start/stop helpers
; Extracted verbatim from Gemini.ahk; loaded via #include into the
; Gemini.ahk process, which remains the entry point / source of truth.
; =============================================================================

; =============================================================================
; GeminiDelayedSubmitMonitor – background completion monitor for Ctrl+Alt+Win+L
; Reuses #!+8 completion detection; on completion shows "Copy? [N] [R]" with 4s timeout (N = no copy, R = copy + read aloud).
; =============================================================================
class GeminiDelayedSubmitMonitor {
    __New() {
        this.OriginalHwnd := 0
        this.GeminiHwnd := 0
        this.TimerCallback := ""
        this.RetryCount := 0
        this.MaxRetries := 300   ; 300 * 500ms = 150s timeout (covers Gemini Pro deep thinking)
        this.ButtonEverFound := false
        this.CopyBannerGui := ""
        this.CopyTimeoutTimer := ""
        this.HasCopiedForThisResponse := false
        this.CopyBannerShownForThisResponse := false
    }

    Start(originalHwnd, geminiHwnd) {
        if (!originalHwnd || !geminiHwnd)
            return
        this.OriginalHwnd := originalHwnd
        this.GeminiHwnd := geminiHwnd
        this.RetryCount := 0
        this.ButtonEverFound := false
        this.HasCopiedForThisResponse := false
        this.CopyBannerShownForThisResponse := false
        GeminiBackgroundSetTimer(this, this.CheckCompletion.Bind(this), GEMINI_ASYNC_POLL_MS)
    }

    ; Stop polling; used when user chose S or N at 6s so "Copy response?" is never shown for this flow.
    Stop() {
        GeminiBackgroundStopTimer(this)
    }

    CheckCompletion() {
        state := GeminiMonitorStreamingTransition(this, this.OnStreamingCompleted.Bind(this))
        if (state = "timeout")
            return
    }

    OnStreamingCompleted() {
        try {
            ScriptSoundPlay(A_ScriptDir . "\assets\sounds\gemini-completion.wav")
        } catch {
            PlayCopyCompletedChime()
        }
        this.ShowCopyDecisionBanner()
    }

    ShowCopyDecisionBanner() {
        if (this.CopyBannerShownForThisResponse)
            return
        this.CopyBannerShownForThisResponse := true
        this.CopyBannerGui := ""
        this.CopyTimeoutTimer := ""
        copyKeyCallbacks := Map("N", this.CancelCopy.Bind(this), "Y", this.DoCopyOnly.Bind(this), "R", this.CopyAndReadAloud
        .Bind(this), "C", this.CopyAndTransferToCursor.Bind(this), "F", this.CopyAndFavorite.Bind(this))
        StandardLoadingBar_ShowWithKeys("❓ Copy response?", copyKeyCallbacks, 5000, 0, this.DoCopyOnTimeout
            .Bind(this), BANNER_ACCENT_INTERMEDIATE, 520, 17, "", false,
            "[Y] Copy  [N] No  [R] Copy+Read  [C] Transfer  [F] Copy+Favorite",
            true)
    }

    ; Y key: copy latest response only (same as timeout default), then close banner and restore focus.
    DoCopyOnly(*) {
        this.DoCopyOnTimeout()
    }

    ; Shared cleanup: clear Gemini-side state only. Key/timer unregister and overlay hide are handled by Utils.
    CleanupCopyBanner() {
        this.CopyBannerGui := ""
        this.CopyTimeoutTimer := ""
    }

    ; N key: close overlay and stop (no copy/read/transfer). Explicitly close Utils overlay so cancel always takes effect.
    CancelCopy(*) {
        try StandardLoadingBar_CloseKeysOverlay()
        catch {
        }
        try StandardLoadingBar_Hide(0)
        catch {
        }
        this.CleanupCopyBanner()
    }

    DoCopyCore(readAloud := false, skipRestoreFocus := false) {
        if (this.HasCopiedForThisResponse)
            return
        this.HasCopiedForThisResponse := true
        GeminiState.Invalidate()
        if !WinExist("ahk_id " this.GeminiHwnd) {
            if (WinExist("ahk_id " this.OriginalHwnd))
                WinActivate("ahk_id " this.OriginalHwnd)
            return
        }
        ; Hands off cue before activating Gemini to copy the last response (manual Y/R/C and timeout).
        PlayPreMovementWarning("Gemini")
        ; If Gemini is not active when the monitor fires, activate it now.
        if !WinActive("ahk_id " this.GeminiHwnd) {
            try {
                WinActivate("ahk_id " this.GeminiHwnd)
            } catch {
                if (WinExist("ahk_id " this.OriginalHwnd))
                    WinActivate("ahk_id " this.OriginalHwnd)
                return
            }
            if !WinWaitActive("ahk_exe chrome.exe", , 0.5) {
                if (WinExist("ahk_id " this.OriginalHwnd))
                    WinActivate("ahk_id " this.OriginalHwnd)
                return
            }
        }
        copyOpt := { restoreWindow: false, playChimeAndNotify: false, alreadyActive: true }
        if CopyLastGeminiMessageWithRetry(copyOpt, this.GeminiHwnd) {
            PlayCopyCompletedChime()
        }
        if (readAloud) {
            GeminiTriggerReadAloud(false, false, { originalHwnd: this.OriginalHwnd, geminiHwnd: this.GeminiHwnd,
                alreadyActive: true, verifyMaxRetries: GEMINI_DICTATION_READ_ALOUD_MAX_RETRIES })
            ; Do not WinActivate(original) here: async read-aloud must keep Gemini foreground until Pause is found.
        } else if (WinActive("ahk_id " this.GeminiHwnd))
            FocusGeminiAskFieldForHwnd(this.GeminiHwnd, false)
        ; Gemini/Clipboard → Original: return transitions are immediate (no warning).
        if (!skipRestoreFocus && !readAloud && WinExist("ahk_id " this.OriginalHwnd) && !WinActive("ahk_id " this.OriginalHwnd
        )) {
            WinActivate("ahk_id " this.OriginalHwnd)
            ; Fast-path: avoid WinWaitActive if we are already active.
            if (!WinActive("ahk_id " this.OriginalHwnd))
                WinWaitActive("ahk_id " this.OriginalHwnd, , 0.5)
        }
    }

    DoCopyOnTimeout(*) {
        this.CleanupCopyBanner()
        this.DoCopyCore(false)
    }

    ; R key: copy last message and read it aloud, then restore focus (same tab as delayed submit).
    CopyAndReadAloud(*) {
        this.CleanupCopyBanner()
        this.DoCopyCore(true)
    }

    ; F key: copy last response, then mark the new top clip as favorite in Clip Angel (same as MarkLastClipAsFavorite).
    CopyAndFavorite(*) {
        this.CleanupCopyBanner()
        this.DoCopyCore(false)
        clipRaw := A_Clipboard
        clip := Trim(clipRaw)
        if (clip = "" || StrLen(clip) < GEMINI_TRANSFER_MIN_CLIPBOARD_LENGTH) {
            ShowCenteredOverlay_Utils("❌ Copy failed or empty – try again", 2000, BANNER_ACCENT_ERROR)
            if (WinExist("ahk_id " this.OriginalHwnd))
                WinActivate("ahk_id " this.OriginalHwnd)
            return
        }
        MarkLastClipAsFavorite("first", true)
    }

    ; C key: copy response, then show Cursor window selector (1–9), activate selected window, focus AI field, paste and send.
    CopyAndTransferToCursor(*) {
        this.CleanupCopyBanner()
        ; Skip restoring focus so clipboard is not overwritten by the previously focused window before we read it.
        this.DoCopyCore(false, true)
        clipRaw := A_Clipboard
        clip := Trim(clipRaw)
        if (clip = "" || StrLen(clip) < GEMINI_TRANSFER_MIN_CLIPBOARD_LENGTH) {
            ShowCenteredOverlay_Utils("❌ Copy failed or empty – try again", 2000, BANNER_ACCENT_ERROR)
            if (WinExist("ahk_id " this.OriginalHwnd))
                WinActivate("ahk_id " this.OriginalHwnd)
            return
        }

        ; Restore the pre-handoff anchored window so the user sees the selector/paste in the exact context.
        if (this.OriginalHwnd && WinExist("ahk_id " this.OriginalHwnd)) {
            try {
                WinActivate("ahk_id " this.OriginalHwnd)
                ; Fast-path: avoid WinWaitActive if we are already active.
                if (!WinActive("ahk_id " this.OriginalHwnd))
                    WinWaitActive("ahk_id " this.OriginalHwnd, , 0.5)
            } catch {
            }
        }
        try A_Clipboard := clipRaw

        hwnd := CursorTransfer_ShowWindowSelector(0)
        if (!hwnd) {
            if (WinExist("ahk_id " this.OriginalHwnd))
                WinActivate("ahk_id " this.OriginalHwnd)
            try A_Clipboard := clipRaw
            return
        }
        ; Gemini → Cursor: no pre-movement warning (source is not Original).
        try A_Clipboard := clipRaw
        CursorTransfer_ActivateFocusPaste(hwnd, this.OriginalHwnd)
    }
}

; Current monitor instance so we can stop it when user chooses S or N at 6s (no copy/transfer follow-up).
global g_GeminiDelayedSubmitMonitor := ""

; Callable from Utils.ahk after successful auto-send (Ctrl+Alt+Win+L).
GeminiDelayedSubmitMonitorStart(originalHwnd, geminiHwnd) {
    global g_GeminiDelayedSubmitMonitor
    g_GeminiDelayedSubmitMonitor := GeminiDelayedSubmitMonitor()
    g_GeminiDelayedSubmitMonitor.Start(originalHwnd, geminiHwnd)
}

; Stop any running monitor so "Copy response?" does not show (e.g. after S or N at 6s dictation confirm).
GeminiDelayedSubmitMonitorStop() {
    global g_GeminiDelayedSubmitMonitor
    if (g_GeminiDelayedSubmitMonitor)
        try g_GeminiDelayedSubmitMonitor.Stop()
    g_GeminiDelayedSubmitMonitor := ""
}
