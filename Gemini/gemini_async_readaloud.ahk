; =============================================================================
; Gemini module: gemini_async_readaloud.ahk
; GeminiAsyncReadAloud async read-aloud class
; Extracted verbatim from Gemini.ahk; loaded via #include into the
; Gemini.ahk process, which remains the entry point / source of truth.
; =============================================================================

; =============================================================================
; GeminiAsyncReadAloud – async read aloud / pause / resume (D2C R / IPC / TTS)
; =============================================================================
class GeminiAsyncReadAloud {
    __New(copyFirst := true, useTrashTab := false, options := "") {
        this.CopyFirst := copyFirst
        this.UseTrashTab := useTrashTab
        this.OriginalHwnd := (options != "" && options.HasProp("originalHwnd")) ? options.originalHwnd : 0
        this.GeminiHwnd := (options != "" && options.HasProp("geminiHwnd")) ? options.geminiHwnd : 0
        this.AlreadyActive := (options != "" && options.HasProp("alreadyActive")) ? options.alreadyActive : false
        this.VerifyMaxRetries := (options != "" && options.HasProp("verifyMaxRetries")) ? options.verifyMaxRetries :
            GEMINI_READ_ALOUD_START_MAX_RETRIES
        this.TimerCallback := ""
        this.StartCallback := ""
        this.StepCallback := ""
        this.StartRetryAttempted := false
        this.StartRetryCount := 0
        this.StartTick := 0
        this.SkipCopy := false
        this.CopyRetryCount := 0
        this.MenuRetryCount := 0
        this.MenuPollCount := 0
        this.QueueTaskId := ""
        this.QueuePollCount := 0
        this.Uia := 0
        this.Started := false
        this.ListenPhaseReplays := 0
    }

    Start() {
        if (!this.OriginalHwnd)
            this.OriginalHwnd := GeminiResolveOriginalHwnd()
        if (!this.StartTick)
            this.StartTick := A_TickCount
        SetTitleMatchMode(2)
        if (!this.GeminiHwnd)
            this.GeminiHwnd := GetGeminiWindowHwnd()
        if (!this.GeminiHwnd) {
            ShowNotification("Read aloud failed – Gemini is not open", 1800, "FF6666", "FFFFFF", 22)
            return false
        }
        ; If Gemini is already foreground, skip deferred IPC/Launch — same fast path as #!+p-style "already there".
        if (this.GeminiHwnd && WinActive("ahk_id " this.GeminiHwnd))
            return this.TryStartReadAloud(false)
        ; IPC queue + pipe connect can block the calling thread (CreateFile on pipe waits for server).
        ; Defer so the caller (IPC / TTS / delayed submit) returns immediately like copy (#!+p); timers run the blocking work.
        this.StartCallback := this.DeferredQueueAndLaunch.Bind(this)
        SetTimer(this.StartCallback, -1)
        return true
    }

    ; Runs GeminiQueueBackgroundTask off the hotkey thread — avoids indefinite pipe-client blocking.
    DeferredQueueAndLaunch(*) {
        this.StartCallback := ""
        ; EnsureReady/Connect block on CreateFile until a pipe server exists — stalls the main thread and never reaches Launch.
        ; Only enqueue when IPC is enabled and we already have a pipe (#!+p and similar never touch this).
        ; Python IPC is only for consumer Gemini — not Copilot or Gemini Enterprise.
        queuedTask := false
        pipeOpen := GeminiIPC_HasOpenPipe()
        if (GEMINI_USE_PYTHON_IPC && pipeOpen && ResolveGlobalAICompanion() = "gemini")
            queuedTask := GeminiQueueBackgroundTask("ReadAloud", Map("geminiHwnd", this.GeminiHwnd, "originalHwnd",
                this.OriginalHwnd, "copyFirst", this.CopyFirst ? 1 : 0, "useTrashTab", this.UseTrashTab ? 1 : 0))
        if (queuedTask is Map && queuedTask.Has("taskId")) {
            this.QueueTaskId := String(queuedTask["taskId"])
            this.QueuePollCount := 0
            this.StartCallback := this.WaitForQueuedTask.Bind(this)
            SetTimer(this.StartCallback, -1)
            return
        }
        this.StartCallback := this.Launch.Bind(this)
        SetTimer(this.StartCallback, -1)
    }

    ScheduleStep(callback, delayMs := 1) {
        this.ClearStep()
        this.StepCallback := callback
        SetTimer(this.StepCallback, -delayMs)
    }

    ClearStep() {
        cb := ""
        try cb := this.StepCallback
        catch
            cb := ""
        if (cb)
            SetTimer(cb, 0)
        this.StepCallback := ""
    }

    WaitForQueuedTask(*) {
        this.StartCallback := ""
        if (!this.QueueTaskId) {
            this.Launch()
            return
        }
        resp := GeminiIPC_GetTaskStatus(this.QueueTaskId)
        if (GeminiIPC_ResponseOk(resp)) {
            result := GeminiIPC_ResponseResultMap(resp)
            status := result.Has("status") ? String(result["status"]) : ""
            if (status = "ready") {
                this.QueueTaskId := ""
                this.Launch()
                return
            }
        } else {
            this.QueueTaskId := ""
            this.Launch()
            return
        }
        this.QueuePollCount++
        if (this.QueuePollCount >= 40) {
            this.QueueTaskId := ""
            this.Launch()
            return
        }
        this.StartCallback := this.WaitForQueuedTask.Bind(this)
        SetTimer(this.StartCallback, -50)
    }

    Launch(*) {
        this.StartCallback := ""
        this.ListenPhaseReplays := 0
        this.TryStartReadAloud(false)
    }

    RetryLaunch(*) {
        this.StartCallback := ""
        this.ListenPhaseReplays := 0
        this.TryStartReadAloud(true)
    }

    TryStartReadAloud(skipCopy := false) {
        try {
            this.SkipCopy := skipCopy
            this.CopyRetryCount := 0
            this.MenuRetryCount := 0
            this.MenuPollCount := 0
            this.Uia := 0
            this.Started := false
            GeminiState.Invalidate()
            ; HWND can go stale after tab reload / window recreation; GetGeminiWindowHwnd re-resolves.
            if (!this.GeminiHwnd || !WinExist("ahk_id " this.GeminiHwnd)) {
                nh := GetGeminiWindowHwnd()
                if (nh)
                    this.GeminiHwnd := nh
            }
            if (!this.GeminiHwnd || !WinExist("ahk_id " this.GeminiHwnd)) {
                this.Fail()
                return false
            }
            ; Only show "Switching to Gemini" when another window is foreground — not when already in Gemini.
            if (!WinActive("ahk_id " this.GeminiHwnd)) {
                ; Loading Indication (not Hands off overlay): per standard_information_display.md
                StandardLoadingBar_Show("⏳ Switching to Gemini…", BANNER_ACCENT_INTERMEDIATE, { centerOnHwnd: this.GeminiHwnd ?
                    this.GeminiHwnd : 0 })
                if !GeminiActivateWindow(this.GeminiHwnd, GEMINI_ACTIVATE_WAIT_MS) {
                    try StandardLoadingBar_Hide(0)
                    ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
                    return false
                }
                this.ScheduleStep(this.AfterActivation.Bind(this), GEMINI_TAB_SWITCH_MS)
                return true
            }
            this.ScheduleStep(this.AfterActivation.Bind(this), 1)
            return true
        } catch {
            this.Fail()
            return false
        }
    }

    AfterActivation(*) {
        this.ClearStep()
        if (this.UseTrashTab) {
            Send("^2")
            ShowGeminiTabBanner(2, this.GeminiHwnd)
            this.ScheduleStep(this.BuildUIA.Bind(this), GEMINI_TAB_SWITCH_MS)
            return
        }
        this.ScheduleStep(this.BuildUIA.Bind(this), 1)
    }

    BuildUIA(*) {
        this.ClearStep()
        try
            this.Uia := UIA_Browser("ahk_id " this.GeminiHwnd)
        catch {
            this.Fail()
            return false
        }
        this.ScheduleStep(this.InspectControls.Bind(this), GEMINI_UIA_SETTLE_MS)
        return true
    }

    InspectControls(*) {
        this.ClearStep()
        try StandardLoadingBar_Hide(0) ; dismiss "⏳ Switching to Gemini…" before pause/scroll/listen paths
        pauseButton := FindGeminiPauseResumeButton(this.Uia, "Pause")
        if (pauseButton) {
            try pauseButton.Click()
            ShowNotification("Paused", 800, "FFFF00", "000000", 24)
            this.RestoreOriginalFocus()
            return true
        }

        resumeButton := FindGeminiPauseResumeButton(this.Uia, "Resume")
        if (resumeButton) {
            try resumeButton.Click()
            ShowNotification("Resumed", 800, "FFFF00", "000000", 24)
            this.RestoreOriginalFocus()
            return true
        }

        Send "^{End}"
        this.ScheduleStep(this.AfterScroll.Bind(this), GEMINI_SCROLL_SETTLE_MS)
        return true
    }

    AfterScroll(*) {
        this.ClearStep()
        if (this.CopyFirst && !this.SkipCopy) {
            this.ScheduleStep(this.RunCopyAttempt.Bind(this), 1)
            return
        }
        this.BeginListenMenuPhase()
    }

    RunCopyAttempt(*) {
        this.ClearStep()
        copyOpt := { restoreWindow: false, playChimeAndNotify: false, alreadyActive: true }
        if (CopyLastGeminiMessageToClipboard(copyOpt, this.GeminiHwnd)) {
            PlayCopyCompletedChime()
            this.BeginListenMenuPhase()
            return true
        }
        this.CopyRetryCount++
        if (this.CopyRetryCount < GEMINI_COPY_MAX_RETRIES) {
            backoffMs := GEMINI_COPY_RETRY_SLEEP_MS * (1 << (this.CopyRetryCount - 1))
            this.ScheduleStep(this.RunCopyAttempt.Bind(this), backoffMs)
            return false
        }
        this.BeginListenMenuPhase()
        return false
    }

    BeginListenMenuPhase() {
        StandardLoadingBar_Show(this.CopyFirst && !this.SkipCopy ? "🔍 Finding read aloud button and copying..." :
            "🔍 Finding read aloud button...", BANNER_ACCENT_INTERMEDIATE, { passive: true, centerOnHwnd: 0 })
        this.ScheduleStep(this.OpenListenMenu.Bind(this), GEMINI_WAIT_BUTTON_POLL_MS)
    }

    OpenListenMenu(*) {
        this.ClearStep()
        lastMoreOptionsButton := GetLastGeminiMoreOptionsButton(this.Uia)
        if (!lastMoreOptionsButton) {
            StandardLoadingBar_Hide(0)
            this.Fail()
            return false
        }
        try lastMoreOptionsButton.Click()
        catch {
            this.HandleListenMenuRetry()
            return false
        }
        this.MenuPollCount := 0
        this.ScheduleStep(this.WaitForListenMenuReady.Bind(this), GEMINI_MENU_OPEN_MS)
        return true
    }

    WaitForListenMenuReady(*) {
        this.ClearStep()
        listenMenuItem := GetLastGeminiListenMenuItem(this.Uia)
        if (listenMenuItem) {
            try {
                listenMenuItem.Click()
                StandardLoadingBar_Hide(0)
                ; Listen mutates the DOM (TTS / Pause). Cached UIA_Browser from BuildUIA is stale.
                this.Uia := 0
                ; Do not RestoreOriginalFocus here: keep Gemini foreground until CheckStarted confirms Pause.
                this.StartRetryCount := 0
                this.Started := false
                ; Common case (post-Gemini-reliability fix): Pause button is already present at click time.
                ; Run one synchronous check before paying the first 150ms timer tick; only poll if it misses.
                this.CheckStarted()
                if (!this.Started)
                    GeminiBackgroundSetTimer(this, this.CheckStarted.Bind(this), GEMINI_READ_ALOUD_START_POLL_MS)
                return true
            } catch {
                this.HandleListenMenuRetry()
                return false
            }
        }

        this.MenuPollCount++
        maxPolls := Ceil(GEMINI_LISTEN_MENU_WAIT_MS / GEMINI_WAIT_BUTTON_POLL_MS)
        if (this.MenuPollCount < maxPolls) {
            this.ScheduleStep(this.WaitForListenMenuReady.Bind(this), GEMINI_WAIT_BUTTON_POLL_MS)
            return false
        }
        this.HandleListenMenuRetry()
        return false
    }

    HandleListenMenuRetry() {
        this.ClearStep()
        this.MenuRetryCount++
        if (this.MenuRetryCount < GEMINI_LISTEN_MENU_MAX_ATTEMPTS) {
            SendEscape()
            this.ScheduleStep(this.OpenListenMenu.Bind(this), GEMINI_MENU_OPEN_MS)
            return
        }
        try StandardLoadingBar_Hide(0)
        this.Fail()
    }

    CheckStarted() {
        this.StartRetryCount++
        ; Re-resolve HWND if Chrome replaced the window; keeps RetryLaunch / verify from targeting a dead ahk_id.
        if (!this.GeminiHwnd || !WinExist("ahk_id " this.GeminiHwnd)) {
            nh := GetGeminiWindowHwnd()
            if (nh && nh != this.GeminiHwnd)
                this.Uia := 0
            this.GeminiHwnd := nh ? nh : this.GeminiHwnd
        }
        if (!this.GeminiHwnd || !WinExist("ahk_id " this.GeminiHwnd)) {
            GeminiBackgroundStopTimer(this)
            if (!this.StartRetryAttempted) {
                this.StartRetryAttempted := true
                ShowNotification("Retrying read aloud…", 800, "FFFF00", "000000", 24)
                this.StartCallback := this.RetryLaunch.Bind(this)
                SetTimer(this.StartCallback, -1)
                return
            }
            this.Fail()
            return
        }
        ; Keep Gemini foreground during Pause detection so UIA/DOM match what the user sees; no overlay.
        if (!WinActive("ahk_id " this.GeminiHwnd)) {
            try WinActivate("ahk_id " this.GeminiHwnd)
            try WinWaitActive("ahk_id " this.GeminiHwnd, , 0.4)
        }
        ; Reuse the UIA element attached in BuildUIA (canon §4 cache-first); only re-resolve once if it's gone.
        root := IsObject(this.Uia) ? this.Uia : 0
        if (!root)
            root := GeminiReadRootFromHwnd(this.GeminiHwnd)
        if (root) {
            try {
                if (FindGeminiPauseResumeButton(root, "Pause")) {
                    this.Started := true
                    this.ListenPhaseReplays := 0
                    GeminiBackgroundStopTimer(this)
                    GeminiPerfLog("read_aloud", this.StartTick)
                    GeminiEndAutomationSwitch("gemini_read_aloud_started")
                    ShowNotification(this.CopyFirst ? "Copied & Reading aloud" : "Reading aloud", 800, "FFFF00",
                        "000000", 24)
                    ; As soon as read-aloud is confirmed (Pause), return focus to the window that was active at hotkey
                    ; time (dictation R, TTS, …). Do not leave keyboard focus in Gemini unless there is nowhere else to go.
                    orig := this.OriginalHwnd
                    if (orig && orig != this.GeminiHwnd && WinExist("ahk_id " orig))
                        this.RestoreOriginalFocus()
                    else if (this.GeminiHwnd && WinActive("ahk_id " this.GeminiHwnd))
                        FocusGeminiAskFieldForHwnd(this.GeminiHwnd, false)
                    return
                }
            } catch {
            }
        }

        if (this.StartRetryCount < this.VerifyMaxRetries)
            return

        GeminiBackgroundStopTimer(this)
        ; Gemini often ignores the first Listen — replay Listen (scroll + menu) up to LISTEN_PHASE_MAX times before full RetryLaunch.
        if (this.ListenPhaseReplays < GEMINI_READ_ALOUD_LISTEN_PHASE_MAX - 1) {
            this.ListenPhaseReplays++
            this.StartRetryCount := 0
            this.ScheduleStep(this.BuildUIA.Bind(this), GEMINI_UIA_SETTLE_MS)
            return
        }
        if (!this.StartRetryAttempted) {
            this.StartRetryAttempted := true
            ShowNotification("Retrying read aloud...", 800, "FFFF00", "000000", 24)
            this.StartCallback := this.RetryLaunch.Bind(this)
            SetTimer(this.StartCallback, -1)
            return
        }
        this.Fail()
    }

    RestoreOriginalFocus() {
        this.ClearStep()
        this.AlreadyActive := false
        if (this.OriginalHwnd)
            GeminiRestoreWindow(this.OriginalHwnd)
        if (this.GeminiHwnd && WinActive("ahk_id " this.GeminiHwnd))
            FocusGeminiAskFieldForHwnd(this.GeminiHwnd, false)
    }

    Fail() {
        this.ClearStep()
        try StandardLoadingBar_Hide(0)
        ShowNotification("Read aloud failed – Gemini UI not ready", 2000, "FF6666", "FFFFFF", 22)
        this.RestoreOriginalFocus()
    }
}
