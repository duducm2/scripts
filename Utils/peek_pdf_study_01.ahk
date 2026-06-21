; =============================================================================
; Utils module: peek_pdf_study_01.ahk
; Peek PDF / QuickLook study helpers (part 1)
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; Peek PDF - Win+Alt+Shift+X
; If Peek is open: activate it. Otherwise: show study-topic selector (same aesthetic as Win+Alt+Shift+C).
; =============================================================================

; Study topics for Win+Alt+Shift+X selector. Paths are relative to notes repo (GetNotesRepoPath()).
; plansPath values match filenames in the notes repo (see studies/*/ *-plan.md, plan-english.md, learning-techniques.md).
global g_StudyTopics := Map(
    0, { name: "Technique (how to create studies)", mnemonicsPath: "\studies\technique\README.md",
        plansPath: "\studies\technique\plans.md" },
    1, { name: "Mnemonics", mnemonicsPath: "\studies\skills\mnemonics-skills.md",
        plansPath: "\studies\skills\learning-techniques.md" },
    2, { name: "Science", mnemonicsPath: "\studies\science\mnemonics-science.md",
        plansPath: "\studies\science\science-plan.md" },
    3, { name: "Piano", mnemonicsPath: "\studies\piano\mnemonics-piano.md",
        plansPath: "\studies\piano\piano-plan.md" },
    4, { name: "English", mnemonicsPath: "\studies\english\mnemonics-english.md",
        plansPath: "\studies\english\plan-english.md" },
    5, { name: "Communication", mnemonicsPath: "\studies\communication\mnemonics-communication.md",
        plansPath: "\studies\communication\communication-plan.md" },
    6, { name: "German", mnemonicsPath: "\studies\german\mnemonics-german.md",
        plansPath: "\studies\german\german-plan.md" }
)
#include %A_ScriptDir%\StudyArticleLink.ahk
#include %A_ScriptDir%\StudyFavoriteLink.ahk

global g_StudyTopicSelectorGui := false
global g_StudyTopicSelectorActive := false
global g_StudyTopicSelectorPhase := ""           ; "category" | "topic"
global g_StudyTopicSelectorCategory := ""        ; "mnemonics" | "plans"
global g_StudyTopicSelectorLastForegroundMonitorIdx := 0   ; for trackActiveMonitor-style follow (standard_information_display.md)
global g_StudyTopicEscPollPrev := false   ; edge-detect Esc for StudyTopicSelector_EscapePoll (parity with OutlookCopilotSelector)
global STUDY_TOPIC_BLACKOUT_DELAY_MS := 3000
; Study Topic QuickLook: strict bounded waits + shared layout (false = legacy 2s WinWait + inline scroll).
global STUDY_TOPIC_QL_STRICT_LAYOUT := true
global g_QuickLookDeferredLayoutScroll := true
global g_QuickLookDeferredLayoutPath := ""

; PDF focus monitoring for automatic blackout cancellation (Win+Alt+Shift+X)
global g_PdfFocusMonitorTimer := false
global g_PdfFocusTrackedHwnd := 0
global g_PdfFocusLossMode := "Immediate"      ; "Immediate" or "Debounced"
global g_PdfFocusDebounceMs := 1200            ; Allow transient focus loss without un-blackouting
global g_PdfFocusLostSinceTick := 0

; Monitor foreground vs keep-clear display: relocate only when anchor HWND moves; else disable.
MonitorPdfFocus() {
    global g_FocusModeOn, g_FocusModeActiveMonitor, g_PdfFocusTrackedHwnd, g_FocusModeTrackedWindow
    global g_FocusModeAnchorHwnd, g_PdfFocusLossMode, g_PdfFocusDebounceMs, g_PdfFocusLostSinceTick

    if (!g_FocusModeOn) {
        if (g_PdfFocusTrackedHwnd)
            StopPdfFocusMonitor()
        return
    }

    if (FocusMode_CheckCrossProcessRequests())
        return

    fg := WinExist("A")
    if (!fg)
        return

    fgMon := GetActiveMonitorIndex()
    if (!fgMon)
        return

    keepMon := g_FocusModeActiveMonitor
    if (keepMon && fgMon != keepMon) {
        anchor := g_FocusModeAnchorHwnd
        sameAnchor := anchor && fg = anchor && WinExist("ahk_id " . anchor)
        if (sameAnchor) {
            if (g_PdfFocusLossMode = "Debounced") {
                if (!g_PdfFocusLostSinceTick)
                    g_PdfFocusLostSinceTick := A_TickCount
                if ((A_TickCount - g_PdfFocusLostSinceTick) >= g_PdfFocusDebounceMs) {
                    FocusMode_SetKeepMonitor(fgMon)
                    g_PdfFocusLostSinceTick := 0
                }
            } else {
                FocusMode_SetKeepMonitor(fgMon)
                g_PdfFocusLostSinceTick := 0
            }
        } else {
            DisableFocusMode()
            return
        }
    } else {
        g_PdfFocusLostSinceTick := 0
    }

    g_FocusModeTrackedWindow := fg
    g_PdfFocusTrackedHwnd := fg
}

; Start monitoring PDF window focus
StartPdfFocusMonitor(hwnd := 0, focusLossMode := "Immediate") {
    global g_PdfFocusMonitorTimer, g_PdfFocusTrackedHwnd, g_PdfFocusLossMode, g_PdfFocusLostSinceTick
    global g_FocusModeAnchorHwnd

    StopPdfFocusMonitor()

    g_PdfFocusTrackedHwnd := hwnd ? hwnd : WinExist("A")
    if (!g_PdfFocusTrackedHwnd)
        return

    if (hwnd)
        g_FocusModeAnchorHwnd := hwnd
    else if (!g_FocusModeAnchorHwnd || !WinExist("ahk_id " . g_FocusModeAnchorHwnd))
        g_FocusModeAnchorHwnd := g_PdfFocusTrackedHwnd

    g_PdfFocusLossMode := focusLossMode
    g_PdfFocusLostSinceTick := 0

    g_PdfFocusMonitorTimer := MonitorPdfFocus
    SetTimer(g_PdfFocusMonitorTimer, 200)
}

; Stop monitoring PDF window focus
StopPdfFocusMonitor() {
    global g_PdfFocusMonitorTimer, g_PdfFocusTrackedHwnd

    ; First-call safety: global may be unset
    if (!IsSet(g_PdfFocusMonitorTimer))
        g_PdfFocusMonitorTimer := false

    if (g_PdfFocusMonitorTimer) {
        SetTimer(g_PdfFocusMonitorTimer, 0)
        g_PdfFocusMonitorTimer := false
    }
    g_PdfFocusTrackedHwnd := 0
}

StudyTopic_GetBlackoutKeepMonitorIndex() {
    try return MonitorGetPrimary()
    catch
        return 1
}

; Shared countdown flag for FocusBlackoutWatcher (Study Topic path clears without setting).
BlackoutCountdown_Begin() {
    global g_FocusBlackoutWatcherCountdownActive
    g_FocusBlackoutWatcherCountdownActive := true
}

BlackoutCountdown_End() {
    global g_FocusBlackoutWatcherCountdownActive
    g_FocusBlackoutWatcherCountdownActive := false
}

StudyTopic_CancelBlackoutCountdown(targetHwnd := 0, *) {
    global g_FocusBlackoutWatcherDeniedHwnd
    BlackoutCountdown_End()
    if (targetHwnd && WinExist("ahk_id " . targetHwnd))
        g_FocusBlackoutWatcherDeniedHwnd := targetHwnd
    else {
        fg := WinExist("A")
        if (fg)
            g_FocusBlackoutWatcherDeniedHwnd := fg
    }
    StandardLoadingBar_CloseKeysOverlay()
    StandardLoadingBar_Hide(0)
}

; Apply blackout using foreground at timeout (user may juggle monitors during the 3s banner).
StudyTopic_ApplyBlackoutCountdownTimeout(targetHwnd := 0, pdfFocusLossMode := "Debounced") {
    global g_BlackoutSuppressedUntil, g_FocusModeOn

    BlackoutCountdown_End()

    fg := WinExist("A")
    if (!fg && targetHwnd && WinExist("ahk_id " . targetHwnd))
        fg := targetHwnd
    if (!fg)
        return

    if (Blackout_IsSuppressed())
        return

    keepIdx := GetActiveMonitorIndex()
    if (!keepIdx)
        keepIdx := StudyTopic_GetBlackoutKeepMonitorIndex()
    if (g_FocusModeOn)
        FocusMode_SetKeepMonitor(keepIdx)
    else
        EnableFocusMode(keepIdx)
    StartPdfFocusMonitor(fg, pdfFocusLossMode)
}

; --- Blackout suppression logic ---
global g_BlackoutSuppressedUntil := 0
global BLACKOUT_SUPPRESS_MS := 7 * 60 * 1000

Blackout_IsSuppressed() {
    global g_BlackoutSuppressedUntil
    if (!g_BlackoutSuppressedUntil)
        return false
    return (g_BlackoutSuppressedUntil - A_TickCount) > 0
}

; D key handler: suppress all blackout banners for BLACKOUT_SUPPRESS_MS; reset dwell for post-suppress window.
Blackout_Disable7Min(*) {
    global g_BlackoutSuppressedUntil, g_FocusBlackoutWatcherDwellStartTick
    g_BlackoutSuppressedUntil := A_TickCount + BLACKOUT_SUPPRESS_MS
    g_FocusBlackoutWatcherDwellStartTick := A_TickCount
    BlackoutCountdown_End()
    FocusBlackoutWatcher_DebugLog("Blackout_Disable7Min until=" . g_BlackoutSuppressedUntil . " tick=" . A_TickCount .
        " remaining=" . (g_BlackoutSuppressedUntil - A_TickCount))
    StandardLoadingBar_CloseKeysOverlay()
    StandardLoadingBar_Hide(0)
}

StudyTopic_StartBlackoutCountdown(targetHwnd) {
    if (!targetHwnd || !WinExist("ahk_id " . targetHwnd))
        return
    if (Blackout_IsSuppressed()) {
        BlackoutCountdown_End()
        return
    }
    BlackoutCountdown_Begin()
    global g_FocusBlackoutWatcherDwellStartTick
    g_FocusBlackoutWatcherDwellStartTick := A_TickCount
    StandardLoadingBar_CloseKeysOverlay()
    StandardLoadingBar_Hide(0)
    Sleep 50
    cancelCb := StudyTopic_CancelBlackoutCountdown.Bind(targetHwnd)
    keyCallbacks := Map("N", cancelCb, "*Escape", cancelCb)
    keyCallbacks["D"] := Blackout_Disable7Min
    timeoutCb := StudyTopic_ApplyBlackoutCountdownTimeout.Bind(, "Immediate")
    StandardLoadingBar_ShowWithKeys(
        "⏳ Blacking out secondary monitors in 3s",
        keyCallbacks,
        STUDY_TOPIC_BLACKOUT_DELAY_MS,
        0,
        timeoutCb,
        BANNER_BLACKOUT_BORDER,
        420,
        17,
        BANNER_BLACKOUT_BORDER,
        false,
        "[N] Cancel blackout    [D] Disable for 7 min",
        true,
        true,
        true,
        BANNER_BLACKOUT_PANEL)
}

; Focus dwell watcher: after continuous foreground on one window, offer same blackout banner as Study Topic (#!+X).
global g_FocusBlackoutWatcherStarted := false
global g_FocusBlackoutWatcherLastHwnd := 0
global g_FocusBlackoutWatcherDwellStartTick := 0
global g_FocusBlackoutWatcherDeniedHwnd := 0
global g_FocusBlackoutWatcherCountdownActive := false
global FOCUS_BLACKOUT_DWELL_MS := 20000
global FOCUS_BLACKOUT_DEBUG_LOG := false

FocusBlackoutWatcher_DebugLog(message) {
    global FOCUS_BLACKOUT_DEBUG_LOG
    if (!FOCUS_BLACKOUT_DEBUG_LOG)
        return
    try {
        FileAppend "[FBW] " . message . "`n", A_ScriptDir "\focus_blackout_debug.log", "UTF-8"
    } catch {
    }
}

FocusBlackoutWatcher_OnCancel(hwnd, *) {
    StudyTopic_CancelBlackoutCountdown(hwnd)
}

FocusBlackoutWatcher_OnBlackoutTimeout(hwnd, *) {
    StudyTopic_ApplyBlackoutCountdownTimeout(, "Immediate")
}

FocusBlackoutWatcher_StartCountdown(hwnd) {
    global g_FocusBlackoutWatcherDwellStartTick
    if (!hwnd || !WinExist("ahk_id " . hwnd))
        return
    if (Blackout_IsSuppressed()) {
        BlackoutCountdown_End()
        FocusBlackoutWatcher_DebugLog("StartCountdown skipped (suppressed)")
        return
    }
    BlackoutCountdown_Begin()
    g_FocusBlackoutWatcherDwellStartTick := A_TickCount
    FocusBlackoutWatcher_DebugLog("StartCountdown for hwnd " . hwnd)
    StandardLoadingBar_CloseKeysOverlay()
    StandardLoadingBar_Hide(0)
    Sleep 50
    cancelCb := FocusBlackoutWatcher_OnCancel.Bind(hwnd)
    keyCallbacks := Map("N", cancelCb, "*Escape", cancelCb)
    keyCallbacks["D"] := Blackout_Disable7Min
    timeoutCb := FocusBlackoutWatcher_OnBlackoutTimeout.Bind(hwnd)  ; hwnd unused at apply; foreground at timeout wins
    StandardLoadingBar_ShowWithKeys(
        "⏳ Blacking out secondary monitors in 3s",
        keyCallbacks,
        STUDY_TOPIC_BLACKOUT_DELAY_MS,
        0,
        timeoutCb,
        BANNER_BLACKOUT_BORDER,
        420,
        17,
        BANNER_BLACKOUT_BORDER,
        false,
        "[N] Cancel blackout    [D] Disable for 7 min",
        true,
        true,
        true,
        BANNER_BLACKOUT_PANEL)
}

FocusBlackoutWatcher_Tick() {
    global g_FocusBlackoutWatcherLastHwnd, g_FocusBlackoutWatcherDwellStartTick
    global g_FocusBlackoutWatcherDeniedHwnd, g_FocusBlackoutWatcherCountdownActive, FOCUS_BLACKOUT_DWELL_MS
    global g_FocusModeOn, g_FocusModeActiveMonitor, g_FocusModeTrackedWindow

    try {
        if (MonitorGetCount() <= 1)
            return
    } catch {
        return
    }

    try {
        hwnd := WinExist("A")
        if (!hwnd) {
            g_FocusBlackoutWatcherLastHwnd := 0
            FocusBlackoutWatcher_DebugLog("No active window")
            return
        }

        if (g_FocusBlackoutWatcherDeniedHwnd && hwnd != g_FocusBlackoutWatcherDeniedHwnd) {
            FocusBlackoutWatcher_DebugLog("Reset denied hwnd (user switched window)")
            g_FocusBlackoutWatcherDeniedHwnd := 0
        }

        if (hwnd != g_FocusBlackoutWatcherLastHwnd) {
            FocusBlackoutWatcher_DebugLog("Window changed. Reset dwell timer.")
            g_FocusBlackoutWatcherLastHwnd := hwnd
            g_FocusBlackoutWatcherDwellStartTick := A_TickCount
            return
        }

        ; Suppression before countdown-active so dwell resets even if the flag was left set (e.g. Esc).
        if (Blackout_IsSuppressed()) {
            BlackoutCountdown_End()
            FocusBlackoutWatcher_DebugLog("Blackout suppressed. Reset dwell timer.")
            g_FocusBlackoutWatcherDwellStartTick := A_TickCount
            return
        }

        if (g_FocusBlackoutWatcherCountdownActive) {
            FocusBlackoutWatcher_DebugLog("Countdown already active")
            return
        }

        if (g_FocusBlackoutWatcherDeniedHwnd && hwnd = g_FocusBlackoutWatcherDeniedHwnd) {
            FocusBlackoutWatcher_DebugLog("Blackout denied for this hwnd")
            return
        }

        if (g_FocusModeOn && g_FocusModeActiveMonitor) {
            curMon := GetActiveMonitorIndex()
            if (curMon && curMon = g_FocusModeActiveMonitor) {
                FocusBlackoutWatcher_DebugLog("Focus mode already on for this monitor")
                return
            }
        }

        elapsed := (A_TickCount - g_FocusBlackoutWatcherDwellStartTick)
        if (elapsed >= FOCUS_BLACKOUT_DWELL_MS) {
            FocusBlackoutWatcher_DebugLog("Dwell met (" . elapsed . " ms). Starting blackout countdown for hwnd " .
                hwnd)
            FocusBlackoutWatcher_StartCountdown(hwnd)
            return
        }
        FocusBlackoutWatcher_DebugLog("Dwell not yet met: " . elapsed . " ms")
    } catch as e {
        FocusBlackoutWatcher_DebugLog("tick error: " . e.Message)
    }
}

FocusBlackoutWatcher_Start() {
    global g_FocusBlackoutWatcherStarted
    if (g_FocusBlackoutWatcherStarted)
        return
    g_FocusBlackoutWatcherStarted := true
    SetTimer(FocusBlackoutWatcher_Tick, 200)
}

FocusBlackoutWatcher_Stop() {
    global g_FocusBlackoutWatcherStarted
    if (!g_FocusBlackoutWatcherStarted)
        return
    SetTimer(FocusBlackoutWatcher_Tick, 0)
    g_FocusBlackoutWatcherStarted := false
}

; YouTube focus session (Win+Alt+Shift+H): toggle on/off; SMTC for Spotify play/pause (not toggle).
global g_YoutubeFocusMonitorTimer := false
global g_YoutubeFocusTrackedHwnd := 0
global g_YoutubeSpotifyPausePending := false
global g_YoutubeFocusSessionActive := false

; Find Spotify's Windows.Media.Control session (SourceAppUserModelId contains "Spotify").
YouTube_FindSpotifyMediaSession() {
    try {
        for session in Media.GetSessions() {
            try {
                id := session.SourceAppUserModelId
                if InStr(id, "Spotify")
                    return session
            } catch {
                continue
            }
        }
    } catch {
        return 0
    }
    return 0
}

; Pause Spotify via SMTC only when status is Playing (avoids starting playback when already paused).
YouTube_PauseSpotifyBeforeYoutube() {
    global g_YoutubeSpotifyPausePending
    if (g_YoutubeSpotifyPausePending)
        return
    try {
        session := YouTube_FindSpotifyMediaSession()
        if !session
            return
        if (session.PlaybackStatus != Media.PlaybackStatus.Playing)
            return
        session.Pause()
        g_YoutubeSpotifyPausePending := true
    } catch {
        ; WinRT/SMTC unavailable - do not fall back to Media_Play_Pause (toggle bug).
    }
}

; Resume Spotify with SMTC Play() only when we paused it and session reports Paused.
YouTube_ResumeSpotifyAfterYoutubeIfPending(restoreHwnd := 0) {
    global g_YoutubeSpotifyPausePending
    if (!g_YoutubeSpotifyPausePending)
        return
    g_YoutubeSpotifyPausePending := false
    try {
        session := YouTube_FindSpotifyMediaSession()
        if (session && session.PlaybackStatus == Media.PlaybackStatus.Paused)
            session.Play()
    } catch {
        ;
    }
    if (restoreHwnd && WinExist("ahk_id " restoreHwnd)) {
        try
            WinActivate("ahk_id " restoreHwnd)
    }
}

; End YouTube focus session: pause YouTube (k), resume Spotify if pending, stop window monitor.
YouTube_EndFocusSession() {
    global g_YoutubeFocusTrackedHwnd, g_YoutubeFocusSessionActive
    restoreHwnd := WinExist("A")
    if (g_YoutubeFocusTrackedHwnd && WinExist("ahk_id " . g_YoutubeFocusTrackedHwnd)) {
        try {
            WinActivate("ahk_id " . g_YoutubeFocusTrackedHwnd)
            Sleep(50)
            Send("k")
            Sleep(100)
        }
    }
    YouTube_ResumeSpotifyAfterYoutubeIfPending(restoreHwnd)
    StopYoutubeFocusMonitor()
    g_YoutubeFocusSessionActive := false
}

; Monitor tracked YouTube window: only when it is destroyed, run full session teardown (same as second hotkey).
MonitorYoutubeFocus() {
    global g_YoutubeFocusTrackedHwnd

    if (g_YoutubeFocusTrackedHwnd && !WinExist("ahk_id " . g_YoutubeFocusTrackedHwnd))
        YouTube_EndFocusSession()
}

; Start monitoring YouTube window focus
StartYoutubeFocusMonitor(hwnd := 0) {
    global g_YoutubeFocusMonitorTimer, g_YoutubeFocusTrackedHwnd

    StopYoutubeFocusMonitor()

    g_YoutubeFocusTrackedHwnd := hwnd ? hwnd : WinExist("A")
    if (!g_YoutubeFocusTrackedHwnd)
        return

    g_YoutubeFocusMonitorTimer := MonitorYoutubeFocus
    SetTimer(g_YoutubeFocusMonitorTimer, 200)
}

; Stop monitoring YouTube window focus
StopYoutubeFocusMonitor() {
    global g_YoutubeFocusMonitorTimer, g_YoutubeFocusTrackedHwnd

    if (g_YoutubeFocusMonitorTimer) {
        SetTimer(g_YoutubeFocusMonitorTimer, 0)
        g_YoutubeFocusMonitorTimer := false
    }
    g_YoutubeFocusTrackedHwnd := 0
}

StudyTopic_GetRelPath(topic, category) {
    if (category = "plans")
        return topic.plansPath
    return topic.mnemonicsPath
}

; Opens notes-repo-relative path in QuickLook (PDF sibling → .md). Returns false on failure.
; scrollToEnd: mnemonics jump to bottom of long docs; plans stay at top.
StudyTopic_OpenRepoRelativeMarkdown(relPath, scrollToEnd := true) {
    basePath := GetNotesRepoPath()
    if (basePath = "") {
        try ShowCenteredOverlay_Utils("⚠ Notes repo path not set (env.ahk).", 3000, BANNER_ACCENT_INTERMEDIATE)
        return false
    }
    fullPath := RTrim(basePath, "\") . relPath
    if (StrLower(SubStr(fullPath, -3)) = "pdf") {
        mdPath := SubStr(fullPath, 1, StrLen(fullPath) - 3) . "md"
    } else {
        mdPath := fullPath
    }
    if (!FileExist(mdPath)) {
        try ShowCenteredOverlay_Utils("❌ Markdown not found: " mdPath, 3500, BANNER_ACCENT_ERROR)
        return false
    }
    QuickLook_OpenPath(mdPath, scrollToEnd)
    return true
}

; Center in work-area rect using Round + clamp (horizontal matches StandardLoadingBar_RepositionToActiveMonitor; vertical added for true center).
StudyTopicSelector_ComputeCenterTopLeftInWorkArea(ml, mt, mr, mb, gw, gh, &cx, &cy) {
    monitorWidth := mr - ml
    monitorHeight := mb - mt
    cx := Round(ml + (monitorWidth - gw) / 2)
    if (cx < ml)
        cx := ml
    if (cx + gw > mr)
        cx := mr - gw
    cy := Round(mt + (monitorHeight - gh) / 2)
    if (cy < mt)
        cy := mt
    if (cy + gh > mb)
        cy := mb - gh
}

; Initial placement: GetActiveMonitorWorkArea_StandardBar (same source as StandardLoadingBar / standard_information_display.md); Outlook Copilot uses an equivalent MonitorGetWorkArea loop in Shift keys.ahk.
StudyTopicSelector_PositionGuiLikeOutlook(gui) {
    GetActiveMonitorWorkArea_StandardBar(&ml, &mt, &mr, &mb)
    gui.Show("AutoSize Hide")
    gui.GetPos(, , &gw, &gh)
    StudyTopicSelector_ComputeCenterTopLeftInWorkArea(ml, mt, mr, mb, gw, gh, &cx, &cy)
    ; Avoid "NA": if focus stays in QuickLook, Esc is consumed there first (ShowOutlookCopilotSelector comment).
    gui.Show("x" . cx . " y" . cy)
    try WinActivate(gui.Hwnd)
}

StudyTopicSelector_StopActiveMonitorTracking() {
    try SetTimer(StudyTopicSelector_TrackActiveMonitorTick, 0)
}

; forMonitorIdx: 1-based index from StudyTopicSelector_TrackActiveMonitorTick (same as GetMonitorIndexForForeground_StandardBar).
; Use Show("AutoSize Hide") then Show("x y") like StudyTopicSelector_PositionGuiLikeOutlook — Move() alone can mis-center across mixed-DPI monitors.
StudyTopicSelector_RepositionToActiveMonitor(forMonitorIdx := 0) {
    global g_StudyTopicSelectorGui, g_StudyTopicSelectorActive
    if (!IsObject(g_StudyTopicSelectorGui) || !g_StudyTopicSelectorGui.Hwnd)
        return
    idx := forMonitorIdx
    if (idx < 1 || idx > MonitorGetCount())
        idx := GetMonitorIndexForForeground_StandardBar()
    MonitorGetWorkArea(idx, &ml, &mt, &mr, &mb)
    try {
        g_StudyTopicSelectorGui.Show("AutoSize Hide")
        g_StudyTopicSelectorGui.GetPos(, , &gw, &gh)
    } catch {
        return
    }
    StudyTopicSelector_ComputeCenterTopLeftInWorkArea(ml, mt, mr, mb, gw, gh, &cx, &cy)
    try {
        g_StudyTopicSelectorGui.Show("x" . cx . " y" . cy)
        if (g_StudyTopicSelectorActive)
            try WinActivate(g_StudyTopicSelectorGui.Hwnd)
    } catch {
    }
}

; Follow foreground window's monitor while the selector is open (parity with StandardLoadingBar trackActiveMonitor).
StudyTopicSelector_TrackActiveMonitorTick() {
    global g_StudyTopicSelectorGui, g_StudyTopicSelectorActive, g_StudyTopicSelectorLastForegroundMonitorIdx
    if (!g_StudyTopicSelectorActive || !IsObject(g_StudyTopicSelectorGui) || !g_StudyTopicSelectorGui.Hwnd) {
        StudyTopicSelector_StopActiveMonitorTracking()
        return
    }
    newIdx := GetMonitorIndexForForeground_StandardBar()
    if (newIdx != g_StudyTopicSelectorLastForegroundMonitorIdx) {
        StudyTopicSelector_RepositionToActiveMonitor(newIdx)
        g_StudyTopicSelectorLastForegroundMonitorIdx := newIdx
    }
}

StudyTopicSelector_UnbindCategoryHotkeys() {
    try Hotkey("0", "Off")
    try Hotkey("1", "Off")
    try Hotkey("2", "Off")
    try Hotkey("3", "Off")
    try Hotkey("4", "Off")
    try Hotkey("5", "Off")
    try Hotkey("6", "Off")
}

StudyTopicSelector_UnbindDigitHotkeys() {
    loop 7 {
        try Hotkey(String(A_Index - 1), "Off")
    }
}

; Poll Esc — fallback when $*Escape / g_OnEscapePressed miss (OutlookCopilotSelector_EscapePoll).
StudyTopicSelector_EscapePoll() {
    global g_StudyTopicSelectorActive, g_StudyTopicEscPollPrev
    if (!g_StudyTopicSelectorActive) {
        SetTimer(StudyTopicSelector_EscapePoll, 0)
        return
    }
    escSync := GetKeyState("Escape", "P")
    escAsync := (DllCall("user32\GetAsyncKeyState", "int", 0x1B) & 0x8000) != 0
    escDown := escSync || escAsync
    if (escDown) {
        if (!g_StudyTopicEscPollPrev) {
            g_StudyTopicEscPollPrev := true
            StudyTopicSelector_Cancel()
        }
    } else {
        g_StudyTopicEscPollPrev := false
    }
}

StudyTopicSelector_EscapeFromHotkey(*) {
    StudyTopicSelector_Cancel()
}

StudyTopicSelector_GlobalEscapeCallback(*) {
    StudyTopicSelector_Cancel()
}

StudyTopicSelector_GuiEscape(*) {
    StudyTopicSelector_Cancel()
}

; Same escape registration order as ShowOutlookCopilotSelector (after Gui.OnEvent registered at build time).
