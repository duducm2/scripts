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
; mnemonicsUrl / plansUrl: GitHub blob URLs (links.md). Menu [1]/[2]/[6] open via Chrome --new-window.
; plansPath values match filenames in the notes repo (see studies/*/ *-plan.md, plan-english.md, learning-techniques.md).
global g_StudyTopics := Map(
    0, { name: "Technique (how to create studies)", mnemonicsPath: "\studies\technique\README.md",
        plansPath: "\studies\technique\plans.md",
        mnemonicsUrl: "https://github.com/duducm2/my-notes/blob/main/studies/technique/README.md",
        plansUrl: "" },
    1, { name: "Skills", mnemonicsPath: "\studies\skills\mnemonics-skills.md",
        plansPath: "\studies\skills\skills-plan.md",
        mnemonicsUrl: "https://github.com/duducm2/my-notes/blob/main/studies/skills/mnemonics-skills.md",
        plansUrl: "https://github.com/duducm2/my-notes/blob/main/studies/skills/skills-plan.md" },
    2, { name: "Science", mnemonicsPath: "\studies\science\mnemonics-science.md",
        plansPath: "\studies\science\science-plan.md",
        mnemonicsUrl: "https://github.com/duducm2/my-notes/blob/main/studies/science/mnemonics-science.md",
        plansUrl: "https://github.com/duducm2/my-notes/blob/main/studies/science/science-plan.md" },
    3, { name: "Piano", mnemonicsPath: "\studies\piano\mnemonics-piano.md",
        plansPath: "\studies\piano\piano-plan.md",
        mnemonicsUrl: "https://github.com/duducm2/my-notes/blob/main/studies/Piano/mnemonics-piano.md",
        plansUrl: "https://github.com/duducm2/my-notes/blob/main/studies/Piano/piano-plan.md" },
    4, { name: "English", mnemonicsPath: "\studies\english\mnemonics-english.md",
        plansPath: "\studies\english\plan-english.md",
        mnemonicsUrl: "https://github.com/duducm2/my-notes/blob/main/studies/English/mnemonics-english.md",
        plansUrl: "https://github.com/duducm2/my-notes/blob/main/studies/English/english-plan.md" },
    5, { name: "Communication", mnemonicsPath: "\studies\communication\mnemonics-communication.md",
        plansPath: "\studies\communication\communication-plan.md",
        mnemonicsUrl: "https://github.com/duducm2/my-notes/blob/main/studies/Communication/mnemonics-communication.md",
        plansUrl: "https://github.com/duducm2/my-notes/blob/main/studies/Communication/communication-plan.md" },
    6, { name: "German", mnemonicsPath: "\studies\german\mnemonics-german.md",
        plansPath: "\studies\german\german-plan.md",
        mnemonicsUrl: "https://github.com/duducm2/my-notes/blob/main/studies/german/mnemonics-german.md",
        plansUrl: "https://github.com/duducm2/my-notes/blob/main/studies/german/german-plan.md" },
    7, { name: "Entertainment", mnemonicsPath: "", plansPath: "\studies\entertainment\entertainment-plan.md",
        mnemonicsUrl: "", plansUrl: "" }
)
#include %A_ScriptDir%\lib\study\StudyArticleLink.ahk
#include %A_ScriptDir%\lib\study\StudyFavoriteLink.ahk

global g_StudyTopicSelectorGui := false
global g_StudyTopicSelectorActive := false
global g_StudyTopicSelectorPhase := ""           ; "category" | "topic"
global g_StudyTopicSelectorCategory := ""        ; "mnemonics" | "plans"
global g_StudyTopicSelectorLastForegroundMonitorIdx := 0   ; for trackActiveMonitor-style follow (standard_information_display.md)
global g_StudyTopicEscPollPrev := false   ; edge-detect Esc for StudyTopicSelector_EscapePoll (parity with ShowAiModelSelector)
global g_StudyTopicSelectorPendingRemove := false
; category ("mnemonics"|"plans") → Array of { name, url }; filled on first StudyLinks_GetEntries.
global g_StudyLinksCache := Map()
; Study Topic QuickLook: strict bounded waits + shared layout (false = legacy 2s WinWait + inline scroll).
global STUDY_TOPIC_QL_STRICT_LAYOUT := true
global g_QuickLookDeferredLayoutScroll := true
global g_QuickLookDeferredLayoutPath := ""

; PDF focus monitoring for focus-mode relocate/disable (#!+Y)
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

; --- Blackout suppression logic (Outlook reminder banners, etc.) ---
global g_BlackoutSuppressedUntil := 0
global BLACKOUT_SUPPRESS_MS := 7 * 60 * 1000

Blackout_IsSuppressed() {
    global g_BlackoutSuppressedUntil
    if (!g_BlackoutSuppressedUntil)
        return false
    return (g_BlackoutSuppressedUntil - A_TickCount) > 0
}

; D key handler: suppress blackout-related banners for BLACKOUT_SUPPRESS_MS.
Blackout_Disable7Min(*) {
    global g_BlackoutSuppressedUntil
    g_BlackoutSuppressedUntil := A_TickCount + BLACKOUT_SUPPRESS_MS
    StandardLoadingBar_CloseKeysOverlay()
    StandardLoadingBar_Hide(0)
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

StudyTopic_GetUrl(topic, category) {
    if (category = "plans")
        return topic.plansUrl
    return topic.mnemonicsUrl
}

; Dynamic Mnemonics/Plans lists — shares study_links.ini with [Links] API keys; uses [Mnemonics]/[Plans].
StudyLinks_GetIniPath() {
    return A_ScriptDir "\assets\data\study_links.ini"
}

StudyLinks_SectionName(category) {
    return (category = "plans") ? "Plans" : "Mnemonics"
}

StudyLinks_SeedFromTopics(category) {
    global g_StudyTopics
    entries := []
    for num, topic in g_StudyTopics {
        n := Integer(num)
        if (n = 0)
            continue
        if (category = "mnemonics") {
            url := topic.mnemonicsUrl
            if (url = "" && topic.mnemonicsPath = "")
                continue
            if (url = "")
                continue
            entries.Push({ name: topic.name, url: url })
        } else {
            ; Include Entertainment even with empty URL (placeholder).
            entries.Push({ name: topic.name, url: topic.plansUrl })
        }
    }
    return entries
}

StudyLinks_ParseSection(raw) {
    entries := []
    byKey := Map()
    maxKey := 0
    for line in StrSplit(raw, "`n", "`r") {
        line := Trim(line)
        if (line = "" || SubStr(line, 1, 1) = ";")
            continue
        eq := InStr(line, "=")
        if (!eq)
            continue
        keyStr := Trim(SubStr(line, 1, eq - 1))
        val := Trim(SubStr(line, eq + 1))
        if (!RegExMatch(keyStr, "^\d+$"))
            continue
        k := Integer(keyStr)
        pipe := InStr(val, "|")
        if (pipe) {
            name := Trim(SubStr(val, 1, pipe - 1))
            url := Trim(SubStr(val, pipe + 1))
        } else {
            name := val
            url := ""
        }
        byKey[k] := { name: name, url: url }
        if (k > maxKey)
            maxKey := k
    }
    loop maxKey {
        if byKey.Has(A_Index)
            entries.Push(byKey[A_Index])
    }
    return entries
}

StudyLinks_GetEntries(category) {
    global g_StudyLinksCache
    if (category != "mnemonics" && category != "plans")
        return []
    if g_StudyLinksCache.Has(category)
        return g_StudyLinksCache[category]

    iniPath := StudyLinks_GetIniPath()
    section := StudyLinks_SectionName(category)
    raw := ""
    try raw := IniRead(iniPath, section)
    catch {
        raw := ""
    }
    if (raw = "" || raw = "ERROR") {
        entries := StudyLinks_SeedFromTopics(category)
        StudyLinks_SaveEntries(category, entries)
        return g_StudyLinksCache[category]
    }
    entries := StudyLinks_ParseSection(raw)
    if (entries.Length = 0) {
        entries := StudyLinks_SeedFromTopics(category)
        StudyLinks_SaveEntries(category, entries)
        return g_StudyLinksCache[category]
    }
    g_StudyLinksCache[category] := entries
    return entries
}

StudyLinks_SaveEntries(category, entries) {
    global g_StudyLinksCache
    if (category != "mnemonics" && category != "plans")
        return false
    iniPath := StudyLinks_GetIniPath()
    section := StudyLinks_SectionName(category)
    try DirCreate(A_ScriptDir "\assets\data")
    catch {
    }
    try IniDelete(iniPath, section)
    catch {
    }
    idx := 0
    for entry in entries {
        idx += 1
        line := entry.name . "|" . entry.url
        try IniWrite(line, iniPath, section, String(idx))
        catch {
        }
    }
    g_StudyLinksCache[category] := entries
    return true
}

StudyLinks_AddEntry(category, name, url) {
    name := Trim(name)
    url := Trim(url)
    if (name = "" || url = "")
        return false
    entries := StudyLinks_GetEntries(category)
    ; Cap: 1-9 + letters excluding reserved a (add) / r (remove).
    if (entries.Length >= 33) {
        try ShowCenteredOverlay_Utils("⚠ Study link list is full (max 33).", 3000, BANNER_ACCENT_INTERMEDIATE)
        return false
    }
    entries.Push({ name: name, url: url })
    return StudyLinks_SaveEntries(category, entries)
}

StudyLinks_RemoveEntry(category, index) {
    entries := StudyLinks_GetEntries(category)
    if (index < 1 || index > entries.Length)
        return false
    entries.RemoveAt(index)
    return StudyLinks_SaveEntries(category, entries)
}

; Item letters skip a (add) and r (remove).
StudyLinks_ItemLetters() {
    return "bcdefghijklmnopqstuvwxyz"
}

StudyLinks_IndexFromKey(key) {
    key := StrLower(Trim(key))
    if (RegExMatch(key, "^[1-9]$"))
        return Integer(key)
    letters := StudyLinks_ItemLetters()
    pos := InStr(letters, key, true)
    if (pos)
        return 9 + pos
    return 0
}

StudyLinks_LabelForIndex(index) {
    if (index >= 1 && index <= 9)
        return String(index)
    letters := StudyLinks_ItemLetters()
    offset := index - 9
    if (offset >= 1 && offset <= StrLen(letters))
        return SubStr(letters, offset, 1)
    return String(index)
}

; Open GitHub study URL in a new Chrome window. scrollToEnd: Mnemonics only (wait load then ^{End}).
StudyTopic_OpenGithubInChrome(url, scrollToEnd := false) {
    url := Trim(url)
    if (url = "") {
        try ShowCenteredOverlay_Utils("⚠ No GitHub URL configured for this topic.", 3000, BANNER_ACCENT_INTERMEDIATE)
        return false
    }

    beforeMap := Map()
    try {
        for hwnd in WinGetList("ahk_exe chrome.exe")
            beforeMap[hwnd] := true
    } catch {
    }

    if (!StudyLink_OpenUrlInChrome(url, true)) {
        try ShowCenteredOverlay_Utils("❌ Could not open Chrome.", 3000, BANNER_ACCENT_ERROR)
        return false
    }

    if (!scrollToEnd)
        return true

    StandardLoadingBar_Show("⏳ Opening study page…", BANNER_ACCENT_INTERMEDIATE, { passive: false })
    try {
        newHwnd := 0
        deadline := A_TickCount + 8000
        while (A_TickCount < deadline) {
            try {
                for hwnd in WinGetList("ahk_exe chrome.exe") {
                    if !beforeMap.Has(hwnd) {
                        newHwnd := hwnd
                        break
                    }
                }
            } catch {
            }
            if (newHwnd)
                break
            Sleep 100
        }

        if (!newHwnd) {
            try ShowCenteredOverlay_Utils("⚠ Chrome window opened but could not scroll to end.", 3000,
                BANNER_ACCENT_INTERMEDIATE)
            return true
        }

        StudyTopic_WaitChromeReadyAndScroll(newHwnd, url)
        return true
    } finally {
        try StandardLoadingBar_Hide(0)
    }
}

; Match expected study URL once Chrome has navigated (ignore query/hash differences).
StudyTopic_UrlMatchFragment(expectedUrl) {
    expectedUrl := Trim(expectedUrl)
    if (expectedUrl = "")
        return ""
    ; Prefer path after host so http/https and www variants still match.
    if RegExMatch(expectedUrl, "i)^https?://[^/]+(/.*)$", &m)
        return m[1]
    return expectedUrl
}

StudyTopic_ChromeGetDocument(uia) {
    if !IsObject(uia)
        return 0
    try
        return uia.GetCurrentDocumentElement()
    catch
        return 0
}

StudyTopic_ChromeContentSampleSize(doc) {
    if !IsObject(doc)
        return 0
    n := 0
    try n := StrLen(Trim(doc.Name))
    catch {
        n := 0
    }
    if (n > 0)
        return n
    try {
        ; Fall back to bounding height as a coarse “content present” signal.
        r := doc.GetPropertyValue(UIA.Property.BoundingRectangle)
        if IsObject(r) && r.Has(3)
            return Integer(r[3])
        if IsObject(r) && r.Length >= 4
            return Integer(r[4] - r[2])
    } catch {
    }
    return 0
}

StudyTopic_ChromeGetScrollVerticalPercent(doc) {
    if !IsObject(doc)
        return -1.0
    try {
        if (doc.GetPropertyValue(UIA.Property.IsScrollPatternAvailable)) {
            p := doc.GetPropertyValue(UIA.Property.ScrollVerticalScrollPercent)
            if (p >= 0)
                return p + 0.0
        }
    } catch {
    }
    return -1.0
}

StudyTopic_ChromeScrollViaUIA(doc) {
    if !IsObject(doc)
        return false
    try {
        if (doc.GetPropertyValue(UIA.Property.IsScrollPatternAvailable)) {
            doc.ScrollPattern.SetScrollPercent(-1, 100)
            return true
        }
    } catch {
    }
    return false
}

StudyTopic_ChromeScrollViaKeystroke(hwnd) {
    if (!hwnd)
        return false
    try {
        WinActivate("ahk_id " hwnd)
        WinWaitActive("ahk_id " hwnd, , 1)
    } catch {
    }
    try {
        ControlSend "{Blind}^{End}", , "ahk_id " hwnd
        return true
    } catch {
        try {
            Send "^{End}"
            return true
        } catch {
            return false
        }
    }
}

; Wait for URL + page load + content stable, then scroll until bottom verified (or 20s budget).
StudyTopic_WaitChromeReadyAndScroll(chromeHwnd, expectedUrl) {
    if (!chromeHwnd)
        return false

    try {
        WinActivate("ahk_id " chromeHwnd)
        WinWaitActive("ahk_id " chromeHwnd, , 3)
    } catch {
    }

    try UIA.ActivateChromiumAccessibility("ahk_id " chromeHwnd, 300)
    catch {
    }

    uia := 0
    try uia := UIA_Browser("ahk_id " chromeHwnd)
    catch {
        uia := 0
    }

    fragment := StudyTopic_UrlMatchFragment(expectedUrl)
    try StandardLoadingBar_Update("⏳ Waiting for page…", BANNER_ACCENT_INTERMEDIATE)

    urlReady := false
    if IsObject(uia) && fragment != "" {
        urlDeadline := A_TickCount + 12000
        while (A_TickCount < urlDeadline) {
            cur := ""
            try cur := uia.GetCurrentURL()
            catch {
                cur := ""
            }
            if (cur != "" && InStr(cur, fragment)) {
                urlReady := true
                break
            }
            Sleep 120
        }
    }

    if IsObject(uia) {
        try uia.WaitPageLoad("", 12000, 500)
        catch {
        }
    } else if (!urlReady) {
        Sleep 1500
    }

    ; Content-stable gate: document exists and sample size steady for two polls.
    try StandardLoadingBar_Update("⏳ Waiting for content…", BANNER_ACCENT_INTERMEDIATE)
    if IsObject(uia) {
        contentDeadline := A_TickCount + 8000
        lastSize := -1
        stableHits := 0
        while (A_TickCount < contentDeadline) {
            doc := StudyTopic_ChromeGetDocument(uia)
            if !IsObject(doc) {
                stableHits := 0
                lastSize := -1
                Sleep 400
                continue
            }
            sz := StudyTopic_ChromeContentSampleSize(doc)
            if (sz > 0 && sz = lastSize) {
                stableHits += 1
                if (stableHits >= 2)
                    break
            } else {
                stableHits := 0
            }
            lastSize := sz
            Sleep 400
        }
    }

    try StandardLoadingBar_Update("⏳ Scrolling to end…", BANNER_ACCENT_INTERMEDIATE)
    try {
        WinActivate("ahk_id " chromeHwnd)
        WinWaitActive("ahk_id " chromeHwnd, , 2)
    } catch {
    }
    ; One focus click into the document (not every loop).
    try {
        WinGetPos(&wx, &wy, &ww, &wh, "ahk_id " chromeHwnd)
        if (ww > 0 && wh > 0)
            Click wx + ww // 2, wy + wh // 2
    } catch {
    }
    Sleep 200

    deadline := A_TickCount + 20000
    lastPct := -2.0
    stableHits := 0
    unknownEndCycles := 0
    sawNearBottom := false
    confirmed := false

    while (A_TickCount < deadline) {
        doc := IsObject(uia) ? StudyTopic_ChromeGetDocument(uia) : 0
        StudyTopic_ChromeScrollViaUIA(doc)
        StudyTopic_ChromeScrollViaKeystroke(chromeHwnd)
        Sleep 400
        doc := IsObject(uia) ? StudyTopic_ChromeGetDocument(uia) : 0
        pct := StudyTopic_ChromeGetScrollVerticalPercent(doc)

        if (pct >= 95.0) {
            confirmed := true
            sawNearBottom := true
            break
        }
        if (pct >= 0) {
            if (pct >= 90.0)
                sawNearBottom := true
            if (Abs(pct - lastPct) < 0.5) {
                stableHits += 1
                if (stableHits >= 2 && pct >= 90.0) {
                    confirmed := true
                    break
                }
            } else {
                stableHits := 0
            }
            lastPct := pct
            unknownEndCycles := 0
        } else {
            ; No scroll % available: count End cycles, then one settle End.
            unknownEndCycles += 1
            if (unknownEndCycles >= 2) {
                Sleep 600
                StudyTopic_ChromeScrollViaKeystroke(chromeHwnd)
                confirmed := true  ; best effort when UIA % missing
                break
            }
        }
    }

    if (!confirmed && !sawNearBottom) {
        try ShowCenteredOverlay_Utils("⚠ Could not confirm scroll to end.", 2500, BANNER_ACCENT_INTERMEDIATE)
    }
    return true
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
    ; Avoid "NA": if focus stays in QuickLook, Esc is consumed there first (ShowAiModelSelector comment).
    gui.Show("x" . cx . " y" . cy)
    try {
        if StudyTopicSelector_GuiHasWindow(gui)
            WinActivate(gui.Hwnd)
    } catch {
    }
}

StudyTopicSelector_StopActiveMonitorTracking() {
    try SetTimer(StudyTopicSelector_TrackActiveMonitorTick, 0)
}

; forMonitorIdx: 1-based index from StudyTopicSelector_TrackActiveMonitorTick (same as GetMonitorIndexForForeground_StandardBar).
; Use Show("AutoSize Hide") then Show("x y") like StudyTopicSelector_PositionGuiLikeOutlook — Move() alone can mis-center across mixed-DPI monitors.
StudyTopicSelector_RepositionToActiveMonitor(forMonitorIdx := 0) {
    global g_StudyTopicSelectorGui, g_StudyTopicSelectorActive
    try {
        if (!StudyTopicSelector_GuiHasWindow(g_StudyTopicSelectorGui))
            return
        idx := forMonitorIdx
        if (idx < 1 || idx > MonitorGetCount())
            idx := GetMonitorIndexForForeground_StandardBar()
        MonitorGetWorkArea(idx, &ml, &mt, &mr, &mb)
        g_StudyTopicSelectorGui.Show("AutoSize Hide")
        g_StudyTopicSelectorGui.GetPos(, , &gw, &gh)
        StudyTopicSelector_ComputeCenterTopLeftInWorkArea(ml, mt, mr, mb, gw, gh, &cx, &cy)
        g_StudyTopicSelectorGui.Show("x" . cx . " y" . cy)
        if (g_StudyTopicSelectorActive && StudyTopicSelector_GuiHasWindow(g_StudyTopicSelectorGui))
            try WinActivate(g_StudyTopicSelectorGui.Hwnd)
    } catch {
        StudyTopicSelector_StopActiveMonitorTracking()
        try StudyTopicSelector_ForceReset()
    }
}

; Follow foreground window's monitor while the selector is open (parity with StandardLoadingBar trackActiveMonitor).
StudyTopicSelector_TrackActiveMonitorTick() {
    global g_StudyTopicSelectorGui, g_StudyTopicSelectorActive, g_StudyTopicSelectorLastForegroundMonitorIdx
    try {
        if (!g_StudyTopicSelectorActive || !StudyTopicSelector_GuiHasWindow(g_StudyTopicSelectorGui)) {
            StudyTopicSelector_StopActiveMonitorTracking()
            if (g_StudyTopicSelectorActive)
                StudyTopicSelector_ForceReset()
            return
        }
        newIdx := GetMonitorIndexForForeground_StandardBar()
        if (newIdx != g_StudyTopicSelectorLastForegroundMonitorIdx) {
            StudyTopicSelector_RepositionToActiveMonitor(newIdx)
            g_StudyTopicSelectorLastForegroundMonitorIdx := newIdx
        }
    } catch {
        StudyTopicSelector_StopActiveMonitorTracking()
        try StudyTopicSelector_ForceReset()
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
    loop 10 {
        try Hotkey(String(A_Index - 1), "Off")
    }
    loop 26 {
        try Hotkey(Chr(96 + A_Index), "Off")
    }
}

; Poll Esc — fallback when $*Escape / g_OnEscapePressed miss (ShowAiModelSelector escape poll).
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
    global g_StudyTopicSelectorActive
    if (!g_StudyTopicSelectorActive)
        return false
    StudyTopicSelector_Cancel()
    return true
}

StudyTopicSelector_GuiEscape(*) {
    StudyTopicSelector_Cancel()
}

; Same escape registration order as ShowAiModelSelector (after Gui.OnEvent registered at build time).
