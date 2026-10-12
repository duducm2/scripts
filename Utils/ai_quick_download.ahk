; =============================================================================
; Utils module: ai_quick_download.ahk
; Quick Download: name pick first (same list as #!+p [Y]) → focus companion →
; configured click sequences → Desktop wait → rename staging → cut newest.
; PALACE_QUICK_IMAGE: skip cut; leave file on Desktop → Import Management [Q]
; (same auto-import as #!+p [Y] / ClipAngel export).
; Trigger: hold Win+Alt+Shift+9 (see WindowManagement\audio_bt_menu.ahk).
; Slot chain lives in click_sequences.ini; name is chosen before any download.
; =============================================================================

; Tap-dance interval (ms) — matches Teams_R_DoubleTapThresholdMs / ZMK tap-dance.
AI_QD_DOUBLE_TAP_MS := 400
AI_QD_GATE_SETTLE_MS := 600
AI_QD_DESKTOP_POLL_MS := 250
AI_QD_DESKTOP_TIMEOUT_MS := 60000
AI_QD_DESKTOP_STABLE_POLLS := 3
; Finance [D] auto-import: wait after stop-button gone before hunting Download control.
AI_QD_FINANCE_POST_COMPLETE_MS := 2500
; Settle between finance download-control retries.
AI_QD_FINANCE_GATE_SETTLE_MS := 1200
AI_QD_FINANCE_GATE_ATTEMPTS := 4

global g_AiQuickDownloadBusy := false

; Win+Alt+Shift+9 hold entry (called from audio_bt_menu.ahk).
AiQuickDownload_Run() {
    global g_AiQuickDownloadBusy
    if (g_AiQuickDownloadBusy) {
        try ShowCenteredOverlay_Utils("⏳ Quick Download already running…", 1800, BANNER_ACCENT_INTERMEDIATE)
        return
    }
    g_AiQuickDownloadBusy := true
    try {
        AiQuickDownload_RunInner(true)
    } finally {
        g_AiQuickDownloadBusy := false
    }
}

; Finance daily [D]: same download sequences; leave file on Desktop; return path (or "").
AiQuickDownload_RunForFinanceImport() {
    global g_AiQuickDownloadBusy
    if (g_AiQuickDownloadBusy) {
        try ShowCenteredOverlay_Utils("⏳ Quick Download already running…", 1800, BANNER_ACCENT_INTERMEDIATE)
        return ""
    }
    g_AiQuickDownloadBusy := true
    path := ""
    try {
        path := AiQuickDownload_RunInner(false)
    } finally {
        g_AiQuickDownloadBusy := false
    }
    return path
}

; doCut: true = cut newest after wait (#!+9). false = return Desktop path for finance import.
AiQuickDownload_RunInner(doCut := true) {
    global g_ClickSeqRunCtx

    ; Same name-first UX as #!+p [Y]: pick before download so Esc cancels with no companion focus.
    desktopName := ""
    desktopExt := ""
    if (doCut) {
        picked := ""
        ; Empty default ext: Enter keeps name only; rename uses the downloaded file's real extension.
        try picked := ClipAngelExport_PromptPickName("")
        catch {
            picked := ""
        }
        if (!IsObject(picked))
            return ""
        desktopName := picked.name
        desktopExt := picked.ext
    }

    try StandardLoadingBar_Show("⏳ Quick Download…", BANNER_ACCENT_INTERMEDIATE, { passive: false })
    catch {
    }

    focus := AiQuickDownload_FocusCompanion()
    if (!focus.ok) {
        AiQuickDownload_Fail(focus.err)
        return ""
    }

    desktopPath := AiQuickDownload_ResolveDesktopPath()
    ; Full live listing before any click. Newest-by-mtime misses exports whose
    ; modified time is older than something already on the Desktop.
    beforeSnap := AiQuickDownload_SnapshotDesktop(desktopPath)
    beforePath := AiQuickDownload_NewestHintFromSnap(beforeSnap)
    beforeStamp := ""
    if (beforePath != "")
        beforeStamp := AiQuickDownload_ItemStamp(beforePath)

    extras := {
        doCut: doCut,
        seqAttempts: doCut ? 1 : AI_QD_FINANCE_GATE_ATTEMPTS,
        desktopPath: desktopPath,
        beforePath: beforePath,
        beforeStamp: beforeStamp,
        beforeSnap: beforeSnap,
        desktopName: desktopName,
        desktopExt: desktopExt
    }
    ok := false
    try ok := ClickSeq_RunMacro("ai-quick-download", focus.companion, focus.hwnd, extras)
    catch {
        ok := false
    }
    if (!ok) {
        err := ""
        try err := ClickSeq_LastFailMessage()
        catch {
            err := ""
        }
        if (err = "")
            err := "❌ Quick Download: download control not found"
        AiQuickDownload_Fail(err)
        return ""
    }

    newPath := ""
    try {
        if (IsObject(g_ClickSeqRunCtx) && g_ClickSeqRunCtx.HasProp("lastPath"))
            newPath := g_ClickSeqRunCtx.lastPath
    } catch {
        newPath := ""
    }
    if (newPath = "") {
        AiQuickDownload_Fail("❌ Quick Download: file did not appear on Desktop")
        return ""
    }

    if (doCut)
        return newPath

    preferCsv := AiQuickDownload_NewestCsvPreferring(desktopPath, beforePath, beforeStamp, newPath)
    return preferCsv != "" ? preferCsv : newPath
}

AiQuickDownload_Fail(message) {
    try StandardLoadingBar_Hide(0)
    catch {
    }
    try ShowCenteredOverlay_Utils(message, 2800, BANNER_ACCENT_ERROR)
    catch {
    }
}

; One pass of configured click sequences. Caller may retry (finance).
AiQuickDownload_TryClickDownloadControl(hwnd, companion := "") {
    return ClickSeq_RunMacro("ai-quick-download", companion, hwnd)
}

AiQuickDownload_FocusCompanion() {
    result := { ok: false, err: "", hwnd: 0, uia: 0, companion: "" }
    companion := ""
    try companion := ResolveGlobalAICompanion()
    catch {
        companion := "gemini"
    }
    result.companion := companion

    hwnd := 0
    label := "Gemini"
    switch companion {
        case "enterprise":
            try hwnd := GetGeminiEnterpriseWindowHwnd()
            label := "Gemini Enterprise"
        case "copilot":
            try hwnd := GetCopilotWebWindowHwnd()
            label := "Copilot"
        default:
            try hwnd := FindGeminiChromeHwnd()
            label := "Gemini"
    }

    if (!hwnd) {
        result.err := "❌ " . label . " is not open."
        return result
    }

    try WinActivate("ahk_id " hwnd)
    catch {
        result.err := "❌ Could not activate " . label . "."
        return result
    }
    if !WinWaitActive("ahk_id " hwnd, , 2) {
        result.err := "❌ Could not activate " . label . "."
        return result
    }

    uia := 0
    try uia := UIA_Browser("ahk_id " hwnd)
    catch {
        result.err := "❌ UIA attach failed for " . label . "."
        return result
    }
    if (!IsObject(uia)) {
        result.err := "❌ UIA attach failed for " . label . "."
        return result
    }

    result.ok := true
    result.hwnd := hwnd
    result.uia := uia
    return result
}

AiQuickDownload_ResolveDesktopPath() {
    desktopPath := ""
    try desktopPath := GetDesktopToRecyclePath()
    catch
        desktopPath := A_Desktop
    if (!desktopPath || !DirExist(desktopPath))
        desktopPath := A_Desktop
    return desktopPath
}

AiQuickDownload_ItemStamp(path) {
    if (!path)
        return ""
    try {
        tC := FileGetTime(path, "C")
        tM := FileGetTime(path, "M")
        return (tC >= tM) ? tC : tM
    } catch {
        return ""
    }
}

AiQuickDownload_IsTempDownloadName(name) {
    lower := StrLower(name)
    if (lower = "desktop.ini")
        return true
    if (InStr(lower, ".crdownload"))
        return true
    if (InStr(lower, ".tmp"))
        return true
    if (SubStr(lower, -7) = ".partial")
        return true
    return false
}

AiQuickDownload_ReadFileStat(path) {
    ctime := ""
    mtime := ""
    size := -1
    try ctime := FileGetTime(path, "C")
    catch {
        ctime := ""
    }
    try mtime := FileGetTime(path, "M")
    catch {
        mtime := ""
    }
    try size := FileGetSize(path)
    catch {
        size := -1
    }
    return { size: size, ctime: ctime, mtime: mtime }
}

; Live listing. Does not use the folder-mtime index: a download can arrive
; without updating the folder timestamp, and with a modified time older than
; files already on the Desktop.
AiQuickDownload_SnapshotDesktop(desktopPath) {
    snap := Map()
    if (!desktopPath || !DirExist(desktopPath))
        return snap
    loop files desktopPath "\*", "F" {
        if (StrLower(A_LoopFileName) = "desktop.ini")
            continue
        stat := AiQuickDownload_ReadFileStat(A_LoopFileFullPath)
        snap[StrLower(A_LoopFileFullPath)] := {
            path: A_LoopFileFullPath,
            size: stat.size,
            ctime: stat.ctime,
            mtime: stat.mtime
        }
    }
    return snap
}

; Finance CSV hint only. Detection uses the full snapshot, not this path.
AiQuickDownload_NewestHintFromSnap(beforeSnap) {
    bestPath := ""
    bestStamp := ""
    if !(beforeSnap is Map)
        return ""
    for snapKey, row in beforeSnap {
        if (!IsObject(row) || !row.HasProp("path"))
            continue
        SplitPath(row.path, &name)
        if (AiQuickDownload_IsTempDownloadName(name))
            continue
        stamp := ""
        if (row.ctime != "" && row.mtime != "")
            stamp := (row.ctime >= row.mtime) ? row.ctime : row.mtime
        else
            stamp := row.ctime != "" ? row.ctime : row.mtime
        if (bestPath = "" || (stamp != "" && stamp > bestStamp)) {
            bestPath := row.path
            bestStamp := stamp
        }
    }
    return bestPath
}

AiQuickDownload_CtimeIsLater(ctime, bestCtime) {
    if (bestCtime = "")
        return true
    if (ctime = "")
        return false
    return ctime > bestCtime
}

; status idle: no folder change. pending: temp download or locked candidate.
; ready: non-temp file that is new, or whose size or creation time changed.
AiQuickDownload_ProbeDesktop(desktopPath, beforeSnap) {
    result := { status: "idle", path: "", size: -1, ctime: "" }
    if (!desktopPath || !DirExist(desktopPath))
        return result
    if !(beforeSnap is Map)
        beforeSnap := Map()

    bestPath := ""
    bestCtime := ""
    bestSize := -1
    pending := false

    loop files desktopPath "\*", "F" {
        if (StrLower(A_LoopFileName) = "desktop.ini")
            continue
        full := A_LoopFileFullPath
        key := StrLower(full)
        stat := AiQuickDownload_ReadFileStat(full)
        inBase := beforeSnap.Has(key)
        prior := inBase ? beforeSnap[key] : ""

        if (AiQuickDownload_IsTempDownloadName(A_LoopFileName)) {
            if (!inBase)
                pending := true
            else if (stat.size >= 0 && IsObject(prior) && prior.size >= 0 && stat.size != prior.size)
                pending := true
            else if (stat.ctime != "" && IsObject(prior) && prior.ctime != "" && stat.ctime != prior.ctime)
                pending := true
            continue
        }

        changed := !inBase
        if (inBase && IsObject(prior)) {
            if (stat.size >= 0 && prior.size >= 0 && stat.size != prior.size)
                changed := true
            else if (stat.ctime != "" && prior.ctime != "" && stat.ctime != prior.ctime)
                changed := true
        }
        if (!changed)
            continue
        if (stat.size < 0) {
            pending := true
            continue
        }
        if (bestPath = "" || AiQuickDownload_CtimeIsLater(stat.ctime, bestCtime)) {
            bestPath := full
            bestCtime := stat.ctime
            bestSize := stat.size
        }
    }

    if (bestPath != "") {
        result.status := "ready"
        result.path := bestPath
        result.size := bestSize
        result.ctime := bestCtime
        return result
    }
    if (pending)
        result.status := "pending"
    return result
}

AiQuickDownload_SetWaitBanner(which, &banner) {
    if (which = banner)
        return
    banner := which
    msg := (which = "progress") ? "⏳ Download in progress…" : "⏳ Waiting for Desktop file…"
    try StandardLoadingBar_Update(msg, BANNER_ACCENT_INTERMEDIATE)
    catch {
    }
}

; beforeSnap: Map from AiQuickDownload_SnapshotDesktop, taken before the download.
AiQuickDownload_WaitForNewDesktopFile(desktopPath, beforeSnap) {
    if (!desktopPath || !DirExist(desktopPath))
        return ""
    if !(beforeSnap is Map)
        return ""
    start := A_TickCount
    lastPath := ""
    lastSize := -1
    stableCount := 0
    banner := ""
    while (A_TickCount - start < AI_QD_DESKTOP_TIMEOUT_MS) {
        probe := AiQuickDownload_ProbeDesktop(desktopPath, beforeSnap)
        if (probe.status = "pending") {
            AiQuickDownload_SetWaitBanner("progress", &banner)
            lastPath := ""
            lastSize := -1
            stableCount := 0
            Sleep AI_QD_DESKTOP_POLL_MS
            continue
        }
        if (probe.status != "ready" || probe.path = "") {
            AiQuickDownload_SetWaitBanner("waiting", &banner)
            lastPath := ""
            lastSize := -1
            stableCount := 0
            Sleep AI_QD_DESKTOP_POLL_MS
            continue
        }
        if (probe.path = lastPath && probe.size = lastSize) {
            stableCount += 1
            if (stableCount >= AI_QD_DESKTOP_STABLE_POLLS)
                return probe.path
            AiQuickDownload_SetWaitBanner("waiting", &banner)
        } else {
            if (lastPath != "")
                AiQuickDownload_SetWaitBanner("progress", &banner)
            else
                AiQuickDownload_SetWaitBanner("waiting", &banner)
            lastPath := probe.path
            lastSize := probe.size
            stableCount := 1
        }
        Sleep AI_QD_DESKTOP_POLL_MS
    }
    return ""
}

; If a new .csv exists that is at least as new as candidate, prefer it for finance import.
AiQuickDownload_NewestCsvPreferring(desktopPath, beforePath, beforeStamp, candidate) {
    best := ""
    bestStamp := ""
    loop files desktopPath . "\*.csv", "F" {
        if (AiQuickDownload_IsTempDownloadName(A_LoopFileName))
            continue
        stamp := AiQuickDownload_ItemStamp(A_LoopFileFullPath)
        isNew := (beforePath = "") || (A_LoopFileFullPath != beforePath) || (stamp != "" && beforeStamp != "" &&
            stamp > beforeStamp)
        if (!isNew)
            continue
        if (best = "" || (stamp != "" && stamp >= bestStamp)) {
            best := A_LoopFileFullPath
            bestStamp := stamp
        }
    }
    if (best != "")
        return best
    return candidate
}

; Scroll chat feed to bottom so newest download controls have valid geometry.
; Use keyboard/wheel fallback only — never ChromeChat_ScrollFeedToBottomFast /
; UIA JSExecute (SetURL "javascript:…"), which can type the payload into the
; focused Gemini/Copilot composer when the omnibox is not actually targeted.
AiQuickDownload_ScrollFeedToBottom(hwnd, companion := "") {
    if (!hwnd)
        return
    try ChromeChat_ScrollFeedToBottomFallback(hwnd)
    catch {
    }
}
