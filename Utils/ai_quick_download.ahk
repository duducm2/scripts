; =============================================================================
; Utils module: ai_quick_download.ahk
; Quick Download: focus AI companion → configured click sequences → Desktop wait → cut newest.
; Trigger: single-tap Win+Alt+Shift+9 (see WindowManagement\audio_bt_menu.ahk).
; Slot chain (scroll / sibling clicks / Desktop wait / rename via #!+P name list / cut)
; lives in click_sequences.ini.
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

; Win+Alt+Shift+9 single-tap entry (called from audio_bt_menu.ahk).
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
    try StandardLoadingBar_Show("⏳ Quick Download…", BANNER_ACCENT_INTERMEDIATE, { passive: false })
    catch {
    }

    focus := AiQuickDownload_FocusCompanion()
    if (!focus.ok) {
        AiQuickDownload_Fail(focus.err)
        return ""
    }

    desktopPath := AiQuickDownload_ResolveDesktopPath()
    beforePath := ""
    beforeStamp := ""
    try {
        beforePath := DesktopCutNewest_ResolveNewestPath(desktopPath)
        if (beforePath != "")
            beforeStamp := AiQuickDownload_ItemStamp(beforePath)
    } catch {
    }

    extras := {
        doCut: doCut,
        seqAttempts: doCut ? 1 : AI_QD_FINANCE_GATE_ATTEMPTS,
        desktopPath: desktopPath,
        beforePath: beforePath,
        beforeStamp: beforeStamp
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

AiQuickDownload_WaitForNewDesktopFile(desktopPath, beforePath, beforeStamp) {
    if (!desktopPath || !DirExist(desktopPath))
        return ""
    start := A_TickCount
    lastPath := ""
    lastSize := -1
    stableCount := 0
    while (A_TickCount - start < AI_QD_DESKTOP_TIMEOUT_MS) {
        newest := ""
        try newest := DesktopCutNewest_ResolveNewestPath(desktopPath)
        catch {
            newest := ""
        }
        if (newest = "") {
            Sleep AI_QD_DESKTOP_POLL_MS
            continue
        }
        SplitPath(newest, &name)
        if (AiQuickDownload_IsTempDownloadName(name)) {
            Sleep AI_QD_DESKTOP_POLL_MS
            continue
        }
        stamp := AiQuickDownload_ItemStamp(newest)
        isNew := (beforePath = "") || (newest != beforePath) || (stamp != "" && beforeStamp != "" && stamp >
            beforeStamp)
        if (!isNew) {
            Sleep AI_QD_DESKTOP_POLL_MS
            continue
        }
        size := -1
        try size := FileGetSize(newest)
        catch {
            size := -1
        }
        if (size < 0) {
            Sleep AI_QD_DESKTOP_POLL_MS
            continue
        }
        if (newest = lastPath && size = lastSize) {
            stableCount += 1
            if (stableCount >= AI_QD_DESKTOP_STABLE_POLLS)
                return newest
        } else {
            lastPath := newest
            lastSize := size
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
