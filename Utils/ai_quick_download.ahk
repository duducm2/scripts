; =============================================================================
; Utils module: ai_quick_download.ahk
; Quick Download: focus AI companion → Quality Gates → Desktop wait → cut newest.
; Trigger: double-tap Win+Alt+Shift+9 (see WindowManagement\audio_bt_menu.ahk).
; =============================================================================

; Tap-dance interval (ms) — matches Teams_R_DoubleTapThresholdMs / ZMK tap-dance.
AI_QD_DOUBLE_TAP_MS := 400
AI_QD_GATE_SETTLE_MS := 600
AI_QD_OPEN_SETTLE_MS := 800
AI_QD_DESKTOP_POLL_MS := 250
AI_QD_DESKTOP_TIMEOUT_MS := 60000
AI_QD_DESKTOP_STABLE_POLLS := 3

global g_AiQuickDownloadBusy := false

; Win+Alt+Shift+9 double-tap entry (called from audio_bt_menu.ahk).
AiQuickDownload_Run() {
    global g_AiQuickDownloadBusy
    if (g_AiQuickDownloadBusy) {
        try ShowCenteredOverlay_Utils("⏳ Quick Download already running…", 1800, BANNER_ACCENT_INTERMEDIATE)
        return
    }
    g_AiQuickDownloadBusy := true
    try {
        AiQuickDownload_RunInner()
    } finally {
        g_AiQuickDownloadBusy := false
    }
}

AiQuickDownload_RunInner() {
    try StandardLoadingBar_Show("⏳ Quick Download…", BANNER_ACCENT_INTERMEDIATE, { passive: false })
    catch {
    }

    focus := AiQuickDownload_FocusCompanion()
    if (!focus.ok) {
        AiQuickDownload_Fail(focus.err)
        return
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

    try StandardLoadingBar_Update("⏳ Finding download control…", BANNER_ACCENT_INTERMEDIATE)
    catch {
    }

    uia := focus.uia
    hwnd := focus.hwnd
    clicked := false
    ; Five Quality Gates in order (C–E stubbed until Enterprise/Copilot dumps).
    try clicked := AiQuickDownload_GateA_DirectDownload(uia, hwnd)
    catch {
        clicked := false
    }
    if (!clicked) {
        try uia := UIA_Browser("ahk_id " hwnd)
        catch {
        }
        try clicked := AiQuickDownload_GateB_OpenThenDownload(uia, hwnd)
        catch {
            clicked := false
        }
    }
    if (!clicked) {
        try uia := UIA_Browser("ahk_id " hwnd)
        catch {
        }
        try clicked := AiQuickDownload_GateC_Enterprise(uia, hwnd)
        catch {
            clicked := false
        }
    }
    if (!clicked) {
        try uia := UIA_Browser("ahk_id " hwnd)
        catch {
        }
        try clicked := AiQuickDownload_GateD_CopilotCard(uia, hwnd)
        catch {
            clicked := false
        }
    }
    if (!clicked) {
        try uia := UIA_Browser("ahk_id " hwnd)
        catch {
        }
        try clicked := AiQuickDownload_GateE_CopilotPreview(uia, hwnd)
        catch {
            clicked := false
        }
    }

    if (!clicked) {
        AiQuickDownload_Fail("❌ Quick Download: download control not found")
        return
    }

    try StandardLoadingBar_Update("⏳ Waiting for Desktop file…", BANNER_ACCENT_INTERMEDIATE)
    catch {
    }

    if !AiQuickDownload_WaitForNewDesktopFile(desktopPath, beforePath, beforeStamp) {
        AiQuickDownload_Fail("❌ Quick Download: file did not appear on Desktop")
        return
    }

    try StandardLoadingBar_Hide(0)
    catch {
    }
    try DesktopCutNewest_Trigger()
    catch as e {
        try ShowCenteredOverlay_Utils("❌ Quick Download: cut failed — " SubStr(e.Message, 1, 60), 2500,
        BANNER_ACCENT_ERROR)
        catch {
        }
    }
}

AiQuickDownload_Fail(message) {
    try StandardLoadingBar_Hide(0)
    catch {
    }
    try ShowCenteredOverlay_Utils(message, 2800, BANNER_ACCENT_ERROR)
    catch {
    }
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
        return false
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
                return true
        } else {
            lastPath := newest
            lastSize := size
            stableCount := 1
        }
        Sleep AI_QD_DESKTOP_POLL_MS
    }
    return false
}

AiQuickDownload_Invoke(el) {
    if (!IsObject(el))
        return false
    try {
        if el.GetPropertyValue(UIA.Property.IsInvokePatternAvailable) {
            el.InvokePattern.Invoke()
            return true
        }
    } catch {
    }
    try {
        el.Invoke()
        return true
    } catch {
    }
    try {
        el.Click()
        return true
    } catch {
    }
    return false
}

; Prefer last matching button (document order ≈ newest response).
AiQuickDownload_FindLastButton(uia, predicate) {
    if (!IsObject(uia))
        return 0
    ; Accept Func objects or bare function references.
    pred := predicate
    if (Type(predicate) = "String") {
        try pred := Func(predicate)
        catch {
            return 0
        }
    }
    if (Type(pred) != "Func" && !HasMethod(pred, "Call"))
        return 0
    matches := []
    try {
        allButtons := uia.FindAll({ Type: 50000 })
        for btn in allButtons {
            try {
                if (pred.Call(btn))
                    matches.Push(btn)
            } catch {
            }
        }
    } catch {
        try {
            allButtons := uia.FindAll({ Type: "Button" })
            for btn in allButtons {
                try {
                    if (pred.Call(btn))
                        matches.Push(btn)
                } catch {
                }
            }
        } catch {
            return 0
        }
    }
    if (matches.Length = 0)
        return 0
    return matches[matches.Length]
}

AiQuickDownload_ClassContains(el, needle) {
    if (!IsObject(el) || needle = "")
        return false
    try {
        cn := el.ClassName
        return (cn != "" && InStr(cn, needle))
    } catch {
        return false
    }
}

AiQuickDownload_NameIs(el, names) {
    if (!IsObject(el))
        return false
    n := ""
    try n := el.Name
    catch {
        return false
    }
    if (n = "")
        return false
    for want in names {
        if (n = want)
            return true
    }
    return false
}

; --- Gate A: Direct Download (Personal Gemini Drive viewer open) ---------------
AiQuickDownload_IsDriveViewerDownload(el) {
    return AiQuickDownload_NameIs(el, ["Download", "Baixar"])
    && AiQuickDownload_ClassContains(el, "drive-viewer")
}

AiQuickDownload_IsDownloadOrBaixar(el) {
    return AiQuickDownload_NameIs(el, ["Download", "Baixar"])
}

AiQuickDownload_GateA_DirectDownload(uia, hwnd := 0) {
    btn := AiQuickDownload_FindLastButton(uia, AiQuickDownload_IsDriveViewerDownload)
    if (!IsObject(btn)) {
        ; Broader: last Download/Baixar button (viewer may omit class in some builds).
        btn := AiQuickDownload_FindLastButton(uia, AiQuickDownload_IsDownloadOrBaixar)
    }
    if (!IsObject(btn))
        return false
    if !AiQuickDownload_Invoke(btn)
        return false
    Sleep AI_QD_GATE_SETTLE_MS
    return true
}

; --- Gate B: Open file chip → viewer Download (Personal Gemini closed) --------
AiQuickDownload_IsOpenChipButton(el) {
    return AiQuickDownload_NameIs(el, ["Open", "Abrir"])
    && AiQuickDownload_ClassContains(el, "open-button")
}

AiQuickDownload_IsOpenNamed(el) {
    return AiQuickDownload_NameIs(el, ["Open"])
}

AiQuickDownload_IsDownloadCodeNamed(el) {
    return AiQuickDownload_NameIs(el, ["Download code", "Baixar código", "Baixar codigo"])
}

AiQuickDownload_GateB_OpenThenDownload(uia, hwnd := 0) {
    openBtn := AiQuickDownload_FindLastButton(uia, AiQuickDownload_IsOpenChipButton)
    if (!IsObject(openBtn))
        openBtn := AiQuickDownload_FindLastButton(uia, AiQuickDownload_IsOpenNamed)
    if (IsObject(openBtn)) {
        if AiQuickDownload_Invoke(openBtn) {
            Sleep AI_QD_OPEN_SETTLE_MS
            try {
                if (hwnd)
                    uia := UIA_Browser("ahk_id " hwnd)
            } catch {
            }
            if AiQuickDownload_GateA_DirectDownload(uia, hwnd)
                return true
        }
    }

    ; Fallback: code-block "Download code" under download-button group.
    codeBtn := AiQuickDownload_FindLastButton(uia, AiQuickDownload_IsDownloadCodeNamed)
    if (!IsObject(codeBtn))
        codeBtn := AiQuickDownload_FindLastButton(uia, AiQuickDownload_IsDownloadCodeCandidate)
    if (!IsObject(codeBtn))
        return false
    if !AiQuickDownload_Invoke(codeBtn)
        return false
    Sleep AI_QD_GATE_SETTLE_MS
    return true
}

AiQuickDownload_ParentClassContains(el, needle) {
    try {
        p := el.WalkTree("p")
        if (IsObject(p) && AiQuickDownload_ClassContains(p, needle))
            return true
    } catch {
    }
    return false
}

AiQuickDownload_IsDownloadCodeCandidate(el) {
    if (AiQuickDownload_NameIs(el, ["Download code", "Baixar código", "Baixar codigo"]))
        return true
    n := ""
    try n := el.Name
    catch {
        return false
    }
    if (n = "" || !InStr(n, "Download"))
        return false
    return AiQuickDownload_ClassContains(el, "mdc-icon-button")
    && AiQuickDownload_ParentClassContains(el, "download-button")
}

; --- Gate C: Gemini Enterprise (stub until UIA dump) --------------------------
; Predictive targets when dumps arrive:
;   - Canvas: Button/MenuItem Name "Export" → DOCX / PDF
;   - Library media: Button Name "Download" / "Baixar"
;   - File chip Download analogous to Personal Gate A/B
AiQuickDownload_GateC_Enterprise(uia, hwnd := 0) {
    return false
}

; --- Gate D: Copilot Web file card / blob Download (stub until UIA dump) ------
; Predictive targets when dumps arrive:
;   - Chat file card Button Name "Download" / "Baixar"
;   - Direct blob download control before OneDrive redirect
AiQuickDownload_GateD_CopilotCard(uia, hwnd := 0) {
    return false
}

; --- Gate E: Copilot / Office preview fallback (stub until UIA dump) ----------
; Predictive targets when dumps arrive:
;   - "Download a copy", "Download", File-menu download in Word Online / preview
AiQuickDownload_GateE_CopilotPreview(uia, hwnd := 0) {
    return false
}
