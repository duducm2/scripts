#Requires AutoHotkey v2.0+

; VS Code / Cursor CSV row -> PDF substring search loop.
; Read row via ^l ^c (read-only); extract from char 3, length Floor(len/4); search PDF; Escape; return; Down.
; Preconditions: Tab 1 = CSV, Tab 2 = PDF. Switch via Ctrl+Tab + Enter (picker confirm); fallback Ctrl+1/Ctrl+2 + Enter.
; Hotkey: Ctrl+Alt+Win+O — start/stop toggle (bound via EvidenceSearch_BindHotkey).

global g_EvidenceSearchActive := false
global g_EvidenceSearchBannerGui := 0
global g_EvidenceSearchBannerBorderGui := 0
global g_EvidenceSearchNotFoundRows := []
global EVIDENCE_NOTFOUND_REPORT_MS := 10000

EVIDENCE_ACTION_MS := 600
EVIDENCE_TAB_MS := 900
EVIDENCE_BEFORE_ENTER_MS := 1000
EVIDENCE_TITLE_POLL_MS := 100
EVIDENCE_TITLE_WAIT_MS := 3000
EVIDENCE_SLEEP_CHUNK_MS := 50

EvidenceSearch_GetActiveTitle() {
    try
        return WinGetTitle("A")
    return ""
}

EvidenceSearch_TitleLooksCsv(title) {
    return InStr(title, ".csv", false) > 0
}

EvidenceSearch_TitleLooksPdf(title) {
    return InStr(title, ".pdf", false) > 0
}

; UIA: PDF find bar shows Type 50020 Text "Phrase not found" when search misses.
EvidenceSearch_PdfPhraseNotFound() {
    try {
        hwnd := WinExist("A")
        if (!hwnd)
            return false
        root := UIA.ElementFromHandle(hwnd)
        return root.FindFirst({ Name: "Phrase not found", Type: 50020, cs: false }) != ""
    } catch {
        return false
    }
}

EvidenceSearch_IsEditorActive() {
    return WinActive("ahk_exe Code.exe") || WinActive("ahk_exe Cursor.exe")
}

EvidenceSearch_NotifyUser(msg, durationMs := 2200, isError := false) {
    try {
        ShowCenteredOverlay_Utils(msg, durationMs, isError ? BANNER_ACCENT_ERROR : BANNER_ACCENT_SUCCESS)
        return
    } catch {
    }
    ToolTip msg
    SetTimer(() => ToolTip(), -durationMs)
}

EvidenceSearch_Sleep(ms) {
    global g_EvidenceSearchActive
    while (ms > 0 && g_EvidenceSearchActive) {
        chunk := Min(EVIDENCE_SLEEP_CHUNK_MS, ms)
        Sleep chunk
        ms -= chunk
    }
    return g_EvidenceSearchActive
}

EvidenceSearch_IsActive() {
    global g_EvidenceSearchActive
    return g_EvidenceSearchActive
}

; Poll window title until CSV or PDF tab is active (avoids Enter while still on wrong tab).
EvidenceSearch_WaitForTitle(wantCsv, maxMs := EVIDENCE_TITLE_WAIT_MS) {
    deadline := A_TickCount + maxMs
    while (A_TickCount < deadline) {
        if (!EvidenceSearch_IsActive())
            return false
        title := EvidenceSearch_GetActiveTitle()
        if (wantCsv && EvidenceSearch_TitleLooksCsv(title))
            return true
        if (!wantCsv && EvidenceSearch_TitleLooksPdf(title))
            return true
        if (!EvidenceSearch_Sleep(EVIDENCE_TITLE_POLL_MS))
            return false
    }
    title := EvidenceSearch_GetActiveTitle()
    return wantCsv ? EvidenceSearch_TitleLooksCsv(title) : EvidenceSearch_TitleLooksPdf(title)
}

; Char 3 onward, length Floor(25% of trimmed row length). Returns "" if termLen < 1 or len < 3.
EvidenceSearch_TruncateForDisplay(s, maxLen := 55) {
    s := Trim(s)
    if (StrLen(s) <= maxLen)
        return s
    return SubStr(s, 1, maxLen - 1) . "…"
}

EvidenceSearch_ResetNotFoundRows() {
    global g_EvidenceSearchNotFoundRows
    g_EvidenceSearchNotFoundRows := []
}

EvidenceSearch_RecordNotFound(rowNum, line, term) {
    global g_EvidenceSearchNotFoundRows
    preview := EvidenceSearch_TruncateForDisplay(line)
    g_EvidenceSearchNotFoundRows.Push("Row " . rowNum . ": " . preview . "  (term: " . term . ")")
}

EvidenceSearch_ShowNotFoundReport() {
    global g_EvidenceSearchNotFoundRows
    n := g_EvidenceSearchNotFoundRows.Length
    if (n = 0)
        return

    body := "⚠ " . n . " phrase(s) not found in PDF:`n`n"
    for entry in g_EvidenceSearchNotFoundRows
        body .= "• " . entry . "`n"
    body .= "`n📸 Take a print screen now (Alt+PrintScreen)"

    try {
        StandardLoadingBar_Show(body, BANNER_ACCENT_ERROR, { passive: true, centerOnHwnd: 0, textWidth: 720, fontSize: 17,
            passiveBgColor: BANNER_ACCENT_ERROR })
        StandardLoadingBar_Hide(EVIDENCE_NOTFOUND_REPORT_MS)
    } catch {
        EvidenceSearch_NotifyUser(body, EVIDENCE_NOTFOUND_REPORT_MS, true)
    }
}

EvidenceSearch_ExtractSearchTerm(original) {
    s := Trim(original)
    len := StrLen(s)
    if (len < 3)
        return ""
    termLen := Floor(len / 4)
    if (termLen < 1)
        return ""
    return SubStr(s, 3, termLen)
}

EvidenceSearch_ShowBanner(text := "Evidence search loop (Ctrl+Alt+Win+O to stop)") {
    global g_EvidenceSearchBannerGui, g_EvidenceSearchBannerBorderGui
    EvidenceSearch_HideBanner()

    target := WinExist("A")
    hasWindow := false
    if (target && WinExist("ahk_id " target)) {
        try {
            WinGetPos(&wx, &wy, &ww, &wh, target)
            hasWindow := (ww > 0 && wh > 0)
        } catch {
            hasWindow := false
        }
    }

    ov := Gui("+AlwaysOnTop -Caption +ToolWindow")
    ov.BackColor := "1E1E2E"
    ov.SetFont("s18 cFFFFFF Bold", "Segoe UI")
    ov.Add("Text", "w560 Center", text)
    ov.Show("AutoSize Hide")
    ov.GetPos(, , &gw, &gh)

    if (hasWindow) {
        cx := wx + (ww - gw) // 2
        cy := wy + (wh - gh) // 2
    } else {
        vx := SysGet(76)
        vy := SysGet(77)
        vw := SysGet(78)
        vh := SysGet(79)
        cx := vx + (vw - gw) // 2
        cy := vy + (vh - gh) // 2
    }

    borderWidth := 6
    borderGui := Gui("+AlwaysOnTop -Caption +ToolWindow")
    borderGui.BackColor := "0078D4"
    borderGui.Show("NA x" . (cx - borderWidth) . " y" . (cy - borderWidth) . " w" . (gw + 2 * borderWidth) . " h" . (gh +
        2 * borderWidth))
    g_EvidenceSearchBannerBorderGui := borderGui

    ov.Show("x" . cx . " y" . cy . " NA")
    WinSetTransparent(178, ov)
    g_EvidenceSearchBannerGui := ov
}

EvidenceSearch_HideBanner() {
    global g_EvidenceSearchBannerGui, g_EvidenceSearchBannerBorderGui
    try {
        if (IsObject(g_EvidenceSearchBannerBorderGui))
            g_EvidenceSearchBannerBorderGui.Destroy()
    } catch {
    }
    g_EvidenceSearchBannerBorderGui := 0
    try {
        if (IsObject(g_EvidenceSearchBannerGui))
            g_EvidenceSearchBannerGui.Destroy()
    } catch {
    }
    g_EvidenceSearchBannerGui := 0
}

EvidenceSearch_Stop(reason := "", isError := false) {
    global g_EvidenceSearchActive
    g_EvidenceSearchActive := false
    EvidenceSearch_HideBanner()
    if (reason != "")
        EvidenceSearch_NotifyUser(reason, 2200, isError)
}

; Focus CSV or PDF tab; verify via window title (.csv / .pdf).
; VS Code Ctrl+Tab opens the editor picker — wait, then Enter confirms the tab.
EvidenceSearch_ConfirmEditorPicker(wantCsv) {
    if (!EvidenceSearch_Sleep(EVIDENCE_BEFORE_ENTER_MS))
        return false
    SendInput "{Enter}"
    if (!EvidenceSearch_Sleep(EVIDENCE_TAB_MS))
        return false
    return EvidenceSearch_WaitForTitle(wantCsv)
}

EvidenceSearch_FocusEditorTab(wantCsv) {
    title := EvidenceSearch_GetActiveTitle()
    if (wantCsv && EvidenceSearch_TitleLooksCsv(title))
        return true
    if (!wantCsv && EvidenceSearch_TitleLooksPdf(title))
        return true

    SendInput "^{Tab}"
    if (!EvidenceSearch_Sleep(EVIDENCE_ACTION_MS))
        return false
    if (!EvidenceSearch_ConfirmEditorPicker(wantCsv))
        return false

    titleAfter := EvidenceSearch_GetActiveTitle()
    ok := wantCsv ? EvidenceSearch_TitleLooksCsv(titleAfter) : EvidenceSearch_TitleLooksPdf(titleAfter)

    if (!ok) {
        if (wantCsv)
            SendInput "^1"
        else
            SendInput "^2"
        if (!EvidenceSearch_Sleep(EVIDENCE_ACTION_MS))
            return false
        if (!EvidenceSearch_ConfirmEditorPicker(wantCsv))
            return false
        titleAfter := EvidenceSearch_GetActiveTitle()
        ok := wantCsv ? EvidenceSearch_TitleLooksCsv(titleAfter) : EvidenceSearch_TitleLooksPdf(titleAfter)
    }

    return ok
}

EvidenceSearch_SwitchToPdfTab() {
    return EvidenceSearch_FocusEditorTab(false)
}

EvidenceSearch_SwitchToCsvTab() {
    return EvidenceSearch_FocusEditorTab(true)
}

EvidenceSearch_SearchSubstringInPdf(term, rowNum := 0, line := "") {
    if (!EvidenceSearch_IsActive())
        return false

    if (!EvidenceSearch_FocusEditorTab(false))
        return false
    if (!EvidenceSearch_WaitForTitle(false))
        return false

    SendInput "^f"
    if (!EvidenceSearch_Sleep(EVIDENCE_ACTION_MS))
        return false

    SendInput "^a"
    if (!EvidenceSearch_Sleep(EVIDENCE_ACTION_MS))
        return false

    if (EvidenceSearch_TitleLooksCsv(EvidenceSearch_GetActiveTitle())) {
        SendInput "{Escape}"
        EvidenceSearch_Sleep(EVIDENCE_ACTION_MS)
        return false
    }

    SendText term
    if (!EvidenceSearch_Sleep(EVIDENCE_ACTION_MS))
        return false

    if (!EvidenceSearch_Sleep(EVIDENCE_BEFORE_ENTER_MS))
        return false
    if (EvidenceSearch_TitleLooksCsv(EvidenceSearch_GetActiveTitle())) {
        SendInput "{Escape}"
        EvidenceSearch_Sleep(EVIDENCE_ACTION_MS)
        return false
    }

    SendInput "{Enter}"
    if (!EvidenceSearch_Sleep(EVIDENCE_TAB_MS))
        return false

    if (EvidenceSearch_PdfPhraseNotFound()) {
        if (rowNum > 0)
            EvidenceSearch_RecordNotFound(rowNum, line, term)
    }

    SendInput "{Escape}"
    if (!EvidenceSearch_Sleep(EVIDENCE_ACTION_MS))
        return false

    return true
}

EvidenceSearch_CopyCurrentLine() {
    if (!EvidenceSearch_FocusEditorTab(true))
        return ""
    SendInput "^l"
    if (!EvidenceSearch_Sleep(EVIDENCE_ACTION_MS))
        return ""
    SendInput "^c"
    if (!EvidenceSearch_Sleep(EVIDENCE_ACTION_MS))
        return ""
    ClipWait(1.5)
    return Trim(A_Clipboard, "`r`n `t")
}

EvidenceSearch_RunLoop() {
    global g_EvidenceSearchActive
    savedClip := A_Clipboard
    ClipSaved := true
    rowNum := 0

    try {
        SendMode "Input"
        while (g_EvidenceSearchActive) {
            rowNum += 1
            if (!EvidenceSearch_IsEditorActive()) {
                EvidenceSearch_Stop("Evidence loop stopped: VS Code / Cursor not active", true)
                return
            }

            line := EvidenceSearch_CopyCurrentLine()
            if (!g_EvidenceSearchActive)
                break
            if (line = "") {
                EvidenceSearch_Stop("Evidence search finished (empty line)")
                return
            }

            term := EvidenceSearch_ExtractSearchTerm(line)

            if (term != "") {
                if (!EvidenceSearch_SwitchToPdfTab())
                    break
                if (!EvidenceSearch_SearchSubstringInPdf(term, rowNum, line))
                    break
                if (!EvidenceSearch_SwitchToCsvTab())
                    break
            }

            if (!EvidenceSearch_FocusEditorTab(true))
                break
            SendInput "{Down}"
            Sleep 100
            SendInput "{Up}"
            if (!EvidenceSearch_Sleep(EVIDENCE_ACTION_MS))
                break
        }
    } finally {
        if (ClipSaved)
            A_Clipboard := savedClip
        if (g_EvidenceSearchActive) {
            g_EvidenceSearchActive := false
            EvidenceSearch_HideBanner()
        }
    }
}

EvidenceSearch_Toggle(*) {
    global g_EvidenceSearchActive, g_EvidenceSearchNotFoundRows

    if (g_EvidenceSearchActive) {
        if (g_EvidenceSearchNotFoundRows.Length > 0) {
            EvidenceSearch_Stop("", false)
            EvidenceSearch_ShowNotFoundReport()
        } else
            EvidenceSearch_Stop("Evidence search loop stopped")
        return
    }

    if (!EvidenceSearch_IsEditorActive()) {
        EvidenceSearch_NotifyUser("Start failed: focus VS Code or Cursor first", 2600, true)
        return
    }

    EvidenceSearch_ResetNotFoundRows()
    g_EvidenceSearchActive := true
    if (!EvidenceSearch_FocusEditorTab(true)) {
        g_EvidenceSearchActive := false
        EvidenceSearch_NotifyUser("Start failed: focus CSV tab (Ctrl+1) first", 2600, true)
        return
    }
    EvidenceSearch_ShowBanner()
    EvidenceSearch_NotifyUser("Evidence search loop started", 1400)
    EvidenceSearch_RunLoop()
}

EvidenceSearch_BindHotkey() {
    #MaxThreadsPerHotkey 2
    try Hotkey("^!#o", EvidenceSearch_Toggle.Bind(), "On")
    catch Error as e {
        try EvidenceSearch_NotifyUser("Could not bind Ctrl+Alt+Win+O: " . e.Message, 3000, true)
    }
}

; When run standalone (not #included from Shift keys.ahk), bind hotkey here.
if (!IsSet(EVIDENCE_SEARCH_FROM_SHIFT_KEYS)) {
    #SingleInstance Force
    EvidenceSearch_BindHotkey()
}
