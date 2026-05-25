#Requires AutoHotkey v2.0+

; VS Code / Cursor evidence line -> PDF find loop.
; Preconditions: Tab 1 = text evidence file, Tab 2 = PDF, Ctrl+Tab toggles between them.
; Hotkey: Ctrl+Alt+Win+O — start/stop toggle (bound via EvidenceSearch_BindHotkey).

global g_EvidenceSearchActive := false
global g_EvidenceSearchBannerGui := 0
global g_EvidenceSearchBannerBorderGui := 0

EVIDENCE_ACTION_MS := 450
EVIDENCE_TAB_MS := 700
EVIDENCE_SLEEP_CHUNK_MS := 50

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

EvidenceSearch_FormatString(raw) {
    s := raw
    for q in ['"', "'", Chr(0x60), Chr(0x201C), Chr(0x201D), Chr(0x2018), Chr(0x2019)]
        s := StrReplace(s, q, "")
    return Trim(s)
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

EvidenceSearch_SearchInPdf(formatted) {
    if (!EvidenceSearch_IsActive())
        return false

    SendInput "^f"
    if (!EvidenceSearch_Sleep(EVIDENCE_ACTION_MS))
        return false

    SendInput "^a"
    if (!EvidenceSearch_Sleep(EVIDENCE_ACTION_MS))
        return false

    SendText formatted
    if (!EvidenceSearch_Sleep(EVIDENCE_ACTION_MS))
        return false

    SendInput "{Enter}"
    if (!EvidenceSearch_Sleep(EVIDENCE_TAB_MS))
        return false

    halfLen := Floor(StrLen(formatted) / 2)
    if (halfLen >= 1) {
        half := SubStr(formatted, 1, halfLen)
        SendInput "^a"
        if (!EvidenceSearch_Sleep(EVIDENCE_ACTION_MS))
            return false
        SendText half
        if (!EvidenceSearch_Sleep(EVIDENCE_ACTION_MS))
            return false
        SendInput "{Enter}"
        if (!EvidenceSearch_Sleep(EVIDENCE_TAB_MS))
            return false
    }
    return true
}

EvidenceSearch_CopyCurrentLine() {
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

    try {
        SendMode "Input"
        while (g_EvidenceSearchActive) {
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

            formatted := EvidenceSearch_FormatString(line)

            if (formatted != "") {
                SendInput "^{Tab}"
                if (!EvidenceSearch_Sleep(EVIDENCE_TAB_MS))
                    break
                if (!EvidenceSearch_SearchInPdf(formatted))
                    break
            }

            SendInput "^{Tab}"
            if (!EvidenceSearch_Sleep(EVIDENCE_TAB_MS))
                break

            SendInput "{Down}"
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
    global g_EvidenceSearchActive

    if (g_EvidenceSearchActive) {
        EvidenceSearch_Stop("Evidence search loop stopped")
        return
    }

    if (!EvidenceSearch_IsEditorActive()) {
        EvidenceSearch_NotifyUser("Start failed: focus VS Code or Cursor first", 2600, true)
        return
    }

    g_EvidenceSearchActive := true
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
