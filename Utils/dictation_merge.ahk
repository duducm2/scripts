; =============================================================================
; Utils module: dictation_merge.ahk
; Dictation merge non-favorite clips countdown
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; Dictation: Merge non-favorite clips countdown (at end of loop)
; Same UI pattern as clipboard cleanup: 5s banner, N or End to cancel.
; =============================================================================
global g_DictationMergeGui := 0
global g_DictationMergeBorderGui := 0
global g_DictationMergeTextCtrl := 0
global g_DictationMergeRemaining := 0
global g_DictationMergeCanceled := false

DictationMerge_SetCancelHotkeys(enable := true) {
    ; Removed ~ prefix to prevent key leakage into active applications
    if (enable) {
        Hotkey("*n", DictationMerge_Cancel, "On")
        Hotkey("*y", DictationMerge_Proceed, "On")
        Hotkey("*End", DictationMerge_Cancel, "On")
    } else {
        try Hotkey("*n", "Off")
        catch {
        }
        try Hotkey("*y", "Off")
        catch {
        }
        try Hotkey("*End", "Off")
        catch {
        }
    }
}

DictationMerge_ShowBanner() {
    global g_DictationMergeGui, g_DictationMergeTextCtrl, g_DictationMergeRemaining

    try {
        if IsObject(g_DictationMergeGui)
            g_DictationMergeGui.Destroy()
    } catch {
    }
    g_DictationMergeGui := 0
    g_DictationMergeTextCtrl := 0

    target := WinGetID("A")
    hasWindow := false
    if target && WinExist("ahk_id " target) {
        try {
            WinGetPos(&wx, &wy, &ww, &wh, target)
            hasWindow := (ww > 0 && wh > 0)
        } catch {
            hasWindow := false
        }
    }

    ov := Gui("+AlwaysOnTop -Caption +ToolWindow")
    ov.BackColor := "1E1E2E"
    ov.SetFont("s24 cFFFFFF Bold", "Segoe UI")
    g_DictationMergeTextCtrl := ov.Add("Text", "w650 Center", "Merging non-favorite clips in " g_DictationMergeRemaining "... (press Y to proceed, N or End to cancel)"
    )
    ov.Show("AutoSize Hide")
    ov.GetPos(&gx, &gy, &gw, &gh)
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
    try {
        if IsObject(g_DictationMergeBorderGui)
            g_DictationMergeBorderGui.Destroy()
    } catch {
    }
    borderGui := Gui("+AlwaysOnTop -Caption +ToolWindow")
    borderGui.BackColor := BANNER_ACCENT_INTERMEDIATE
    borderGui.Show("NA x" . (cx - borderWidth) . " y" . (cy - borderWidth) . " w" . (gw + 2 * borderWidth) . " h" . (gh +
        2 * borderWidth))
    g_DictationMergeBorderGui := borderGui
    ov.Show("x" . cx . " y" . cy . " NA")
    g_DictationMergeGui := ov
}

DictationMerge_HideBanner() {
    global g_DictationMergeGui, g_DictationMergeBorderGui, g_DictationMergeTextCtrl
    try {
        if IsObject(g_DictationMergeBorderGui)
            g_DictationMergeBorderGui.Destroy()
    } catch {
    }
    g_DictationMergeBorderGui := 0
    try {
        if IsObject(g_DictationMergeGui)
            g_DictationMergeGui.Destroy()
    } catch {
    }
    g_DictationMergeGui := 0
    g_DictationMergeTextCtrl := 0
}

DictationMerge_UpdateBannerText() {
    global g_DictationMergeTextCtrl, g_DictationMergeRemaining
    try {
        if IsObject(g_DictationMergeTextCtrl) {
            g_DictationMergeTextCtrl.Text := "Merging non-favorite clips in " g_DictationMergeRemaining "... (press Y to proceed, N or End to cancel)"
        }
    } catch {
    }
}

DictationMerge_StartCountdown(seconds := 5) {
    global g_DictationMergeRemaining, g_DictationMergeCanceled

    g_DictationMergeCanceled := false
    g_DictationMergeRemaining := seconds

    DictationMerge_ShowBanner()
    DictationMerge_SetCancelHotkeys(true)

    SetTimer(DictationMerge_Tick, 0)
    SetTimer(DictationMerge_Tick, 1000)
}

DictationMerge_StopCountdown(showCancelledBanner := false) {
    SetTimer(DictationMerge_Tick, 0)
    DictationMerge_SetCancelHotkeys(false)

    if (showCancelledBanner) {
        global g_DictationMergeTextCtrl
        try {
            if IsObject(g_DictationMergeTextCtrl) {
                g_DictationMergeTextCtrl.Text := "Merge cancelled"
            }
        } catch {
        }
        SetTimer(DictationMerge_HideBanner, -900)
    } else {
        DictationMerge_HideBanner()
    }
}

DictationMerge_Cancel(*) {
    global g_DictationMergeCanceled
    g_DictationMergeCanceled := true
    DictationMerge_StopCountdown(true)
}

DictationMerge_Proceed(*) {
    ; Immediately proceed with merge, skipping countdown
    DictationMerge_StopCountdown(false)
    MergeNonFavoriteClips()
}

DictationMerge_Tick() {
    global g_DictationMergeRemaining, g_DictationMergeCanceled

    if (g_DictationMergeCanceled) {
        DictationMerge_StopCountdown(true)
        return
    }

    g_DictationMergeRemaining -= 1

    if (g_DictationMergeRemaining <= 0) {
        DictationMerge_StopCountdown(false)
        MergeNonFavoriteClips()
        return
    }

    DictationMerge_UpdateBannerText()
}
