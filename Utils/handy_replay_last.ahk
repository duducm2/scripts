; =============================================================================
; Utils module: handy_replay_last.ahk
; Send dictation? [R] — open Handy, History tab, play last recording
; =============================================================================

HandyReplay_FindNamed(el, typeId, names) {
    if !el
        return 0
    for nm in names {
        try {
            found := el.FindFirst({ Type: typeId, Name: nm })
            if found
                return found
        } catch {
        }
    }
    return 0
}

HandyReplay_HistoryLoaded(el) {
    if !el
        return false
    if HandyReplay_FindNamed(el, 50020, ["HISTORY", "HISTÓRICO"])
        return true
    if HandyReplay_FindNamed(el, 50000, ["Play", "Reproduzir"])
        return true
    return false
}

HandyReplay_EnsureHistoryTab(hwnd) {
    global UIA
    el := UIA.ElementFromHandle(hwnd)
    if !el
        return false
    if (HandyReplay_HistoryLoaded(el))
        return true
    hist := HandyReplay_FindNamed(el, 50020, ["History", "Histórico"])
    if !hist
        return false
    try hist.Click()
    catch {
        try hist.Invoke()
        catch {
            return false
        }
    }
    Sleep 220
    el2 := UIA.ElementFromHandle(hwnd)
    return HandyReplay_HistoryLoaded(el2)
}

HandyReplay_WaitHistoryReady(hwnd, timeoutMs := 4000) {
    global UIA
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        el := UIA.ElementFromHandle(hwnd)
        if (el && HandyReplay_FindNamed(el, 50000, ["Play", "Reproduzir"]))
            return true
        Sleep 120
    }
    return false
}

HandyReplay_ClickFirstPlay(hwnd) {
    global UIA
    el := UIA.ElementFromHandle(hwnd)
    if !el
        return false
    play := HandyReplay_FindNamed(el, 50000, ["Play", "Reproduzir"])
    if !play
        return false
    try play.Click()
    catch {
        try play.Invoke()
        catch {
            return false
        }
    }
    return true
}

; End-to-end: activate/launch Handy → History → first Play. Leaves Handy open.
HandyReplay_PlayLastRecording() {
    StandardLoadingBar_Show("⏳ Opening Handy — please wait...", BANNER_ACCENT_INTERMEDIATE, {
        fontSize: 17,
        trackActiveMonitor: true
    })
    barOwned := true
    try {
        StandardLoadingBar_Update("⏳ Activating Handy...")
        hwnd := Handy_ActivateOrLaunch()
        if (!hwnd) {
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ Could not open Handy.", 2000, BANNER_ACCENT_ERROR)
            barOwned := false
            return false
        }

        StandardLoadingBar_Update("⏳ Opening History...")
        if (!HandyReplay_EnsureHistoryTab(hwnd)) {
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ Handy History tab not found.", 2200, BANNER_ACCENT_ERROR)
            barOwned := false
            return false
        }

        if (!HandyReplay_WaitHistoryReady(hwnd)) {
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ No recording Play button in History.", 2200, BANNER_ACCENT_ERROR)
            barOwned := false
            return false
        }

        StandardLoadingBar_Update("⏳ Playing last recording...")
        if (!HandyReplay_ClickFirstPlay(hwnd)) {
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ Could not click Play.", 2200, BANNER_ACCENT_ERROR)
            barOwned := false
            return false
        }

        StandardLoadingBar_Hide(0)
        ShowCenteredOverlay_Utils("✅ Playing last recording", 1500, BANNER_ACCENT_SUCCESS)
        barOwned := false
        return true
    } finally {
        if (barOwned) {
            try StandardLoadingBar_Hide(0)
        }
    }
}
