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

; Always click sidebar History so the list refreshes after a new recording.
HandyReplay_EnsureHistoryTab(hwnd) {
    global UIA
    try {
        el := UIA.ElementFromHandle(hwnd)
        if !el
            return false
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
        Sleep 350
        return true
    } catch {
        return false
    }
}

HandyReplay_ClickFirstPlay(hwnd) {
    global UIA
    loop 2 {
        try {
            el := UIA.ElementFromHandle(hwnd)
            play := el ? HandyReplay_FindNamed(el, 50000, ["Play", "Reproduzir"]) : 0
            if play {
                try play.Click()
                catch {
                    try play.Invoke()
                    catch {
                        return false
                    }
                }
                return true
            }
        } catch {
        }
        if (A_Index = 1)
            Sleep 250
    }
    return false
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
