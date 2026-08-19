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

HandyReplay_HistoryListVisible(el) {
    if !el
        return false
    try {
        if (el.FindFirst({ Type: 50000, Name: "Play" }) || el.FindFirst({ Type: 50000, Name: "Reproduzir" }))
            return true
    } catch {
    }
    try {
        if (el.FindFirst({ Type: 50020, Name: "HISTORY" }))
            return true
    } catch {
    }
    return false
}

; Sidebar rows are Groups with cursor-pointer; the "History" Text is a nested child.
HandyReplay_FindCursorPointerAncestor(el) {
    cur := el
    loop 8 {
        if !cur
            return 0
        cn := ""
        try cn := cur.ClassName
        if (InStr(cn, "cursor-pointer"))
            return cur
        try cur := cur.WalkTree("p")
        catch
            return 0
    }
    return 0
}

; Pattern Click/Invoke no-ops or throws on Handy Text/Group; use a real mouse click.
HandyReplay_PhysicalClick(el) {
    if !el
        return false
    try {
        el.Click("left")
        return true
    } catch {
    }
    try {
        el.ControlClick()
        return true
    } catch {
    }
    return false
}

; Always click sidebar History so the list refreshes after a new recording.
HandyReplay_EnsureHistoryTab(hwnd) {
    global UIA
    try {
        el := UIA.ElementFromHandle(hwnd)
        if !el
            return false
        alreadyOpen := HandyReplay_HistoryListVisible(el)
        hist := HandyReplay_FindNamed(el, 50020, ["History", "Histórico"])
        if !hist
            return alreadyOpen
        target := HandyReplay_FindCursorPointerAncestor(hist)
        if !target
            target := hist
        try WinActivate("ahk_id " hwnd)
        clicked := HandyReplay_PhysicalClick(target)
        if (!clicked && !alreadyOpen)
            return false
        Sleep 350
        elAfter := 0
        try elAfter := UIA.ElementFromHandle(hwnd)
        return HandyReplay_HistoryListVisible(elAfter) || alreadyOpen
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
        ; Overlay would intercept Click("left") on Handy's webview.
        StandardLoadingBar_Hide(0)
        barOwned := false
        if (!HandyReplay_EnsureHistoryTab(hwnd)) {
            ShowCenteredOverlay_Utils("❌ Handy History tab not found.", 2200, BANNER_ACCENT_ERROR)
            return false
        }

        if (!HandyReplay_ClickFirstPlay(hwnd)) {
            ShowCenteredOverlay_Utils("❌ Could not click Play.", 2200, BANNER_ACCENT_ERROR)
            return false
        }

        ShowCenteredOverlay_Utils("✅ Playing last recording", 1500, BANNER_ACCENT_SUCCESS)
        return true
    } finally {
        if (barOwned) {
            try StandardLoadingBar_Hide(0)
        }
    }
}
