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

; Sidebar nav row: Group with cursor-pointer + flex gap-2 containing the tab label Text.
HandyReplay_FindSidebarRow(el, labelNames) {
    if !el
        return 0
    try {
        for group in el.FindAll({ Type: 50026 }) {
            cn := ""
            try cn := group.ClassName
            if !(InStr(cn, "cursor-pointer") && InStr(cn, "flex gap-2 items-center"))
                continue
            for nm in labelNames {
                try {
                    if group.FindFirst({ Type: 50020, Name: nm })
                        return group
                } catch {
                }
            }
        }
    } catch {
    }
    histText := HandyReplay_FindNamed(el, 50020, labelNames)
    if !histText
        return 0
    cur := histText
    loop 10 {
        if !cur
            return 0
        try {
            if (cur.Type = 50026 && InStr(cur.ClassName, "cursor-pointer"))
                return cur
        } catch {
        }
        try
            cur := cur.GetParentElement()
        catch
            return 0
    }
    return 0
}

HandyReplay_FindHistorySidebarRow(el) {
    return HandyReplay_FindSidebarRow(el, ["History", "Histórico"])
}

HandyReplay_HistoryTabActive(el) {
    if HandyReplay_HistoryListVisible(el)
        return true
    row := HandyReplay_FindHistorySidebarRow(el)
    if !row
        return false
    cn := ""
    try cn := row.ClassName
    return InStr(cn, "bg-logo-primary")
}

HandyReplay_WaitHistoryTab(hwnd, maxWaitMs := 800) {
    global UIA
    start := A_TickCount
    loop {
        try {
            el := UIA.ElementFromHandle(hwnd)
            if HandyReplay_HistoryTabActive(el)
                return true
        } catch {
        }
        if ((A_TickCount - start) >= maxWaitMs)
            break
        Sleep 80
    }
    return false
}

; One explicit click strategy (invoke / legacy / uiaLeft / control / mouse).
HandyReplay_ApplyClickStrategy(el, hwnd, method) {
    if !el
        return false
    if (method = "invoke") {
        try {
            if (el.GetPropertyValue(UIA.Property.IsInvokePatternAvailable)) {
                el.InvokePattern.Invoke()
                return true
            }
        } catch {
        }
        return false
    }
    if (method = "legacy") {
        try {
            if (el.GetPropertyValue(UIA.Property.IsLegacyIAccessiblePatternAvailable)) {
                el.LegacyIAccessiblePattern.DoDefaultAction()
                return true
            }
        } catch {
        }
        return false
    }
    if (method = "uiaLeft") {
        try {
            el.Click("left")
            return true
        } catch {
        }
        return false
    }
    if (method = "control") {
        try {
            el.ControlClick(, "ahk_id " hwnd)
            return true
        } catch {
        }
        return false
    }
    if (method = "mouse") {
        try {
            pos := el.Location
            if (pos.w <= 0 || pos.h <= 0)
                return false
            saveMode := A_CoordModeMouse
            CoordMode("Mouse", "Screen")
            MouseGetPos(&prevX, &prevY)
            if !WinActive("ahk_id " hwnd) {
                WinActivate("ahk_id " hwnd)
                WinWaitActive("ahk_id " hwnd, , 0.4)
            }
            cx := pos.x + pos.w // 2
            cy := pos.y + pos.h // 2
            Click cx " " cy
            Sleep 40
            MouseMove(prevX, prevY)
            CoordMode("Mouse", saveMode)
            return true
        } catch {
        }
        return false
    }
    return false
}

; Try several UIA + real-mouse click paths until one runs without throwing.
HandyReplay_ClickWithStrategies(el, hwnd) {
    for method in ["invoke", "legacy", "uiaLeft", "control", "mouse"] {
        if HandyReplay_ApplyClickStrategy(el, hwnd, method)
            return true
    }
    return false
}

; Switch to History from General (or any tab); retries with different click methods.
HandyReplay_EnsureHistoryTab(hwnd) {
    global UIA
    try {
        try WinActivate("ahk_id " hwnd)
        if !WinActive("ahk_id " hwnd)
            WinWaitActive("ahk_id " hwnd, , 0.5)
        Sleep 80

        el := UIA.ElementFromHandle(hwnd)
        if !el
            return false
        if HandyReplay_HistoryTabActive(el)
            return true

        row := HandyReplay_FindHistorySidebarRow(el)
        if !row
            return false

        loop 3 {
            for method in ["mouse", "uiaLeft", "control", "legacy", "invoke"] {
                HandyReplay_ApplyClickStrategy(row, hwnd, method)
                if HandyReplay_WaitHistoryTab(hwnd, 500)
                    return true
            }
            Sleep 100
            el := UIA.ElementFromHandle(hwnd)
            row := HandyReplay_FindHistorySidebarRow(el)
            if !row
                break
        }

        try el := UIA.ElementFromHandle(hwnd)
        return HandyReplay_HistoryTabActive(el)
    } catch {
        return false
    }
}

HandyReplay_FindFirstPlay(el) {
    return HandyReplay_FindNamed(el, 50000, ["Play", "Reproduzir"])
}

HandyReplay_WaitFirstPlay(hwnd, maxWaitMs := 2000) {
    global UIA
    start := A_TickCount
    loop {
        try {
            el := UIA.ElementFromHandle(hwnd)
            play := HandyReplay_FindFirstPlay(el)
            if play
                return play
        } catch {
        }
        if ((A_TickCount - start) >= maxWaitMs)
            break
        Sleep 100
    }
    return 0
}

HandyReplay_ClickFirstPlay(hwnd) {
    play := HandyReplay_WaitFirstPlay(hwnd, 2000)
    if !play
        return false
    if HandyReplay_ClickWithStrategies(play, hwnd)
        return true
    try play.Invoke()
    catch {
        return false
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
        ; Overlay would intercept Click("left") / mouse clicks on Handy's webview.
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
