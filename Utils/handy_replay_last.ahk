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

HandyReplay_ElementIsOnScreen(el) {
    if !el
        return false
    try {
        if el.GetPropertyValue(UIA.Property.IsOffscreen)
            return false
    } catch {
        try {
            if el.IsOffscreen
                return false
        } catch {
        }
    }
    try {
        pos := el.Location
        if (!IsObject(pos) || pos.w <= 0 || pos.h <= 0)
            return false
    } catch {
        return false
    }
    return true
}

; History content chrome (title or folder button) — present only on History view.
HandyReplay_HistoryContentChromeVisible(el) {
    if !el
        return false
    if HandyReplay_FindNamed(el, 50020, ["HISTORY", "HISTÓRICO"])
        return true
    if HandyReplay_FindNamed(el, 50000, ["Open Recordings Folder", "Abrir pasta de gravações"])
        return true
    return false
}

HandyReplay_HistoryHasPlay(el) {
    if !el
        return false
    return HandyReplay_FindNamed(el, 50000, ["Play", "Reproduzir"]) != 0
}

HandyReplay_HistoryListVisible(el) {
    return HandyReplay_HistoryContentChromeVisible(el) || HandyReplay_HistoryHasPlay(el)
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

; Last-resort locator from handy.md dump: RootWebArea → root → History Group (index 4).
HandyReplay_FindHistorySidebarRowByPath(el) {
    if !el
        return 0
    try {
        row := el.ElementFromPath("1,1,2,1,1,1,2,1,1,1,1,2,1,4")
        if !row
            return 0
        try {
            if row.FindFirst({ Type: 50020, Name: "History" })
                return row
        } catch {
        }
        try {
            if row.FindFirst({ Type: 50020, Name: "Histórico" })
                return row
        } catch {
        }
    } catch {
    }
    return 0
}

; Name-based row, then text→parent walk (inside FindSidebarRow), then ElementFromPath.
HandyReplay_LocateHistorySidebarRow(el) {
    row := HandyReplay_FindHistorySidebarRow(el)
    if row
        return row
    return HandyReplay_FindHistorySidebarRowByPath(el)
}

HandyReplay_HistorySidebarSelected(el) {
    row := HandyReplay_LocateHistorySidebarRow(el)
    if !row
        return false
    cn := ""
    try cn := row.ClassName
    return InStr(cn, "bg-logo-primary")
}

; Light check: already on History — do not re-click.
HandyReplay_HistoryTabActiveLight(el) {
    if !el
        return false
    if HandyReplay_HistorySidebarSelected(el)
        return true
    return HandyReplay_HistoryContentChromeVisible(el)
}

; Strict post-click gates: selected sidebar + content chrome (+ Play when recordings exist).
; If chrome is present but no Play yet, still accept (empty list / slow paint); Play wait handles R.
HandyReplay_HistoryTabActiveStrict(el) {
    if !el
        return false
    if !HandyReplay_HistorySidebarSelected(el)
        return false
    if !HandyReplay_HistoryContentChromeVisible(el)
        return false
    return true
}

; Backward-compatible alias used by older call sites / Wait helpers.
HandyReplay_HistoryTabActive(el) {
    return HandyReplay_HistoryTabActiveLight(el)
}

HandyReplay_WaitHistoryTab(hwnd, maxWaitMs := 2000, strict := true) {
    global UIA
    start := A_TickCount
    loop {
        try {
            el := UIA.ElementFromHandle(hwnd)
            if (strict ? HandyReplay_HistoryTabActiveStrict(el) : HandyReplay_HistoryTabActiveLight(el))
                return true
        } catch {
        }
        if ((A_TickCount - start) >= maxWaitMs)
            break
        Sleep 80
    }
    return false
}

; Primary quality gate: UIA locate → valid on-screen Location → real screen Click.
HandyReplay_ClickElementWithCursor(el, hwnd) {
    if !el
        return false
    if !HandyReplay_ElementIsOnScreen(el)
        return false

    saveMode := A_CoordModeMouse
    CoordMode("Mouse", "Screen")
    try {
        if !WinExist("ahk_id " hwnd)
            return false
        if !WinActive("ahk_id " hwnd) {
            WinActivate("ahk_id " hwnd)
            WinWaitActive("ahk_id " hwnd, , 0.5)
        }
        Sleep 40

        MouseGetPos(&prevX, &prevY)
        cx := 0
        cy := 0
        gotPoint := false

        try {
            pt := el.GetClickablePoint()
            if (IsObject(pt) && (pt.x || pt.y)) {
                cx := pt.x
                cy := pt.y
                gotPoint := true
            }
        } catch {
        }

        if !gotPoint {
            pos := el.Location
            if (!IsObject(pos) || pos.w <= 0 || pos.h <= 0)
                return false
            cx := pos.x + pos.w // 2
            cy := pos.y + pos.h // 2
        }

        Click(cx " " cy)
        Sleep 60
        MouseMove(prevX, prevY)
        return true
    } catch {
        return false
    } finally {
        try CoordMode("Mouse", saveMode)
    }
}

; One explicit click strategy (cursor / uiaLeft / control / legacy / invoke).
HandyReplay_ApplyClickStrategy(el, hwnd, method) {
    if !el
        return false
    if (method = "cursor" || method = "mouse") {
        return HandyReplay_ClickElementWithCursor(el, hwnd)
    }
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
    return false
}

; Try several UIA + real-mouse click paths until one runs without throwing.
HandyReplay_ClickWithStrategies(el, hwnd) {
    for method in ["cursor", "uiaLeft", "control", "legacy", "invoke"] {
        if HandyReplay_ApplyClickStrategy(el, hwnd, method)
            return true
    }
    return false
}

; Switch to History from General (or any tab); cursor-click first, verify-after-each.
HandyReplay_EnsureHistoryTab(hwnd) {
    global UIA
    try {
        if !hwnd || !WinExist("ahk_id " hwnd)
            return false

        try WinActivate("ahk_id " hwnd)
        if !WinActive("ahk_id " hwnd)
            WinWaitActive("ahk_id " hwnd, , 0.5)

        ; Sidebar / footer must be interactive before we hunt History.
        try Handy_WaitForMainUiReady(hwnd, 2500)

        el := UIA.ElementFromHandle(hwnd)
        if !el
            return false
        if HandyReplay_HistoryTabActiveLight(el)
            return true

        loop 4 {
            el := UIA.ElementFromHandle(hwnd)
            if !el
                break

            if HandyReplay_HistoryTabActiveLight(el)
                return true

            row := HandyReplay_LocateHistorySidebarRow(el)
            if !row {
                Sleep 120
                continue
            }

            ; Cursor click is the primary verified path for Tauri/webview Groups.
            for method in ["cursor", "uiaLeft", "control"] {
                if !HandyReplay_ApplyClickStrategy(row, hwnd, method)
                    continue
                if HandyReplay_WaitHistoryTab(hwnd, 2000, true)
                    return true
                ; Soft accept: chrome appeared even if selected class lagged.
                try {
                    elCheck := UIA.ElementFromHandle(hwnd)
                    if (HandyReplay_HistoryContentChromeVisible(elCheck) && HandyReplay_HistorySidebarSelected(elCheck))
                        return true
                } catch {
                }
            }

            Sleep 150
        }

        try el := UIA.ElementFromHandle(hwnd)
        return HandyReplay_HistoryTabActiveStrict(el)
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
            if (play && HandyReplay_ElementIsOnScreen(play))
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
    play := HandyReplay_WaitFirstPlay(hwnd, 2500)
    if !play
        return false
    ; Prefer real cursor click; fall back to other strategies only if cursor fails.
    if HandyReplay_ClickElementWithCursor(play, hwnd)
        return true
    if HandyReplay_ClickWithStrategies(play, hwnd)
        return true
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
