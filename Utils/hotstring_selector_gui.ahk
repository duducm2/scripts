; =============================================================================
; Utils module: hotstring_selector_gui.ahk
; ShowHotstringSelector ListView GUI (project-selector aesthetic)
; =============================================================================

UtilitySelector_ActiveMonitorWorkArea() {
    activeWin := 0
    try activeWin := WinGetID("A")
    catch {
        activeWin := 0
    }
    MonitorGetWorkArea(1, &monitorLeft, &monitorTop, &monitorRight, &monitorBottom)
    monitorWidth := monitorRight - monitorLeft
    monitorHeight := monitorBottom - monitorTop
    if (activeWin && activeWin != 0) {
        rect := Buffer(16, 0)
        if (DllCall("GetWindowRect", "ptr", activeWin, "ptr", rect)) {
            winLeft := NumGet(rect, 0, "int")
            winTop := NumGet(rect, 4, "int")
            winRight := NumGet(rect, 8, "int")
            winBottom := NumGet(rect, 12, "int")
            centerX := winLeft + (winRight - winLeft) // 2
            centerY := winTop + (winBottom - winTop) // 2
            monitorCount := MonitorGetCount()
            loop monitorCount {
                idx := A_Index
                MonitorGetWorkArea(idx, &l, &t, &r, &b)
                if (centerX >= l && centerX <= r && centerY >= t && centerY <= b) {
                    monitorLeft := l
                    monitorTop := t
                    monitorRight := r
                    monitorBottom := b
                    monitorWidth := r - l
                    monitorHeight := b - t
                    break
                }
            }
        }
    }
    return { left: monitorLeft, top: monitorTop, width: monitorWidth, height: monitorHeight }
}

UtilitySelector_GuiIsAlive() {
    global g_HotstringSelectorGui, g_HotstringSelectorGuiReady
    if (!g_HotstringSelectorGuiReady || !IsObject(g_HotstringSelectorGui))
        return false
    try {
        if !g_HotstringSelectorGui.Hwnd
            return false
    } catch {
        return false
    }
    return true
}

UtilitySelector_ApplyChrome() {
    global g_HotstringSelectorGui, g_HotstringSelectorLv, g_HotstringSelectorHint
    global g_HotstringSelectorBtnAdd, g_HotstringSelectorBtnEdit, g_HotstringSelectorBtnDelete
    global g_UtilitySelectorMode, g_UtilitySelectorCategory

    try g_HotstringSelectorGui.Title := UtilitySelector_WindowTitle()
    catch {
    }
    if (IsObject(g_HotstringSelectorHint)) {
        try g_HotstringSelectorHint.Text := UtilitySelector_HintText()
        catch {
        }
    }

    showCrud := (g_UtilitySelectorMode = "category" && (g_UtilitySelectorCategory = "Prompts" ||
        g_UtilitySelectorCategory = "Hotstrings"))
    if (IsObject(g_HotstringSelectorBtnAdd)) {
        try g_HotstringSelectorBtnAdd.Visible := showCrud
        try g_HotstringSelectorBtnEdit.Visible := showCrud
        try g_HotstringSelectorBtnDelete.Visible := showCrud
    }

    if (!IsObject(g_HotstringSelectorLv))
        return
    if (g_UtilitySelectorMode = "top") {
        try g_HotstringSelectorLv.ModifyCol(1, 50, "Char")
        try g_HotstringSelectorLv.ModifyCol(2, 220, "Category")
        try g_HotstringSelectorLv.ModifyCol(3, 80, "Count")
        try g_HotstringSelectorLv.ModifyCol(4, 0, "")
        return
    }
    if (g_UtilitySelectorCategory = "Prompts") {
        try g_HotstringSelectorLv.ModifyCol(1, 50, "Char")
        try g_HotstringSelectorLv.ModifyCol(2, 100, "Category")
        try g_HotstringSelectorLv.ModifyCol(3, 340, "Name")
        try g_HotstringSelectorLv.ModifyCol(4, 330, "File")
        return
    }
    if (g_UtilitySelectorCategory = "Hotstrings") {
        try g_HotstringSelectorLv.ModifyCol(1, 50, "Char")
        try g_HotstringSelectorLv.ModifyCol(2, 260, "Name")
        try g_HotstringSelectorLv.ModifyCol(3, 420, "Text")
        try g_HotstringSelectorLv.ModifyCol(4, 0, "")
        return
    }
    if (g_UtilitySelectorCategory = "Projects") {
        try g_HotstringSelectorLv.ModifyCol(1, 50, "Char")
        try g_HotstringSelectorLv.ModifyCol(2, 400, "Name")
        try g_HotstringSelectorLv.ModifyCol(3, 0, "")
        try g_HotstringSelectorLv.ModifyCol(4, 0, "")
        return
    }
    try g_HotstringSelectorLv.ModifyCol(1, 50, "Char")
    try g_HotstringSelectorLv.ModifyCol(2, 500, "Title")
    try g_HotstringSelectorLv.ModifyCol(3, 0, "")
    try g_HotstringSelectorLv.ModifyCol(4, 0, "")
}

UtilitySelector_WindowTitle() {
    global g_UtilitySelectorMode, g_UtilitySelectorCategory
    if (g_UtilitySelectorMode = "category" && g_UtilitySelectorCategory != "")
        return "Utility Shortcuts - " . g_UtilitySelectorCategory
    return "Utility Shortcuts"
}

UtilitySelector_CreateGui() {
    global g_HotstringSelectorGui, g_HotstringSelectorLv, g_HotstringSelectorHint
    global g_HotstringSelectorBtnAdd, g_HotstringSelectorBtnEdit, g_HotstringSelectorBtnDelete
    global g_HotstringSelectorGuiReady

    g_HotstringSelectorGui := Gui("+AlwaysOnTop +ToolWindow", "Utility Shortcuts")
    g_HotstringSelectorGui.SetFont("s10", "Segoe UI")
    g_HotstringSelectorHint := g_HotstringSelectorGui.Add("Text", "w820", UtilitySelector_HintText())
    g_HotstringSelectorLv := g_HotstringSelectorGui.Add("ListView", "w820 h420 -Multi", ["Char", "Category", "Count",
        "File"])
    g_HotstringSelectorLv.OnEvent("DoubleClick", UtilitySelector_OnListActivate)

    g_HotstringSelectorBtnAdd := g_HotstringSelectorGui.Add("Button", "w100 Section", "Add")
    g_HotstringSelectorBtnAdd.OnEvent("Click", UtilitySelector_OnAdd)
    g_HotstringSelectorBtnEdit := g_HotstringSelectorGui.Add("Button", "w100 ys", "Edit")
    g_HotstringSelectorBtnEdit.OnEvent("Click", UtilitySelector_OnEdit)
    g_HotstringSelectorBtnDelete := g_HotstringSelectorGui.Add("Button", "w100 ys", "Delete")
    g_HotstringSelectorBtnDelete.OnEvent("Click", UtilitySelector_OnDelete)
    g_HotstringSelectorGui.Add("Button", "w100 ys", "Close").OnEvent("Click", HandleHotstringEscape)
    g_HotstringSelectorGui.OnEvent("Close", HandleHotstringEscape)
    g_HotstringSelectorGui.OnEvent("Escape", HandleHotstringEscape)
    g_HotstringSelectorGuiReady := true
}

UtilitySelector_PositionAndShow() {
    global g_HotstringSelectorGui
    mon := UtilitySelector_ActiveMonitorWorkArea()
    guiW := 850
    guiH := 520
    guiX := mon.left + (mon.width - guiW) // 2
    guiY := mon.top + (mon.height - guiH) // 2
    if (guiX < mon.left + 20)
        guiX := mon.left + 20
    if (guiY < mon.top + 20)
        guiY := mon.top + 20
    g_HotstringSelectorGui.Show("x" . guiX . " y" . guiY)
    try g_HotstringSelectorLv.Focus()
    catch {
    }
}

UtilitySelector_RefreshView() {
    UtilitySelector_ApplyChrome()
    UtilitySelector_PopulateLv()
    UtilitySelector_BindModalHotkeys()
}

UtilitySelector_RebuildGui() {
    global g_HotstringSelectorActive
    if (!UtilitySelector_GuiIsAlive()) {
        UtilitySelector_CreateGui()
        UtilitySelector_RefreshView()
        UtilitySelector_PositionAndShow()
        g_HotstringSelectorActive := true
        return
    }
    UtilitySelector_UnbindModalHotkeys()
    UtilitySelector_RefreshView()
}

UtilitySelector_StartIpc(*) {
    global g_HotstringSelectorActive, g_HS_SelectorOpenFile, g_HS_SelectorCloseCheckTimer
    if (!g_HotstringSelectorActive)
        return
    try {
        sentinel := A_ScriptDir "\.cursor\wm_selector_open"
        if (FileExist(sentinel)) {
            closeReq := A_ScriptDir "\.cursor\wm_selector_close_request"
            try FileAppend("", closeReq)
            catch {
            }
        }
    } catch {
    }
    try {
        if (!FileExist(A_ScriptDir "\.cursor"))
            DirCreate(A_ScriptDir "\.cursor")
        FileAppend("", g_HS_SelectorOpenFile)
    } catch {
    }
    g_HS_SelectorCloseCheckTimer := SetTimer(Utils_CheckHotstringSelectorCloseRequest, 120)
}

ShowHotstringSelector() {
    global g_HotstringSelectorGui, g_HotstringSelectorActive
    global g_UtilitySelectorMode, g_UtilitySelectorCategory, g_UtilitySelectorRestoreHwnd
    global g_OnEscapePressed

    if (g_HotstringSelectorActive && UtilitySelector_GuiIsAlive()) {
        CleanupHotstringSelector()
        return
    }

    try {
        if (IsSet(g_ProjectSelectorActive) && g_ProjectSelectorActive && IsSet(CleanupProjectSelector))
            CleanupProjectSelector()
    } catch {
    }

    g_UtilitySelectorRestoreHwnd := 0
    try g_UtilitySelectorRestoreHwnd := WinGetID("A")
    catch {
        g_UtilitySelectorRestoreHwnd := 0
    }

    g_UtilitySelectorMode := "top"
    g_UtilitySelectorCategory := ""

    if (!UtilitySelector_GuiIsAlive())
        UtilitySelector_CreateGui()

    UtilitySelector_RefreshView()
    UtilitySelector_PositionAndShow()

    g_HotstringSelectorActive := true
    g_OnEscapePressed := HandleHotstringEscape
    ; Drive FileAppend/FileExist after paint so the hotkey is not blocked on Google Drive I/O.
    SetTimer(UtilitySelector_StartIpc, -1)
}
