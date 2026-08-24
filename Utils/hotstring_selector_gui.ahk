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

UtilitySelector_IsPromptsView() {
    global g_UtilitySelectorMode, g_UtilitySelectorCategory
    return (g_UtilitySelectorMode = "category" && g_UtilitySelectorCategory = "Prompts")
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
    UtilitySelector_ApplyFilterChrome()

    if (!IsObject(g_HotstringSelectorLv))
        return
    if (g_UtilitySelectorMode = "top") {
        try g_HotstringSelectorLv.ModifyCol(1, 50, "Char")
        try g_HotstringSelectorLv.ModifyCol(2, 220, "Category")
        try g_HotstringSelectorLv.ModifyCol(3, 80, "Count")
        try g_HotstringSelectorLv.ModifyCol(4, 0, "")
        try g_HotstringSelectorLv.ModifyCol(5, 0, "")
        return
    }
    if (g_UtilitySelectorCategory = "Prompts") {
        try g_HotstringSelectorLv.ModifyCol(1, 50, "Char")
        try g_HotstringSelectorLv.ModifyCol(2, 90, "Category")
        try g_HotstringSelectorLv.ModifyCol(3, 280, "Name")
        try g_HotstringSelectorLv.ModifyCol(4, 90, "Out")
        try g_HotstringSelectorLv.ModifyCol(5, 280, "File")
        return
    }
    if (g_UtilitySelectorCategory = "Hotstrings") {
        try g_HotstringSelectorLv.ModifyCol(1, 50, "Char")
        try g_HotstringSelectorLv.ModifyCol(2, 260, "Name")
        try g_HotstringSelectorLv.ModifyCol(3, 420, "Text")
        try g_HotstringSelectorLv.ModifyCol(4, 0, "")
        try g_HotstringSelectorLv.ModifyCol(5, 0, "")
        return
    }
    if (g_UtilitySelectorCategory = "Projects") {
        try g_HotstringSelectorLv.ModifyCol(1, 50, "Char")
        try g_HotstringSelectorLv.ModifyCol(2, 400, "Name")
        try g_HotstringSelectorLv.ModifyCol(3, 0, "")
        try g_HotstringSelectorLv.ModifyCol(4, 0, "")
        try g_HotstringSelectorLv.ModifyCol(5, 0, "")
        return
    }
    try g_HotstringSelectorLv.ModifyCol(1, 50, "Char")
    try g_HotstringSelectorLv.ModifyCol(2, 500, "Title")
    try g_HotstringSelectorLv.ModifyCol(3, 0, "")
    try g_HotstringSelectorLv.ModifyCol(4, 0, "")
    try g_HotstringSelectorLv.ModifyCol(5, 0, "")
    return
}

UtilitySelector_WindowTitle() {
    global g_UtilitySelectorMode, g_UtilitySelectorCategory
    if (g_UtilitySelectorMode = "category" && g_UtilitySelectorCategory != "")
        return "Utility Shortcuts - " . g_UtilitySelectorCategory
    return "Utility Shortcuts"
}

UtilitySelector_ApplyFilterChrome() {
    global g_HotstringSelectorFilterLabel, g_HotstringSelectorFilterCtrl
    global g_HotstringSelectorFilterEnterBtn

    showFilter := UtilitySelector_IsPromptsView()
    if (IsObject(g_HotstringSelectorFilterLabel)) {
        try g_HotstringSelectorFilterLabel.Visible := showFilter
        catch {
        }
    }
    if (IsObject(g_HotstringSelectorFilterCtrl)) {
        try g_HotstringSelectorFilterCtrl.Visible := showFilter
        try g_HotstringSelectorFilterCtrl.Enabled := showFilter
        catch {
        }
        if (!showFilter) {
            try g_HotstringSelectorFilterCtrl.Value := ""
            catch {
            }
        }
    }
    if (IsObject(g_HotstringSelectorFilterEnterBtn)) {
        try {
            if (showFilter)
                g_HotstringSelectorFilterEnterBtn.Opt("+Default")
            else
                g_HotstringSelectorFilterEnterBtn.Opt("-Default")
        } catch {
        }
    }
    UtilitySelector_LayoutControls()
}

UtilitySelector_LayoutControls() {
    global g_HotstringSelectorHint, g_HotstringSelectorFilterLabel, g_HotstringSelectorFilterCtrl
    global g_HotstringSelectorLv, g_HotstringSelectorBtnAdd, g_HotstringSelectorBtnEdit
    global g_HotstringSelectorBtnDelete, g_HotstringSelectorBtnClose

    if (!IsObject(g_HotstringSelectorHint) || !IsObject(g_HotstringSelectorLv))
        return
    showFilter := UtilitySelector_IsPromptsView()
    try {
        g_HotstringSelectorHint.GetPos(&hx, &hy, &hw, &hh)
    } catch {
        return
    }
    if (hw < 1)
        hw := 820
    filterY := hy + hh + 6
    filterH := 24
    lvY := showFilter ? (filterY + filterH + 8) : filterY
    lvH := showFilter ? 392 : 420
    if (IsObject(g_HotstringSelectorFilterLabel) && showFilter) {
        try g_HotstringSelectorFilterLabel.Move(hx, filterY, 70, filterH)
        catch {
        }
    }
    if (IsObject(g_HotstringSelectorFilterCtrl) && showFilter) {
        try g_HotstringSelectorFilterCtrl.Move(hx + 74, filterY, hw - 74, filterH)
        catch {
        }
    }
    try g_HotstringSelectorLv.Move(hx, lvY, hw, lvH)
    catch {
    }
    btnY := lvY + lvH + 8
    if (IsObject(g_HotstringSelectorBtnAdd)) {
        try {
            g_HotstringSelectorBtnAdd.Move(hx, btnY)
            if (IsObject(g_HotstringSelectorBtnEdit))
                g_HotstringSelectorBtnEdit.Move(hx + 110, btnY)
            if (IsObject(g_HotstringSelectorBtnDelete))
                g_HotstringSelectorBtnDelete.Move(hx + 220, btnY)
            if (IsObject(g_HotstringSelectorBtnClose))
                g_HotstringSelectorBtnClose.Move(hx + 330, btnY)
        } catch {
        }
    }
}

UtilitySelector_CreateGui() {
    global g_HotstringSelectorGui, g_HotstringSelectorLv, g_HotstringSelectorHint
    global g_HotstringSelectorFilterLabel, g_HotstringSelectorFilterCtrl
    global g_HotstringSelectorFilterEnterBtn
    global g_HotstringSelectorBtnAdd, g_HotstringSelectorBtnEdit, g_HotstringSelectorBtnDelete
    global g_HotstringSelectorBtnClose
    global g_HotstringSelectorGuiReady

    g_HotstringSelectorGui := Gui("+AlwaysOnTop +ToolWindow", "Utility Shortcuts")
    g_HotstringSelectorGui.SetFont("s10", "Segoe UI")
    g_HotstringSelectorHint := g_HotstringSelectorGui.Add("Text", "xm w820", UtilitySelector_HintText())
    g_HotstringSelectorFilterLabel := g_HotstringSelectorGui.Add("Text", "xm w70 Hidden", "Filter:")
    g_HotstringSelectorFilterCtrl := g_HotstringSelectorGui.Add("Edit", "yp w740 Hidden")
    g_HotstringSelectorFilterCtrl.OnEvent("Change", UtilitySelector_OnFilterChange)
    g_HotstringSelectorFilterCtrl.OnEvent("Focus", UtilitySelector_OnFilterFocus)
    g_HotstringSelectorFilterCtrl.OnEvent("LoseFocus", UtilitySelector_OnFilterKillFocus)
    g_HotstringSelectorLv := g_HotstringSelectorGui.Add("ListView", "xm w820 h420 -Multi", ["Char", "Category", "Count",
        "Out", "File"])
    g_HotstringSelectorLv.OnEvent("DoubleClick", UtilitySelector_OnListActivate)

    g_HotstringSelectorBtnAdd := g_HotstringSelectorGui.Add("Button", "w100 Section", "Add")
    g_HotstringSelectorBtnAdd.OnEvent("Click", UtilitySelector_OnAdd)
    g_HotstringSelectorBtnEdit := g_HotstringSelectorGui.Add("Button", "w100 ys", "Edit")
    g_HotstringSelectorBtnEdit.OnEvent("Click", UtilitySelector_OnEdit)
    g_HotstringSelectorBtnDelete := g_HotstringSelectorGui.Add("Button", "w100 ys", "Delete")
    g_HotstringSelectorBtnDelete.OnEvent("Click", UtilitySelector_OnDelete)
    g_HotstringSelectorBtnClose := g_HotstringSelectorGui.Add("Button", "w100 ys", "Close")
    g_HotstringSelectorBtnClose.OnEvent("Click", HandleHotstringEscape)
    g_HotstringSelectorFilterEnterBtn := g_HotstringSelectorGui.Add("Button", "Hidden x0 y0 w1 h1", "OK")
    g_HotstringSelectorFilterEnterBtn.OnEvent("Click", UtilitySelector_OnEnter)
    g_HotstringSelectorGui.OnEvent("Close", HandleHotstringEscape)
    g_HotstringSelectorGui.OnEvent("Escape", HandleHotstringEscape)
    g_HotstringSelectorGuiReady := true
}

UtilitySelector_PositionAndShow() {
    global g_HotstringSelectorGui, g_UtilitySelectorNoActivate
    mon := UtilitySelector_ActiveMonitorWorkArea()
    guiW := 850
    guiH := 520
    guiX := mon.left + (mon.width - guiW) // 2
    guiY := mon.top + (mon.height - guiH) // 2
    if (guiX < mon.left + 20)
        guiX := mon.left + 20
    if (guiY < mon.top + 20)
        guiY := mon.top + 20
    if (g_UtilitySelectorNoActivate) {
        g_HotstringSelectorGui.Show("x" . guiX . " y" . guiY . " NA")
    } else {
        g_HotstringSelectorGui.Show("x" . guiX . " y" . guiY)
    }
    UtilitySelector_LayoutControls()
    if (!g_UtilitySelectorNoActivate)
        UtilitySelector_FocusAfterShow()
}

UtilitySelector_FocusAfterShow() {
    if (UtilitySelector_IsPromptsView()) {
        UtilitySelector_FocusFilterField()
        return
    }
    global g_HotstringSelectorLv
    try g_HotstringSelectorLv.Focus()
    catch {
    }
}

UtilitySelector_RefreshView() {
    global g_UtilitySelectorFilterTyping
    UtilitySelector_ApplyChrome()
    UtilitySelector_PopulateLv()
    if (UtilitySelector_IsPromptsView() && (g_UtilitySelectorFilterTyping || UtilitySelector_IsFilterFocused()))
        UtilitySelector_BindFilterTypingHotkeys()
    else
        UtilitySelector_BindModalHotkeys()
}

UtilitySelector_RebuildGui() {
    global g_HotstringSelectorActive, g_UtilitySelectorSuppressFilterKillFocus
    g_UtilitySelectorSuppressFilterKillFocus := true
    if (!UtilitySelector_GuiIsAlive()) {
        UtilitySelector_CreateGui()
        g_HotstringSelectorActive := true
        UtilitySelector_RefreshView()
        UtilitySelector_PositionAndShow()
        g_UtilitySelectorSuppressFilterKillFocus := false
        return
    }
    UtilitySelector_UnbindModalHotkeys()
    UtilitySelector_RefreshView()
    g_UtilitySelectorSuppressFilterKillFocus := false
    UtilitySelector_FocusAfterShow()
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

UtilitySelector_IsFocusFragileWindow() {
    try {
        exe := WinGetProcessName("A")
        if (exe = "PowerToys.PowerLauncher.exe" || exe = "Microsoft.CmdPal.UI.exe")
            return true
    } catch {
    }
    return false
}

ShowHotstringSelector(initialCategory := "") {
    global g_HotstringSelectorGui, g_HotstringSelectorActive
    global g_UtilitySelectorMode, g_UtilitySelectorCategory, g_UtilitySelectorRestoreHwnd
    global g_UtilitySelectorNoActivate, g_OnEscapePressed

    if (g_HotstringSelectorActive && UtilitySelector_GuiIsAlive()) {
        ; Same view again → toggle close. Different category → switch without full reopen.
        if (initialCategory != "") {
            if (g_UtilitySelectorMode = "category" && g_UtilitySelectorCategory = initialCategory) {
                CleanupHotstringSelector()
                return
            }
            UtilitySelector_SwitchToCategory(initialCategory)
            return
        }
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

    g_UtilitySelectorNoActivate := UtilitySelector_IsFocusFragileWindow()

    if (initialCategory != "") {
        g_UtilitySelectorMode := "category"
        g_UtilitySelectorCategory := initialCategory
    } else {
        g_UtilitySelectorMode := "top"
        g_UtilitySelectorCategory := ""
    }

    if (!UtilitySelector_GuiIsAlive())
        UtilitySelector_CreateGui()

    g_HotstringSelectorActive := true
    UtilitySelector_RefreshView()
    UtilitySelector_PositionAndShow()

    g_OnEscapePressed := HandleHotstringEscape
    ; Drive FileAppend/FileExist after paint so the hotkey is not blocked on Google Drive I/O.
    SetTimer(UtilitySelector_StartIpc, -1)
}
