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

UtilitySelector_LvColumns() {
    global g_UtilitySelectorMode, g_UtilitySelectorCategory
    if (g_UtilitySelectorMode = "top")
        return ["Char", "Category", "Count"]
    if (g_UtilitySelectorCategory = "Prompts")
        return ["Char", "Category", "Name", "Author", "File"]
    if (g_UtilitySelectorCategory = "Hotstrings")
        return ["Char", "Name", "Text"]
    if (g_UtilitySelectorCategory = "Projects")
        return ["Char", "Name"]
    return ["Char", "Title"]
}

UtilitySelector_WindowTitle() {
    global g_UtilitySelectorMode, g_UtilitySelectorCategory
    if (g_UtilitySelectorMode = "category" && g_UtilitySelectorCategory != "")
        return "Utility Shortcuts - " . g_UtilitySelectorCategory
    return "Utility Shortcuts"
}

UtilitySelector_CreateGui() {
    global g_HotstringSelectorGui, g_HotstringSelectorLv, g_HotstringSelectorHint
    global g_UtilitySelectorMode, g_UtilitySelectorCategory

    mon := UtilitySelector_ActiveMonitorWorkArea()
    g_HotstringSelectorGui := Gui("+AlwaysOnTop +ToolWindow", UtilitySelector_WindowTitle())
    g_HotstringSelectorGui.SetFont("s10", "Segoe UI")
    g_HotstringSelectorHint := g_HotstringSelectorGui.Add("Text", "w820", UtilitySelector_HintText())
    cols := UtilitySelector_LvColumns()
    g_HotstringSelectorLv := g_HotstringSelectorGui.Add("ListView", "w820 h420 -Multi", cols)
    g_HotstringSelectorLv.OnEvent("DoubleClick", UtilitySelector_OnListActivate)

    showCrud := (g_UtilitySelectorMode = "category" && (g_UtilitySelectorCategory = "Prompts" ||
        g_UtilitySelectorCategory = "Hotstrings"))
    if (showCrud) {
        g_HotstringSelectorGui.Add("Button", "w100 Section", "Add").OnEvent("Click", UtilitySelector_OnAdd)
        g_HotstringSelectorGui.Add("Button", "w100 ys", "Edit").OnEvent("Click", UtilitySelector_OnEdit)
        g_HotstringSelectorGui.Add("Button", "w100 ys", "Delete").OnEvent("Click", UtilitySelector_OnDelete)
        g_HotstringSelectorGui.Add("Button", "w100 ys", "Close").OnEvent("Click", HandleHotstringEscape)
    } else {
        g_HotstringSelectorGui.Add("Button", "w100 Section", "Close").OnEvent("Click", HandleHotstringEscape)
    }
    g_HotstringSelectorGui.OnEvent("Close", HandleHotstringEscape)
    g_HotstringSelectorGui.OnEvent("Escape", HandleHotstringEscape)

    UtilitySelector_PopulateLv()

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
    UtilitySelector_BindModalHotkeys()
}

UtilitySelector_RebuildGui() {
    global g_HotstringSelectorGui, g_HotstringSelectorActive
    keepActive := g_HotstringSelectorActive
    UtilitySelector_UnbindModalHotkeys()
    if (IsObject(g_HotstringSelectorGui)) {
        try g_HotstringSelectorGui.Destroy()
        catch {
        }
        g_HotstringSelectorGui := false
    }
    UtilitySelector_CreateGui()
    g_HotstringSelectorActive := keepActive
}

ShowHotstringSelector() {
    global g_HotstringSelectorGui, g_HotstringSelectorActive
    global g_UtilitySelectorMode, g_UtilitySelectorCategory, g_UtilitySelectorRestoreHwnd
    global g_OnEscapePressed, g_HS_SelectorOpenFile, g_HS_SelectorCloseCheckTimer

    if (g_HotstringSelectorActive && IsObject(g_HotstringSelectorGui)) {
        CleanupHotstringSelector()
        Sleep 50
    }

    try {
        if (IsSet(g_ProjectSelectorActive) && g_ProjectSelectorActive && IsSet(CleanupProjectSelector)) {
            CleanupProjectSelector()
            Sleep 50
        }
    } catch {
    }

    try {
        sentinel := A_ScriptDir "\.cursor\wm_selector_open"
        if (FileExist(sentinel)) {
            closeReq := A_ScriptDir "\.cursor\wm_selector_close_request"
            try FileAppend("", closeReq)
            catch {
            }
            Sleep 50
        }
    } catch {
    }

    g_UtilitySelectorRestoreHwnd := 0
    try g_UtilitySelectorRestoreHwnd := WinGetID("A")
    catch {
        g_UtilitySelectorRestoreHwnd := 0
    }

    PromptData_Load()
    HotstringData_Load()
    ProjectData_Load(true)
    BuildMacroCharMap()
    UtilitySelector_RebuildPromptCharMap()

    g_UtilitySelectorMode := "top"
    g_UtilitySelectorCategory := ""
    UtilitySelector_CreateGui()

    g_HotstringSelectorActive := true
    g_OnEscapePressed := HandleHotstringEscape

    try {
        DirCreate(A_ScriptDir "\.cursor")
        FileAppend("", g_HS_SelectorOpenFile)
    } catch {
    }
    g_HS_SelectorCloseCheckTimer := SetTimer(Utils_CheckHotstringSelectorCloseRequest, 120)
}
