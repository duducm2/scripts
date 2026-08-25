; =============================================================================
; Utils module: handy_ai_model_gui.ahk
; ShowAiModelSelector GUI (Utility Shortcuts / project-selector ListView aesthetic)
; Extracted from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

global g_AiModelSelectorLv := false
global g_AiModelHotkeyHandlers := []

; =============================================================================
; ShowAiModelSelector() - ListView picker with digit / Enter / Esc capture
; =============================================================================
ShowAiModelSelector() {
    global g_AiModelSelectorGui, g_AiModelSelectorActive, g_AiModelSelectorLv
    global g_HandyAiModels, g_OnEscapePressed

    if (g_AiModelSelectorActive)
        return

    currentSlot := Handy_GetPersistedAiModelSlot()

    g_AiModelSelectorGui := Gui("+AlwaysOnTop +ToolWindow", "Handy AI Model")
    g_AiModelSelectorGui.SetFont("s10", "Segoe UI")
    g_AiModelSelectorGui.Add("Text", "w700",
        "Char = select   Enter/double-click = select   Esc = cancel")
    g_AiModelSelectorLv := g_AiModelSelectorGui.Add("ListView", "w700 h120 -Multi", ["Char", "Model",
        "Description", "Status"])
    g_AiModelSelectorLv.OnEvent("DoubleClick", AiModelSelector_OnListActivate)
    g_AiModelSelectorGui.Add("Button", "w100", "Close").OnEvent("Click", AiModelSelector_Cancel)
    g_AiModelSelectorGui.OnEvent("Close", AiModelSelector_Cancel)
    g_AiModelSelectorGui.OnEvent("Escape", AiModelSelector_Cancel)

    focusRow := 0
    loop 3 {
        num := A_Index
        if (!g_HandyAiModels.Has(num))
            continue
        model := g_HandyAiModels[num]
        status := (num = currentSlot) ? "Current" : ""
        row := g_AiModelSelectorLv.Add("", String(num), model.name, model.desc, status)
        if (num = currentSlot)
            focusRow := row
    }
    try g_AiModelSelectorLv.ModifyCol(1, 50)
    try g_AiModelSelectorLv.ModifyCol(2, 180)
    try g_AiModelSelectorLv.ModifyCol(3, 360)
    try g_AiModelSelectorLv.ModifyCol(4, 80)
    if (focusRow > 0) {
        try g_AiModelSelectorLv.Modify(focusRow, "Select Focus Vis")
        catch {
        }
    } else if (g_AiModelSelectorLv.GetCount() > 0) {
        try g_AiModelSelectorLv.Modify(1, "Select Focus Vis")
        catch {
        }
    }

    AiModelSelector_PositionAndShow()
    g_AiModelSelectorActive := true
    g_OnEscapePressed := AiModelSelector_Cancel
    try g_AiModelSelectorLv.Focus()
    catch {
    }
    AiModelSelector_BindModalHotkeys()
}

AiModelSelector_PositionAndShow() {
    global g_AiModelSelectorGui

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

    guiW := 720
    guiH := 220
    guiX := monitorLeft + (monitorWidth - guiW) // 2
    guiY := monitorTop + (monitorHeight - guiH) // 2
    if (guiX < monitorLeft + 20)
        guiX := monitorLeft + 20
    if (guiY < monitorTop + 20)
        guiY := monitorTop + 20
    g_AiModelSelectorGui.Show("x" . guiX . " y" . guiY . " w" . guiW . " h" . guiH)
}

AiModelSelector_SelectorHwnd() {
    global g_AiModelSelectorGui
    if (!IsObject(g_AiModelSelectorGui))
        return 0
    try
        return g_AiModelSelectorGui.Hwnd
    catch
        return 0
}

AiModelSelector_UnbindModalHotkeys() {
    global g_AiModelHotkeyHandlers
    hwnd := AiModelSelector_SelectorHwnd()
    if (hwnd) {
        try HotIfWinActive("ahk_id " hwnd)
        catch {
        }
    }
    for handler in g_AiModelHotkeyHandlers {
        try Hotkey(handler.key, "Off")
        catch {
        }
    }
    if (hwnd) {
        try HotIf()
        catch {
        }
    }
    g_AiModelHotkeyHandlers := []
}

AiModelSelector_BindModalHotkeys() {
    global g_AiModelHotkeyHandlers
    AiModelSelector_UnbindModalHotkeys()
    hwnd := AiModelSelector_SelectorHwnd()
    if (!hwnd)
        return
    try HotIfWinActive("ahk_id " hwnd)
    catch {
        return
    }

    loop 3 {
        key := String(A_Index)
        cb := AiModelSelector_Select.Bind(A_Index)
        try {
            Hotkey(key, cb, "On")
            g_AiModelHotkeyHandlers.Push({ key: key, handler: cb })
        } catch {
        }
    }
    try {
        Hotkey("Enter", AiModelSelector_OnEnter, "On")
        g_AiModelHotkeyHandlers.Push({ key: "Enter", handler: AiModelSelector_OnEnter })
    } catch {
    }
    try {
        Hotkey("Escape", AiModelSelector_Cancel, "On")
        g_AiModelHotkeyHandlers.Push({ key: "Escape", handler: AiModelSelector_Cancel })
    } catch {
    }

    try HotIf()
    catch {
    }
}

AiModelSelector_Select(selection, *) {
    global g_AiModelSelectorActive, g_HandyAiModels
    if (!g_AiModelSelectorActive)
        return
    AiModelSelector_Close()
    if (g_HandyAiModels.Has(selection))
        ExecuteHandyAiModelSelection(selection)
}

AiModelSelector_SelectedSlot() {
    global g_AiModelSelectorLv
    if (!IsObject(g_AiModelSelectorLv))
        return 0
    row := 0
    try row := g_AiModelSelectorLv.GetNext(0, "Focused")
    catch {
        row := 0
    }
    if (row < 1) {
        try row := g_AiModelSelectorLv.GetNext(0, "Selected")
        catch {
            row := 0
        }
    }
    if (row < 1)
        return 0
    ch := ""
    try ch := g_AiModelSelectorLv.GetText(row, 1)
    catch {
        return 0
    }
    if (!RegExMatch(ch, "^\d+$"))
        return 0
    return Integer(ch)
}

AiModelSelector_OnEnter(*) {
    global g_AiModelSelectorActive
    if (!g_AiModelSelectorActive)
        return
    slot := AiModelSelector_SelectedSlot()
    if (slot > 0)
        AiModelSelector_Select(slot)
}

AiModelSelector_OnListActivate(*) {
    AiModelSelector_OnEnter()
}

AiModelSelector_Cancel(*) {
    AiModelSelector_Close()
}

AiModelSelector_Close() {
    global g_AiModelSelectorGui, g_AiModelSelectorActive, g_AiModelSelectorLv, g_OnEscapePressed

    if (!g_AiModelSelectorActive)
        return

    g_AiModelSelectorActive := false
    AiModelSelector_UnbindModalHotkeys()
    if (g_OnEscapePressed = AiModelSelector_Cancel)
        g_OnEscapePressed := ""
    Utils_EnsureGlobalEscapeHotkey()

    if (IsObject(g_AiModelSelectorGui)) {
        try g_AiModelSelectorGui.Destroy()
        catch {
        }
    }
    g_AiModelSelectorGui := false
    g_AiModelSelectorLv := false
}
