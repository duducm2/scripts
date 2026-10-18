; =============================================================================
; Utils module: handy_ai_model_gui.ahk
; ShowAiModelSelector GUI (Utility Shortcuts / project-selector ListView aesthetic)
; Extracted from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; CRUD for ExtraModels (slots 4+): Insert/a add, E edit, Delete remove.
; Slots 1-3 are immutable builtins.
; =============================================================================

global g_AiModelSelectorLv := false
global g_AiModelSelectorPrevHwnd := 0
global g_AiModelHotkeyHandlers := []
global g_AiModelSelectorFocusSlot := 0

; =============================================================================
; ShowAiModelSelector() - ListView picker with digit / Enter / Esc capture
; =============================================================================
ShowAiModelSelector() {
    global g_AiModelSelectorActive, g_AiModelSelectorPrevHwnd

    if (g_AiModelSelectorActive)
        return

    g_AiModelSelectorPrevHwnd := 0
    try g_AiModelSelectorPrevHwnd := WinGetID("A")
    catch {
        g_AiModelSelectorPrevHwnd := 0
    }

    g_AiModelSelectorActive := true
    AiModelSelector_Rebuild()
}

AiModelSelector_SafeDestroyGui(gui) {
    if (!IsObject(gui))
        return
    try gui.Destroy()
    catch {
    }
}

AiModelSelector_Rebuild() {
    global g_AiModelSelectorGui, g_AiModelSelectorLv, g_AiModelSelectorActive
    global g_HandyAiModels, g_OnEscapePressed, g_AiModelSelectorFocusSlot

    if (!g_AiModelSelectorActive)
        return

    AiModelSelector_UnbindModalHotkeys()
    AiModelSelector_SafeDestroyGui(g_AiModelSelectorGui)
    g_AiModelSelectorGui := false
    g_AiModelSelectorLv := false

    currentSlot := Handy_GetPersistedAiModelSlot()
    count := HandyAiModels_Count()
    lvH := Min(320, Max(120, 28 + count * 28))

    g_AiModelSelectorGui := Gui("+AlwaysOnTop +ToolWindow", "Handy AI Model")
    g_AiModelSelectorGui.SetFont("s10", "Segoe UI")
    g_AiModelSelectorGui.Add("Text", "w700",
        "Char = select   a/Insert = add   e = edit   Del = remove   Enter/double-click = select   Esc/Space = cancel")
    g_AiModelSelectorLv := g_AiModelSelectorGui.Add("ListView", "w700 h" . lvH . " -Multi", ["Char", "Model",
        "Description", "Status"])
    g_AiModelSelectorLv.OnEvent("DoubleClick", AiModelSelector_OnListActivate)
    g_AiModelSelectorGui.Add("Button", "w100 Section", "Add").OnEvent("Click", AiModelSelector_AddEntry)
    g_AiModelSelectorGui.Add("Button", "w100 ys", "Edit").OnEvent("Click", AiModelSelector_EditEntry)
    g_AiModelSelectorGui.Add("Button", "w100 ys", "Delete").OnEvent("Click", AiModelSelector_DeleteEntry)
    g_AiModelSelectorGui.Add("Button", "w100 ys", "Close").OnEvent("Click", AiModelSelector_Cancel)
    g_AiModelSelectorGui.OnEvent("Close", AiModelSelector_Cancel)
    g_AiModelSelectorGui.OnEvent("Escape", AiModelSelector_Cancel)

    focusRow := 0
    preferredSlot := g_AiModelSelectorFocusSlot
    g_AiModelSelectorFocusSlot := 0
    maxSlots := HandyAiModels_MaxSlots()
    loop maxSlots {
        num := A_Index
        if (!g_HandyAiModels.Has(num))
            continue
        model := g_HandyAiModels[num]
        status := (num = currentSlot) ? "Current" : ""
        label := HandyAiModels_LabelForIndex(num)
        desc := (model.desc != "") ? model.desc : ""
        row := g_AiModelSelectorLv.Add("", label, model.name, desc, status)
        if (preferredSlot > 0 && num = preferredSlot)
            focusRow := row
        else if (preferredSlot < 1 && num = currentSlot)
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

    g_AiModelSelectorGui.Show("AutoSize Hide")
    g_AiModelSelectorGui.GetPos(, , &guiW, &guiH)
    guiX := monitorLeft + (monitorWidth - guiW) // 2
    guiY := monitorTop + (monitorHeight - guiH) // 2
    if (guiX < monitorLeft + 20)
        guiX := monitorLeft + 20
    if (guiY < monitorTop + 20)
        guiY := monitorTop + 20
    g_AiModelSelectorGui.Show("x" . guiX . " y" . guiY)
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
    global g_AiModelHotkeyHandlers, g_HandyAiModels
    AiModelSelector_UnbindModalHotkeys()
    hwnd := AiModelSelector_SelectorHwnd()
    if (!hwnd)
        return
    try HotIfWinActive("ahk_id " hwnd)
    catch {
        return
    }

    maxSlots := HandyAiModels_MaxSlots()
    loop maxSlots {
        slot := A_Index
        if (!g_HandyAiModels.Has(slot))
            continue
        key := HandyAiModels_LabelForIndex(slot)
        cb := AiModelSelector_Select.Bind(slot)
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
    try {
        Hotkey("Space", AiModelSelector_Cancel, "On")
        g_AiModelHotkeyHandlers.Push({ key: "Space", handler: AiModelSelector_Cancel })
    } catch {
    }
    try {
        Hotkey("a", AiModelSelector_AddEntry, "On")
        g_AiModelHotkeyHandlers.Push({ key: "a", handler: AiModelSelector_AddEntry })
    } catch {
    }
    try {
        Hotkey("Insert", AiModelSelector_AddEntry, "On")
        g_AiModelHotkeyHandlers.Push({ key: "Insert", handler: AiModelSelector_AddEntry })
    } catch {
    }
    try {
        Hotkey("e", AiModelSelector_EditEntry, "On")
        g_AiModelHotkeyHandlers.Push({ key: "e", handler: AiModelSelector_EditEntry })
    } catch {
    }
    try {
        Hotkey("Delete", AiModelSelector_DeleteEntry, "On")
        g_AiModelHotkeyHandlers.Push({ key: "Delete", handler: AiModelSelector_DeleteEntry })
    } catch {
    }

    try HotIf()
    catch {
    }
}

AiModelSelector_Select(selection, *) {
    global g_AiModelSelectorActive, g_HandyAiModels, g_AiModelSelectorPrevHwnd
    if (!g_AiModelSelectorActive)
        return
    ; Keep the pre-selector window (dictation paste target) across Close + Handy focus steal.
    restoreHwnd := g_AiModelSelectorPrevHwnd
    AiModelSelector_Close()
    if (g_HandyAiModels.Has(selection))
        ExecuteHandyAiModelSelection(selection, false, restoreHwnd)
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
    return HandyAiModels_IndexFromKey(ch)
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

AiModelSelector_SuspendGuiForInput() {
    global g_AiModelSelectorGui, g_AiModelSelectorLv
    AiModelSelector_UnbindModalHotkeys()
    AiModelSelector_SafeDestroyGui(g_AiModelSelectorGui)
    g_AiModelSelectorGui := false
    g_AiModelSelectorLv := false
}

AiModelSelector_AddEntry(*) {
    global g_AiModelSelectorActive, g_AiModelSelectorFocusSlot

    if (!g_AiModelSelectorActive)
        return
    AiModelSelector_SuspendGuiForInput()

    nameBox := InputBox("Model name match string (substring of Handy UIA name):",
        "Add Handy AI model", "w440 h120")
    if (nameBox.Result != "OK") {
        AiModelSelector_Rebuild()
        return
    }
    name := Trim(nameBox.Value)
    if (name = "") {
        try ShowCenteredOverlay_Utils("⚠ Name cannot be empty.", 2500, BANNER_ACCENT_INTERMEDIATE)
        AiModelSelector_Rebuild()
        return
    }
    if (HandyAiModels_Count() >= HandyAiModels_MaxSlots()) {
        try ShowCenteredOverlay_Utils("❌ Model list is full.", 2500, BANNER_ACCENT_ERROR)
        AiModelSelector_Rebuild()
        return
    }
    if (HandyAiModels_NameExists(name)) {
        try ShowCenteredOverlay_Utils("❌ Duplicate model name.", 2500, BANNER_ACCENT_ERROR)
        AiModelSelector_Rebuild()
        return
    }
    if HandyAiModels_Add(name) {
        g_AiModelSelectorFocusSlot := HandyAiModels_Count()
        try ShowCenteredOverlay_Utils("✅ Added: " . name, 2000, BANNER_ACCENT_SUCCESS)
    } else {
        try ShowCenteredOverlay_Utils("❌ Could not add model.", 2500, BANNER_ACCENT_ERROR)
    }
    AiModelSelector_Rebuild()
}

AiModelSelector_EditEntry(*) {
    global g_AiModelSelectorActive, g_HandyAiModels, g_AiModelSelectorFocusSlot

    if (!g_AiModelSelectorActive)
        return
    slot := AiModelSelector_SelectedSlot()
    if (slot < 1) {
        try ShowCenteredOverlay_Utils("Select a model row to edit.", 2200, BANNER_ACCENT_INTERMEDIATE)
        return
    }
    if (HandyAiModels_IsBuiltin(slot)) {
        try ShowCenteredOverlay_Utils("Built-in models cannot be edited.", 2500, BANNER_ACCENT_INTERMEDIATE)
        return
    }
    if !g_HandyAiModels.Has(slot)
        return

    currentName := g_HandyAiModels[slot].name
    AiModelSelector_SuspendGuiForInput()

    nameBox := InputBox("Model name match string (substring of Handy UIA name):",
        "Edit Handy AI model", "w440 h120", currentName)
    if (nameBox.Result != "OK") {
        g_AiModelSelectorFocusSlot := slot
        AiModelSelector_Rebuild()
        return
    }
    name := Trim(nameBox.Value)
    if (name = "") {
        try ShowCenteredOverlay_Utils("⚠ Name cannot be empty.", 2500, BANNER_ACCENT_INTERMEDIATE)
        g_AiModelSelectorFocusSlot := slot
        AiModelSelector_Rebuild()
        return
    }
    if (name = currentName) {
        g_AiModelSelectorFocusSlot := slot
        AiModelSelector_Rebuild()
        return
    }
    if (HandyAiModels_Rename(slot, name)) {
        g_AiModelSelectorFocusSlot := slot
        try ShowCenteredOverlay_Utils("✅ Renamed: " . name, 2000, BANNER_ACCENT_SUCCESS)
    } else {
        try ShowCenteredOverlay_Utils("❌ Could not rename (duplicate or save failed).", 2500, BANNER_ACCENT_ERROR)
        g_AiModelSelectorFocusSlot := slot
    }
    AiModelSelector_Rebuild()
}

AiModelSelector_DeleteEntry(*) {
    global g_AiModelSelectorActive, g_HandyAiModels, g_AiModelSelectorGui, g_AiModelSelectorFocusSlot

    if (!g_AiModelSelectorActive)
        return
    slot := AiModelSelector_SelectedSlot()
    if (slot < 1) {
        try ShowCenteredOverlay_Utils("Select a model row to delete.", 2200, BANNER_ACCENT_INTERMEDIATE)
        return
    }
    if (HandyAiModels_IsBuiltin(slot)) {
        try ShowCenteredOverlay_Utils("Built-in models cannot be removed.", 2500, BANNER_ACCENT_INTERMEDIATE)
        return
    }
    if !g_HandyAiModels.Has(slot)
        return

    name := g_HandyAiModels[slot].name
    hwnd := AiModelSelector_SelectorHwnd()
    msgOpts := "YesNo Icon! Default2"
    if (hwnd)
        msgOpts .= " Owner" . hwnd

    AiModelSelector_UnbindModalHotkeys()
    confirmed := (MsgBox("Delete model '" . name . "'?", "Delete Handy AI model", msgOpts) = "Yes")
    if (!confirmed) {
        g_AiModelSelectorFocusSlot := slot
        AiModelSelector_Rebuild()
        return
    }
    if (HandyAiModels_Remove(slot)) {
        ; Focus previous extra row when possible.
        g_AiModelSelectorFocusSlot := Max(HANDY_AI_BUILTIN_COUNT + 1, slot - 1)
        if (!g_HandyAiModels.Has(g_AiModelSelectorFocusSlot))
            g_AiModelSelectorFocusSlot := HANDY_AI_SLOT_ENGLISH
        try ShowCenteredOverlay_Utils("✅ Removed: " . name, 2000, BANNER_ACCENT_SUCCESS)
    } else {
        try ShowCenteredOverlay_Utils("❌ Could not remove model.", 2200, BANNER_ACCENT_ERROR)
        g_AiModelSelectorFocusSlot := slot
    }
    AiModelSelector_Rebuild()
}

AiModelSelector_Cancel(*) {
    AiModelSelector_Close()
}

AiModelSelector_Close() {
    global g_AiModelSelectorGui, g_AiModelSelectorActive, g_AiModelSelectorLv, g_AiModelSelectorPrevHwnd
    global g_OnEscapePressed, g_AiModelSelectorFocusSlot

    if (!g_AiModelSelectorActive)
        return

    g_AiModelSelectorActive := false
    g_AiModelSelectorFocusSlot := 0
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

    prevHwnd := g_AiModelSelectorPrevHwnd
    g_AiModelSelectorPrevHwnd := 0
    if (prevHwnd && WinExist("ahk_id " prevHwnd)) {
        try WinActivate("ahk_id " prevHwnd)
        catch {
        }
    }
}
