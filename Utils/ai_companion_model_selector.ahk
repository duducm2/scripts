; =============================================================================
; Utils module: ai_companion_model_selector.ahk
; Shared Shift+L model list manager (Utility Shortcuts ListView aesthetic).
; CRUD mirrors Utility Shortcuts: Insert/a add, E edit, Delete remove.
; Included from Utils.ahk after Lib\AiCompanionModels.ahk.
; =============================================================================

global g_AiCompanionModelSelectorGui := false
global g_AiCompanionModelSelectorLv := false
global g_AiCompanionModelSelectorActive := false
global g_AiCompanionModelSelectorCompanion := ""
global g_AiCompanionModelSelectorLastForegroundMonitorIdx := 0
global g_AiCompanionModelEscPollPrev := false

ShowAiCompanionModelSelector(companion) {
    global g_AiCompanionModelSelectorActive, g_AiCompanionModelSelectorCompanion

    if !AiCompanionModels_IsValidCompanion(companion)
        return
    if (g_AiCompanionModelSelectorActive)
        AiCompanionModelSelector_Close()

    g_AiCompanionModelSelectorCompanion := companion
    g_AiCompanionModelSelectorActive := true
    AiCompanionModelSelector_Rebuild()
}

AiCompanionModelSelector_SafeDestroyGui(gui) {
    if (!IsObject(gui))
        return
    try gui.Destroy()
    catch {
    }
}

AiCompanionModelSelector_GuiHasWindow(gui) {
    if !IsObject(gui)
        return false
    try
        return !!gui.Hwnd
    catch
        return false
}

AiCompanionModelSelector_PositionGui(gui) {
    GetActiveMonitorWorkArea_StandardBar(&ml, &mt, &mr, &mb)
    gui.Show("AutoSize Hide")
    gui.GetPos(, , &gw, &gh)
    cx := ml + ((mr - ml) - gw) // 2
    cy := mt + ((mb - mt) - gh) // 2
    if (cx < ml)
        cx := ml
    if (cy < mt)
        cy := mt
    gui.Show("x" . cx . " y" . cy)
    try {
        if AiCompanionModelSelector_GuiHasWindow(gui)
            WinActivate(gui.Hwnd)
    } catch {
    }
}

AiCompanionModelSelector_UnbindKeys() {
    ; Clear caller #HotIf (Gemini/Enterprise/Copilot) so Off applies to globally registered modal keys.
    try HotIf()
    catch {
    }
    loop 10 {
        dig := String(A_Index - 1)
        try Hotkey(dig, "Off")
        try Hotkey("$*" . dig, "Off")
    }
    loop 26 {
        letter := Chr(96 + A_Index)
        try Hotkey(letter, "Off")
        try Hotkey("$*" . letter, "Off")
    }
    try Hotkey("Escape", "Off")
    catch {
    }
    try Hotkey("$*Escape", "Off")
    catch {
    }
    try Hotkey("Enter", "Off")
    catch {
    }
    try Hotkey("$*Enter", "Off")
    catch {
    }
    try Hotkey("Insert", "Off")
    catch {
    }
    try Hotkey("$*Insert", "Off")
    catch {
    }
    try Hotkey("Delete", "Off")
    catch {
    }
    try Hotkey("$*Delete", "Off")
    catch {
    }
    try HotIf()
    catch {
    }
}

AiCompanionModelSelector_BindKeys(count) {
    ; Register as GLOBAL while modal is open (caller companion #HotIf would block otherwise).
    try HotIf()
    catch {
    }
    loop count {
        label := AiCompanionModels_LabelForIndex(A_Index)
        try Hotkey("$*" . label, AiCompanionModelSelector_HandleKey, "On")
    }
    try Hotkey("$*a", AiCompanionModelSelector_AddEntry, "On")
    try Hotkey("$*Insert", AiCompanionModelSelector_AddEntry, "On")
    try Hotkey("$*e", AiCompanionModelSelector_EditEntry, "On")
    try Hotkey("$*Delete", AiCompanionModelSelector_DeleteEntry, "On")
    try Hotkey("$*f", AiCompanionModelSelector_SetFast, "On")
    try Hotkey("$*d", AiCompanionModelSelector_SetDeep, "On")
    try Hotkey("$*Enter", AiCompanionModelSelector_OnEnter, "On")
    try HotIf()
    catch {
    }
}

AiCompanionModelSelector_BindRobustEscape() {
    global g_AiCompanionModelSelectorGui, g_OnEscapePressed, g_AiCompanionModelEscPollPrev
    SetTimer(AiCompanionModelSelector_EscapePoll, 0)
    if (!AiCompanionModelSelector_GuiHasWindow(g_AiCompanionModelSelectorGui))
        return
    try HotIf()
    catch {
    }
    try Hotkey("$*Escape", AiCompanionModelSelector_EscapeFromHotkey, "On")
    catch {
    }
    try HotIf()
    catch {
    }
    g_OnEscapePressed := AiCompanionModelSelector_GlobalEscapeCallback
    Utils_EnsureGlobalEscapeHotkey()
    g_AiCompanionModelEscPollPrev := false
    SetTimer(AiCompanionModelSelector_EscapePoll, 50)
}

AiCompanionModelSelector_UnbindRobustEscape() {
    global g_OnEscapePressed, g_AiCompanionModelEscPollPrev
    SetTimer(AiCompanionModelSelector_EscapePoll, 0)
    g_AiCompanionModelEscPollPrev := false
    try HotIf()
    catch {
    }
    try Hotkey("$*Escape", AiCompanionModelSelector_EscapeFromHotkey, "Off")
    catch {
    }
    try HotIf()
    catch {
    }
    g_OnEscapePressed := ""
    Utils_EnsureGlobalEscapeHotkey()
}

AiCompanionModelSelector_EscapeFromHotkey(*) {
    AiCompanionModelSelector_Cancel()
}

AiCompanionModelSelector_GlobalEscapeCallback(*) {
    AiCompanionModelSelector_Cancel()
}

AiCompanionModelSelector_EscapePoll() {
    global g_AiCompanionModelSelectorActive, g_AiCompanionModelEscPollPrev
    if (!g_AiCompanionModelSelectorActive) {
        SetTimer(AiCompanionModelSelector_EscapePoll, 0)
        return
    }
    escSync := GetKeyState("Escape", "P")
    escAsync := (DllCall("user32\GetAsyncKeyState", "int", 0x1B) & 0x8000) != 0
    escDown := escSync || escAsync
    if (escDown) {
        if (!g_AiCompanionModelEscPollPrev) {
            g_AiCompanionModelEscPollPrev := true
            AiCompanionModelSelector_Cancel()
        }
    } else {
        g_AiCompanionModelEscPollPrev := false
    }
}

AiCompanionModelSelector_GuiEscape(*) {
    AiCompanionModelSelector_Cancel()
}

AiCompanionModelSelector_StopMonitorTracking() {
    try SetTimer(AiCompanionModelSelector_TrackActiveMonitorTick, 0)
}

AiCompanionModelSelector_TrackActiveMonitorTick() {
    global g_AiCompanionModelSelectorGui, g_AiCompanionModelSelectorActive,
        g_AiCompanionModelSelectorLastForegroundMonitorIdx
    try {
        if (!g_AiCompanionModelSelectorActive || !AiCompanionModelSelector_GuiHasWindow(g_AiCompanionModelSelectorGui)) {
            AiCompanionModelSelector_StopMonitorTracking()
            if (g_AiCompanionModelSelectorActive)
                AiCompanionModelSelector_ForceReset()
            return
        }
        newIdx := GetMonitorIndexForForeground_StandardBar()
        if (newIdx != g_AiCompanionModelSelectorLastForegroundMonitorIdx) {
            MonitorGetWorkArea(newIdx, &ml, &mt, &mr, &mb)
            g_AiCompanionModelSelectorGui.Show("AutoSize Hide")
            g_AiCompanionModelSelectorGui.GetPos(, , &gw, &gh)
            cx := ml + ((mr - ml) - gw) // 2
            cy := mt + ((mb - mt) - gh) // 2
            g_AiCompanionModelSelectorGui.Show("x" . cx . " y" . cy)
            g_AiCompanionModelSelectorLastForegroundMonitorIdx := newIdx
        }
    } catch {
        AiCompanionModelSelector_StopMonitorTracking()
        try AiCompanionModelSelector_ForceReset()
    }
}

AiCompanionModelSelector_Rebuild() {
    global g_AiCompanionModelSelectorGui, g_AiCompanionModelSelectorLv, g_AiCompanionModelSelectorActive,
        g_AiCompanionModelSelectorCompanion, g_AiCompanionModelSelectorLastForegroundMonitorIdx

    if (!g_AiCompanionModelSelectorActive)
        return

    AiCompanionModelSelector_StopMonitorTracking()
    AiCompanionModelSelector_UnbindKeys()
    AiCompanionModelSelector_SafeDestroyGui(g_AiCompanionModelSelectorGui)
    g_AiCompanionModelSelectorGui := false
    g_AiCompanionModelSelectorLv := false

    companion := g_AiCompanionModelSelectorCompanion
    cfg := AiCompanionModels_Load(companion)
    title := AiCompanionModels_DisplayName(companion)
    fastLabel := (cfg.fast != "") ? cfg.fast : "(not set)"
    deepLabel := (cfg.deep != "") ? cfg.deep : "(not set)"

    hint :=
        "Char = select   Enter/double-click = select   Insert/a = add   E = edit   Delete = remove   f Fast   d Deep   Esc = cancel"

    g_AiCompanionModelSelectorGui := Gui("+AlwaysOnTop +ToolWindow", title . " models")
    g_AiCompanionModelSelectorGui.SetFont("s10", "Segoe UI")
    g_AiCompanionModelSelectorGui.Add("Text", "w700", hint)
    g_AiCompanionModelSelectorLv := g_AiCompanionModelSelectorGui.Add("ListView", "w700 h280 -Multi", ["Char",
        "Model", "Detail"])
    g_AiCompanionModelSelectorLv.OnEvent("DoubleClick", AiCompanionModelSelector_OnListActivate)
    g_AiCompanionModelSelectorGui.Add("Button", "w100 Section", "Add").OnEvent("Click",
        AiCompanionModelSelector_AddEntry)
    g_AiCompanionModelSelectorGui.Add("Button", "w100 ys", "Edit").OnEvent("Click", AiCompanionModelSelector_EditEntry)
    g_AiCompanionModelSelectorGui.Add("Button", "w100 ys", "Delete").OnEvent("Click",
        AiCompanionModelSelector_DeleteEntry)
    g_AiCompanionModelSelectorGui.Add("Button", "w100 ys", "Close").OnEvent("Click", AiCompanionModelSelector_Cancel)
    g_AiCompanionModelSelectorGui.OnEvent("Close", AiCompanionModelSelector_Cancel)
    g_AiCompanionModelSelectorGui.OnEvent("Escape", AiCompanionModelSelector_GuiEscape)

    g_AiCompanionModelSelectorLv.Add("", "f", "Set Fast", fastLabel)
    g_AiCompanionModelSelectorLv.Add("", "d", "Set Deep", deepLabel)

    maxSlots := AiCompanionModels_MaxSlots()
    if (cfg.models.Length = 0) {
        g_AiCompanionModelSelectorLv.Add("", "", "(no extra models)", "Insert or a to add")
    } else {
        for i, name in cfg.models {
            if (i > maxSlots)
                break
            label := AiCompanionModels_LabelForIndex(i)
            g_AiCompanionModelSelectorLv.Add("", label, name, "")
        }
    }

    try g_AiCompanionModelSelectorLv.ModifyCol(1, 50)
    try g_AiCompanionModelSelectorLv.ModifyCol(2, 280)
    try g_AiCompanionModelSelectorLv.ModifyCol(3, 340)
    focusRow := (cfg.models.Length > 0) ? 3 : 1
    if (g_AiCompanionModelSelectorLv.GetCount() > 0) {
        try g_AiCompanionModelSelectorLv.Modify(focusRow, "Select Focus Vis")
        catch {
        }
    }

    AiCompanionModelSelector_PositionGui(g_AiCompanionModelSelectorGui)
    try g_AiCompanionModelSelectorLv.Focus()
    catch {
    }
    g_AiCompanionModelSelectorLastForegroundMonitorIdx := GetMonitorIndexForForeground_StandardBar()

    count := Min(cfg.models.Length, maxSlots)
    AiCompanionModelSelector_BindKeys(count)
    AiCompanionModelSelector_BindRobustEscape()
    if (g_AiCompanionModelSelectorActive)
        SetTimer(AiCompanionModelSelector_TrackActiveMonitorTick, 115)
}

; Returns { kind: "fast"|"deep"|"model"|"", index: 0|n, name: "" }.
AiCompanionModelSelector_SelectedTarget() {
    global g_AiCompanionModelSelectorLv, g_AiCompanionModelSelectorCompanion
    out := { kind: "", index: 0, name: "" }
    if (!IsObject(g_AiCompanionModelSelectorLv))
        return out
    row := 0
    try row := g_AiCompanionModelSelectorLv.GetNext(0, "Focused")
    catch {
        row := 0
    }
    if (row < 1) {
        try row := g_AiCompanionModelSelectorLv.GetNext(0, "Selected")
        catch {
            row := 0
        }
    }
    if (row < 1)
        return out
    ch := ""
    name := ""
    try ch := StrLower(Trim(g_AiCompanionModelSelectorLv.GetText(row, 1)))
    catch {
        return out
    }
    try name := Trim(g_AiCompanionModelSelectorLv.GetText(row, 2))
    catch {
    }
    if (ch = "f")
        return { kind: "fast", index: 0, name: name }
    if (ch = "d")
        return { kind: "deep", index: 0, name: name }
    idx := AiCompanionModels_IndexFromKey(ch)
    if (idx < 1)
        return out
    models := AiCompanionModels_GetModels(g_AiCompanionModelSelectorCompanion)
    if (idx > models.Length)
        return out
    return { kind: "model", index: idx, name: models[idx] }
}

AiCompanionModelSelector_SelectedChar() {
    t := AiCompanionModelSelector_SelectedTarget()
    if (t.kind = "fast")
        return "f"
    if (t.kind = "deep")
        return "d"
    if (t.kind = "model")
        return AiCompanionModels_LabelForIndex(t.index)
    return ""
}

AiCompanionModelSelector_OnEnter(*) {
    global g_AiCompanionModelSelectorActive
    if (!g_AiCompanionModelSelectorActive)
        return
    ch := AiCompanionModelSelector_SelectedChar()
    if (ch = "")
        return
    if (ch = "f") {
        AiCompanionModelSelector_SetFast()
        return
    }
    if (ch = "d") {
        AiCompanionModelSelector_SetDeep()
        return
    }
    AiCompanionModelSelector_HandleKey(ch)
}

AiCompanionModelSelector_OnListActivate(*) {
    AiCompanionModelSelector_OnEnter()
}

AiCompanionModelSelector_SuspendGuiForInput() {
    global g_AiCompanionModelSelectorGui, g_AiCompanionModelSelectorLv
    ; AlwaysOnTop modal covers AHK InputBox unless we tear it down first.
    AiCompanionModelSelector_StopMonitorTracking()
    AiCompanionModelSelector_SafeDestroyGui(g_AiCompanionModelSelectorGui)
    g_AiCompanionModelSelectorGui := false
    g_AiCompanionModelSelectorLv := false
}

AiCompanionModelSelector_AddEntry(*) {
    global g_AiCompanionModelSelectorActive, g_AiCompanionModelSelectorCompanion

    if (!g_AiCompanionModelSelectorActive)
        return
    AiCompanionModelSelector_UnbindKeys()
    AiCompanionModelSelector_UnbindRobustEscape()
    AiCompanionModelSelector_SuspendGuiForInput()

    companion := g_AiCompanionModelSelectorCompanion
    nameBox := InputBox("Exact UIA-visible model name:", "Add AI model — " .
        AiCompanionModels_DisplayName(companion), "w440 h120")
    if (nameBox.Result != "OK") {
        AiCompanionModelSelector_Rebuild()
        return
    }
    name := Trim(nameBox.Value)
    if (name = "") {
        try ShowCenteredOverlay_Utils("⚠ Name cannot be empty.", 2500, BANNER_ACCENT_INTERMEDIATE)
        AiCompanionModelSelector_Rebuild()
        return
    }

    cfg := AiCompanionModels_Load(companion)
    if (cfg.models.Length >= AiCompanionModels_MaxSlots()) {
        try ShowCenteredOverlay_Utils("❌ Model list is full.", 2500, BANNER_ACCENT_ERROR)
        AiCompanionModelSelector_Rebuild()
        return
    }

    if AiCompanionModels_AddModel(companion, name, "")
        try ShowCenteredOverlay_Utils("✅ Added: " . name, 2000, BANNER_ACCENT_SUCCESS)
    AiCompanionModelSelector_Rebuild()
}

AiCompanionModelSelector_EditEntry(*) {
    global g_AiCompanionModelSelectorActive, g_AiCompanionModelSelectorCompanion

    if (!g_AiCompanionModelSelectorActive)
        return
    t := AiCompanionModelSelector_SelectedTarget()
    if (t.kind = "fast") {
        AiCompanionModelSelector_SetRoleFromInput("fast")
        return
    }
    if (t.kind = "deep") {
        AiCompanionModelSelector_SetRoleFromInput("deep")
        return
    }
    if (t.kind != "model" || t.index < 1) {
        try ShowCenteredOverlay_Utils("Select a model row to edit.", 2200, BANNER_ACCENT_INTERMEDIATE)
        return
    }

    companion := g_AiCompanionModelSelectorCompanion
    AiCompanionModelSelector_UnbindKeys()
    AiCompanionModelSelector_UnbindRobustEscape()
    AiCompanionModelSelector_SuspendGuiForInput()

    nameBox := InputBox("Exact UIA-visible model name:", "Edit AI model — " .
        AiCompanionModels_DisplayName(companion), "w440 h120", t.name)
    if (nameBox.Result != "OK") {
        AiCompanionModelSelector_Rebuild()
        return
    }
    name := Trim(nameBox.Value)
    if (name = "") {
        try ShowCenteredOverlay_Utils("⚠ Name cannot be empty.", 2500, BANNER_ACCENT_INTERMEDIATE)
        AiCompanionModelSelector_Rebuild()
        return
    }
    if (name = t.name) {
        AiCompanionModelSelector_Rebuild()
        return
    }
    if (AiCompanionModels_RenameModel(companion, t.index, name)) {
        try ShowCenteredOverlay_Utils("✅ Renamed: " . name, 2000, BANNER_ACCENT_SUCCESS)
    } else {
        try ShowCenteredOverlay_Utils("❌ Could not rename (duplicate or save failed).", 2500, BANNER_ACCENT_ERROR)
    }
    AiCompanionModelSelector_Rebuild()
}

AiCompanionModelSelector_DeleteEntry(*) {
    global g_AiCompanionModelSelectorActive, g_AiCompanionModelSelectorCompanion, g_AiCompanionModelSelectorGui

    if (!g_AiCompanionModelSelectorActive)
        return
    t := AiCompanionModelSelector_SelectedTarget()
    if (t.kind != "model" || t.index < 1) {
        try ShowCenteredOverlay_Utils("Select a model row to delete.", 2200, BANNER_ACCENT_INTERMEDIATE)
        return
    }

    companion := g_AiCompanionModelSelectorCompanion
    hwnd := 0
    try {
        if AiCompanionModelSelector_GuiHasWindow(g_AiCompanionModelSelectorGui)
            hwnd := g_AiCompanionModelSelectorGui.Hwnd
    } catch {
    }
    msgOpts := "YesNo Icon! Default2"
    if (hwnd)
        msgOpts .= " Owner" . hwnd

    AiCompanionModelSelector_UnbindKeys()
    AiCompanionModelSelector_UnbindRobustEscape()
    confirmed := (MsgBox("Delete model '" . t.name . "'?", "Delete AI model", msgOpts) = "Yes")
    if (!confirmed) {
        AiCompanionModelSelector_Rebuild()
        return
    }
    if (AiCompanionModels_RemoveModel(companion, t.index)) {
        try ShowCenteredOverlay_Utils("✅ Removed: " . t.name, 2000, BANNER_ACCENT_SUCCESS)
    } else {
        try ShowCenteredOverlay_Utils("❌ Could not remove model.", 2200, BANNER_ACCENT_ERROR)
    }
    AiCompanionModelSelector_Rebuild()
}

AiCompanionModelSelector_SetFast(*) {
    AiCompanionModelSelector_SetRoleFromInput("fast")
}

AiCompanionModelSelector_SetDeep(*) {
    AiCompanionModelSelector_SetRoleFromInput("deep")
}

AiCompanionModelSelector_SetRoleFromInput(role) {
    global g_AiCompanionModelSelectorActive, g_AiCompanionModelSelectorCompanion

    if (!g_AiCompanionModelSelectorActive)
        return

    role := StrLower(Trim(role))
    if (role != "fast" && role != "deep")
        return

    companion := g_AiCompanionModelSelectorCompanion
    AiCompanionModelSelector_UnbindKeys()
    AiCompanionModelSelector_UnbindRobustEscape()
    AiCompanionModelSelector_SuspendGuiForInput()

    current := (role = "fast") ? AiCompanionModels_GetFast(companion) : AiCompanionModels_GetDeep(companion)
    roleLabel := (role = "fast") ? "Fast" : "Deep"
    nameBox := InputBox(
        "Exact UIA-visible " . roleLabel . " model name:`n(Shift+Q / Shift+M use this value)",
        "Set " . roleLabel . " — " . AiCompanionModels_DisplayName(companion),
        "w460 h140",
        current)
    if (nameBox.Result != "OK") {
        AiCompanionModelSelector_Rebuild()
        return
    }
    name := Trim(nameBox.Value)
    if (name = "") {
        try ShowCenteredOverlay_Utils("⚠ Name cannot be empty.", 2500, BANNER_ACCENT_INTERMEDIATE)
        AiCompanionModelSelector_Rebuild()
        return
    }

    if (AiCompanionModels_SetRole(companion, role, name)) {
        try ShowCenteredOverlay_Utils("✅ " . roleLabel . " set: " . name, 2000, BANNER_ACCENT_SUCCESS)
    } else {
        try ShowCenteredOverlay_Utils("❌ Could not save " . roleLabel . " name", 2200, BANNER_ACCENT_ERROR)
    }
    AiCompanionModelSelector_Rebuild()
}

AiCompanionModelSelector_HandleKey(thisHotkey := "") {
    global g_AiCompanionModelSelectorActive, g_AiCompanionModelSelectorCompanion

    if (!g_AiCompanionModelSelectorActive)
        return

    key := thisHotkey
    if (key = "")
        key := A_ThisHotkey
    key := RegExReplace(key, "^[\$\*]*", "")
    key := StrLower(key)

    idx := AiCompanionModels_IndexFromKey(key)
    if (idx < 1)
        return

    companion := g_AiCompanionModelSelectorCompanion
    models := AiCompanionModels_GetModels(companion)
    if (idx > models.Length)
        return

    modelName := models[idx]
    AiCompanionModelSelector_Close()
    ok := AiCompanionModels_Apply(companion, modelName)
    if (!ok)
        try ShowCenteredOverlay_Utils("Could not select " . modelName, 2200, BANNER_ACCENT_ERROR)
}

AiCompanionModelSelector_Cancel(*) {
    AiCompanionModelSelector_Close()
}

AiCompanionModelSelector_ForceReset() {
    global g_AiCompanionModelSelectorGui, g_AiCompanionModelSelectorLv, g_AiCompanionModelSelectorActive,
        g_AiCompanionModelSelectorCompanion, g_AiCompanionModelSelectorLastForegroundMonitorIdx

    g_AiCompanionModelSelectorActive := false
    g_AiCompanionModelSelectorCompanion := ""
    g_AiCompanionModelSelectorLastForegroundMonitorIdx := 0

    try AiCompanionModelSelector_StopMonitorTracking()
    try AiCompanionModelSelector_UnbindRobustEscape()
    try AiCompanionModelSelector_UnbindKeys()
    try Utils_EnsureGlobalEscapeHotkey()
    catch {
    }
    try AiCompanionModelSelector_SafeDestroyGui(g_AiCompanionModelSelectorGui)
    g_AiCompanionModelSelectorGui := false
    g_AiCompanionModelSelectorLv := false
}

AiCompanionModelSelector_Close() {
    global g_AiCompanionModelSelectorActive
    if (!g_AiCompanionModelSelectorActive)
        return
    AiCompanionModelSelector_ForceReset()
}
