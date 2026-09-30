; =============================================================================
; Utils module: ai_companion_model_selector.ahk
; Shared Shift+L model list manager (Utility Shortcuts ListView aesthetic).
; CRUD mirrors Utility Shortcuts: Insert/a add, E edit, Delete remove.
; Button rows: Shift+N / Shift+D / Shift+S capture New Chat / Menu / Search.
; Included from Utils.ahk after Lib\AiCompanionModels.ahk and AiCompanionButtons.ahk.
; =============================================================================

global g_AiCompanionModelSelectorGui := false
global g_AiCompanionModelSelectorLv := false
global g_AiCompanionModelSelectorActive := false
global g_AiCompanionModelSelectorCompanion := ""
global g_AiCompanionModelSelectorTargetHwnd := 0
global g_AiCompanionModelSelectorLastForegroundMonitorIdx := 0
global g_AiCompanionModelSelectorRowMeta := Map()
global g_AiCompanionModelEscPollPrev := false
global g_AiCompanionButtonCaptureActive := false
global g_AiCompanionButtonCaptureAction := ""
global g_AiCompanionButtonCaptureArmed := false
global g_AiCompanionButtonCaptureEscLatch := false

ShowAiCompanionModelSelector(companion) {
    global g_AiCompanionModelSelectorActive, g_AiCompanionModelSelectorCompanion,
        g_AiCompanionModelSelectorTargetHwnd

    if !AiCompanionModels_IsValidCompanion(companion)
        return
    ; Remember the companion window before the menu takes focus.
    priorHwnd := g_AiCompanionModelSelectorActive ? g_AiCompanionModelSelectorTargetHwnd : WinExist("A")
    if (g_AiCompanionModelSelectorActive)
        AiCompanionModelSelector_Close()

    g_AiCompanionModelSelectorTargetHwnd := priorHwnd
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
    for action in AiCompanionButtons_Actions() {
        try Hotkey("$*+" . action.chord, "Off")
        catch {
        }
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
    for action in AiCompanionButtons_Actions() {
        try Hotkey("$*+" . action.chord, AiCompanionModelSelector_CaptureChord, "On")
    }
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
    global g_AiCompanionButtonCaptureActive, g_AiCompanionButtonCaptureEscLatch
    if (g_AiCompanionButtonCaptureActive)
        AiCompanionButtonCapture_Abort()
    else if (!g_AiCompanionButtonCaptureEscLatch)
        AiCompanionModelSelector_Cancel()
}

AiCompanionModelSelector_GlobalEscapeCallback(*) {
    global g_AiCompanionButtonCaptureActive, g_AiCompanionButtonCaptureEscLatch
    if (g_AiCompanionButtonCaptureActive)
        AiCompanionButtonCapture_Abort()
    else if (!g_AiCompanionButtonCaptureEscLatch)
        AiCompanionModelSelector_Cancel()
}

AiCompanionModelSelector_EscapePoll() {
    global g_AiCompanionModelSelectorActive, g_AiCompanionModelEscPollPrev, g_AiCompanionButtonCaptureActive,
        g_AiCompanionButtonCaptureEscLatch
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
            if (g_AiCompanionButtonCaptureActive)
                AiCompanionButtonCapture_Abort()
            else if (!g_AiCompanionButtonCaptureEscLatch)
                AiCompanionModelSelector_Cancel()
        }
    } else {
        g_AiCompanionModelEscPollPrev := false
        g_AiCompanionButtonCaptureEscLatch := false
    }
}

AiCompanionModelSelector_GuiEscape(*) {
    global g_AiCompanionButtonCaptureActive, g_AiCompanionButtonCaptureEscLatch
    if (g_AiCompanionButtonCaptureActive)
        AiCompanionButtonCapture_Abort()
    else if (!g_AiCompanionButtonCaptureEscLatch)
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
        g_AiCompanionModelSelectorCompanion, g_AiCompanionModelSelectorLastForegroundMonitorIdx,
        g_AiCompanionButtonCaptureActive, g_AiCompanionModelSelectorRowMeta

    if (g_AiCompanionButtonCaptureActive) {
        AiCompanionButtonCapture_Stop()
        try StandardLoadingBar_Hide(0)
    }

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
        "Char/Enter = select model   Shift+N New Chat   Shift+D Menu   Shift+S Search   Insert/a add   E edit   Delete remove   f Fast   d Deep   Esc cancel"

    g_AiCompanionModelSelectorGui := Gui("+AlwaysOnTop +ToolWindow", title . " models")
    g_AiCompanionModelSelectorGui.SetFont("s10", "Segoe UI")
    g_AiCompanionModelSelectorGui.Add("Text", "w700 Wrap", hint)
    g_AiCompanionModelSelectorLv := g_AiCompanionModelSelectorGui.Add("ListView", "w700 h340 -Multi", ["Char",
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

    g_AiCompanionModelSelectorRowMeta := Map()
    for action in AiCompanionButtons_Actions() {
        spec := AiCompanionButtons_Get(companion, action.id)
        row := g_AiCompanionModelSelectorLv.Add("", StrUpper(action.chord), action.label,
        AiCompanionButtons_Summary(spec))
        g_AiCompanionModelSelectorRowMeta[row] := action.id
    }
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
    ; Button rows occupy 1-3. Fast is row 4; the first extra model is row 6.
    focusRow := (cfg.models.Length > 0) ? 6 : 4
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

; Returns { kind: "fast"|"deep"|"model"|"button"|"", index: 0|n, name: "", actionId: "" }.
AiCompanionModelSelector_SelectedTarget() {
    global g_AiCompanionModelSelectorLv, g_AiCompanionModelSelectorCompanion, g_AiCompanionModelSelectorRowMeta
    out := { kind: "", index: 0, name: "", actionId: "" }
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
    if (IsObject(g_AiCompanionModelSelectorRowMeta) && g_AiCompanionModelSelectorRowMeta.Has(row))
        return { kind: "button", index: 0, name: name, actionId: g_AiCompanionModelSelectorRowMeta[row] }
    if (ch = "f")
        return { kind: "fast", index: 0, name: name, actionId: "" }
    if (ch = "d")
        return { kind: "deep", index: 0, name: name, actionId: "" }
    idx := AiCompanionModels_IndexFromKey(ch)
    if (idx < 1)
        return out
    models := AiCompanionModels_GetModels(g_AiCompanionModelSelectorCompanion)
    if (idx > models.Length)
        return out
    return { kind: "model", index: idx, name: models[idx], actionId: "" }
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
    t := AiCompanionModelSelector_SelectedTarget()
    if (t.kind = "button") {
        AiCompanionModelSelector_BeginCapture(t.actionId)
        return
    }
    if (t.kind = "fast") {
        AiCompanionModelSelector_SetFast()
        return
    }
    if (t.kind = "deep") {
        AiCompanionModelSelector_SetDeep()
        return
    }
    if (t.kind != "model")
        return
    AiCompanionModelSelector_HandleKey(AiCompanionModels_LabelForIndex(t.index))
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
    if (t.kind = "button") {
        AiCompanionModelSelector_BeginCapture(t.actionId)
        return
    }
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
    if (t.kind = "button") {
        AiCompanionModelSelector_ClearButton(t)
        return
    }
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
        g_AiCompanionModelSelectorCompanion, g_AiCompanionModelSelectorTargetHwnd,
        g_AiCompanionModelSelectorLastForegroundMonitorIdx, g_AiCompanionButtonCaptureActive,
        g_AiCompanionButtonCaptureEscLatch, g_AiCompanionModelSelectorRowMeta

    wasCapture := g_AiCompanionButtonCaptureActive
    g_AiCompanionButtonCaptureEscLatch := false
    g_AiCompanionModelSelectorActive := false
    g_AiCompanionModelSelectorCompanion := ""
    g_AiCompanionModelSelectorTargetHwnd := 0
    g_AiCompanionModelSelectorLastForegroundMonitorIdx := 0
    g_AiCompanionModelSelectorRowMeta := Map()
    AiCompanionButtonCapture_Stop()
    if (wasCapture) {
        try StandardLoadingBar_Hide(0)
    }

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

AiCompanionModelSelector_CaptureChord(*) {
    key := A_ThisHotkey
    key := RegExReplace(key, "^[\$\*]*", "")
    key := RegExReplace(key, "^\+", "")
    key := StrLower(key)
    action := AiCompanionButtons_ActionByChord(key)
    if (!IsObject(action))
        return
    AiCompanionModelSelector_BeginCapture(action.id)
}

AiCompanionModelSelector_BeginCapture(actionId) {
    global g_AiCompanionModelSelectorActive, g_AiCompanionModelSelectorCompanion,
        g_AiCompanionModelSelectorTargetHwnd, g_AiCompanionModelSelectorGui, g_AiCompanionModelSelectorLv,
        g_AiCompanionButtonCaptureActive, g_AiCompanionButtonCaptureAction, g_AiCompanionButtonCaptureArmed

    if (!g_AiCompanionModelSelectorActive || g_AiCompanionButtonCaptureActive)
        return
    action := AiCompanionButtons_ActionById(actionId)
    if (!IsObject(action))
        return

    hwnd := g_AiCompanionModelSelectorTargetHwnd
    companion := g_AiCompanionModelSelectorCompanion
    AiCompanionModelSelector_StopMonitorTracking()
    AiCompanionModelSelector_UnbindKeys()
    ; UnbindKeys turns Escape off. Keep it so Esc aborts capture.
    try HotIf()
    catch {
    }
    try Hotkey("$*Escape", AiCompanionModelSelector_EscapeFromHotkey, "On")
    catch {
    }
    try HotIf()
    catch {
    }
    AiCompanionModelSelector_SafeDestroyGui(g_AiCompanionModelSelectorGui)
    g_AiCompanionModelSelectorGui := false
    g_AiCompanionModelSelectorLv := false

    if (!hwnd || !WinExist("ahk_id " hwnd)) {
        try ShowCenteredOverlay_Utils("Companion window is not available.", 2200, BANNER_ACCENT_ERROR)
        AiCompanionModelSelector_Rebuild()
        return
    }

    g_AiCompanionButtonCaptureActive := true
    g_AiCompanionButtonCaptureAction := action.id
    g_AiCompanionButtonCaptureArmed := false
    try WinActivate("ahk_id " hwnd)
    title := AiCompanionModels_DisplayName(companion)
    try StandardLoadingBar_Show(
        "Click " . action.label . " in " . title . ". Esc cancels.",
        BANNER_ACCENT_INTERMEDIATE, { passive: true, centerOnHwnd: hwnd, fontSize: 17, textWidth: 640, passiveBgColor: BANNER_ACCENT_INTERMEDIATE }
    )
    SetTimer(AiCompanionButtonCapture_Poll, 30)
}

AiCompanionModelSelector_ClearButton(t) {
    global g_AiCompanionModelSelectorActive, g_AiCompanionModelSelectorCompanion, g_AiCompanionModelSelectorGui

    if (!g_AiCompanionModelSelectorActive || t.kind != "button")
        return
    companion := g_AiCompanionModelSelectorCompanion
    action := AiCompanionButtons_ActionById(t.actionId)
    if (!IsObject(action))
        return
    if (!AiCompanionButtons_HasSpec(AiCompanionButtons_Get(companion, action.id))) {
        try ShowCenteredOverlay_Utils(action.label . " is already using the built-in finder.", 2200,
            BANNER_ACCENT_INTERMEDIATE)
        return
    }

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
    confirmed := (MsgBox("Clear saved " . action.label . " button and use the built-in finder?",
        "Clear button mapping", msgOpts) = "Yes")
    if (!confirmed) {
        AiCompanionModelSelector_Rebuild()
        return
    }
    if (AiCompanionButtons_Clear(companion, action.id)) {
        try ShowCenteredOverlay_Utils("Cleared " . action.label . ". Using the built-in finder.", 2000,
            BANNER_ACCENT_SUCCESS)
    } else {
        try ShowCenteredOverlay_Utils("Could not clear " . action.label . ".", 2200, BANNER_ACCENT_ERROR)
    }
    AiCompanionModelSelector_Rebuild()
}

AiCompanionButtonCapture_Stop() {
    global g_AiCompanionButtonCaptureActive, g_AiCompanionButtonCaptureAction, g_AiCompanionButtonCaptureArmed
    g_AiCompanionButtonCaptureActive := false
    g_AiCompanionButtonCaptureAction := ""
    g_AiCompanionButtonCaptureArmed := false
    try SetTimer(AiCompanionButtonCapture_Poll, 0)
}

AiCompanionButtonCapture_Abort() {
    global g_AiCompanionButtonCaptureActive, g_AiCompanionButtonCaptureEscLatch,
        g_AiCompanionModelSelectorActive, g_AiCompanionModelEscPollPrev
    if (!g_AiCompanionButtonCaptureActive)
        return
    g_AiCompanionButtonCaptureEscLatch := true
    AiCompanionButtonCapture_Stop()
    try StandardLoadingBar_Hide(0)
    if (g_AiCompanionModelSelectorActive)
        AiCompanionModelSelector_Rebuild()
    ; Esc is still down; don't let the restored menu treat that as a close.
    escDown := GetKeyState("Escape", "P") || ((DllCall("user32\GetAsyncKeyState", "int", 0x1B) & 0x8000) != 0)
    if (escDown)
        g_AiCompanionModelEscPollPrev := true
}

AiCompanionButtonCapture_Poll() {
    global g_AiCompanionButtonCaptureActive, g_AiCompanionButtonCaptureArmed
    if (!g_AiCompanionButtonCaptureActive) {
        SetTimer(AiCompanionButtonCapture_Poll, 0)
        return
    }
    down := GetKeyState("LButton", "P") || ((DllCall("user32\GetAsyncKeyState", "int", 0x01) & 0x8000) != 0)
    if (!g_AiCompanionButtonCaptureArmed) {
        if (!down)
            g_AiCompanionButtonCaptureArmed := true
        return
    }
    if (down) {
        AiCompanionButtonCapture_OnClick()
    }
}

AiCompanionButtonCapture_CursorInTarget(hwnd) {
    prevMode := A_CoordModeMouse
    CoordMode("Mouse", "Screen")
    under := 0
    try MouseGetPos(, , &under)
    CoordMode("Mouse", prevMode)
    if (!under)
        return false
    if (under = hwnd)
        return true
    root := 0
    try root := DllCall("GetAncestor", "Ptr", under, "UInt", 2, "Ptr")
    return root = hwnd
}

AiCompanionButtonCapture_OnClick() {
    global g_AiCompanionButtonCaptureActive, g_AiCompanionButtonCaptureAction,
        g_AiCompanionModelSelectorActive, g_AiCompanionModelSelectorCompanion,
        g_AiCompanionModelSelectorTargetHwnd

    if (!g_AiCompanionButtonCaptureActive)
        return
    actionId := g_AiCompanionButtonCaptureAction
    companion := g_AiCompanionModelSelectorCompanion
    hwnd := g_AiCompanionModelSelectorTargetHwnd
    AiCompanionButtonCapture_Stop()
    action := AiCompanionButtons_ActionById(actionId)
    label := IsObject(action) ? action.label : actionId
    title := AiCompanionModels_DisplayName(companion)

    if (!AiCompanionButtonCapture_CursorInTarget(hwnd)) {
        try StandardLoadingBar_Hide(0)
        try ShowCenteredOverlay_Utils("Click was outside " . title . ". Mapping unchanged.", 2200,
            BANNER_ACCENT_INTERMEDIATE)
        if (g_AiCompanionModelSelectorActive)
            AiCompanionModelSelector_Rebuild()
        return
    }

    el := AiCompanionButtons_ButtonFromPoint(hwnd)
    try StandardLoadingBar_Hide(0)
    if (!IsObject(el)) {
        try ShowCenteredOverlay_Utils("Not a button in " . title . ". Mapping unchanged.", 2200,
            BANNER_ACCENT_INTERMEDIATE)
        if (g_AiCompanionModelSelectorActive)
            AiCompanionModelSelector_Rebuild()
        return
    }

    spec := AiCompanionButtons_SpecFromElement(el)
    if (!AiCompanionButtons_HasSpec(spec) || !AiCompanionButtons_Save(companion, actionId, spec)) {
        try ShowCenteredOverlay_Utils("Could not save " . label . ".", 2200, BANNER_ACCENT_ERROR)
        if (g_AiCompanionModelSelectorActive)
            AiCompanionModelSelector_Rebuild()
        return
    }
    try ShowCenteredOverlay_Utils("Saved " . label . ": " . AiCompanionButtons_Summary(spec), 2000,
    BANNER_ACCENT_SUCCESS)
    if (g_AiCompanionModelSelectorActive)
        AiCompanionModelSelector_Rebuild()
}
