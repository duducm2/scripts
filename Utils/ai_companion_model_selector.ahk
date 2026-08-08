; =============================================================================
; Utils module: ai_companion_model_selector.ahk
; Shared Shift+L model list manager (study-topic style: 1-9/letters + a/r/f/d).
; Included from Utils.ahk after Lib\AiCompanionModels.ahk.
; =============================================================================

global g_AiCompanionModelSelectorGui := false
global g_AiCompanionModelSelectorActive := false
global g_AiCompanionModelSelectorCompanion := ""
global g_AiCompanionModelSelectorPendingRemove := false
global g_AiCompanionModelSelectorLastForegroundMonitorIdx := 0
global g_AiCompanionModelEscPollPrev := false

ShowAiCompanionModelSelector(companion) {
    global g_AiCompanionModelSelectorActive, g_AiCompanionModelSelectorCompanion,
        g_AiCompanionModelSelectorPendingRemove

    if !AiCompanionModels_IsValidCompanion(companion)
        return
    if (g_AiCompanionModelSelectorActive)
        AiCompanionModelSelector_Close()

    g_AiCompanionModelSelectorCompanion := companion
    g_AiCompanionModelSelectorPendingRemove := false
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
    loop 10 {
        try Hotkey(String(A_Index - 1), "Off")
    }
    loop 26 {
        try Hotkey(Chr(96 + A_Index), "Off")
    }
    try Hotkey("Escape", "Off")
    catch {
    }
}

AiCompanionModelSelector_BindRobustEscape() {
    global g_AiCompanionModelSelectorGui, g_OnEscapePressed, g_AiCompanionModelEscPollPrev
    SetTimer(AiCompanionModelSelector_EscapePoll, 0)
    if (!AiCompanionModelSelector_GuiHasWindow(g_AiCompanionModelSelectorGui))
        return
    Hotkey("$*Escape", AiCompanionModelSelector_EscapeFromHotkey, "On")
    g_OnEscapePressed := AiCompanionModelSelector_GlobalEscapeCallback
    Utils_EnsureGlobalEscapeHotkey()
    g_AiCompanionModelEscPollPrev := false
    SetTimer(AiCompanionModelSelector_EscapePoll, 50)
}

AiCompanionModelSelector_UnbindRobustEscape() {
    global g_OnEscapePressed, g_AiCompanionModelEscPollPrev
    SetTimer(AiCompanionModelSelector_EscapePoll, 0)
    g_AiCompanionModelEscPollPrev := false
    try Hotkey("$*Escape", AiCompanionModelSelector_EscapeFromHotkey, "Off")
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
            GetActiveMonitorWorkArea_StandardBar(&ml, &mt, &mr, &mb)
            ; Reposition via same path as PositionGui using current monitor index.
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
    global g_AiCompanionModelSelectorGui, g_AiCompanionModelSelectorActive,
        g_AiCompanionModelSelectorCompanion, g_AiCompanionModelSelectorPendingRemove,
        g_AiCompanionModelSelectorLastForegroundMonitorIdx

    if (!g_AiCompanionModelSelectorActive)
        return

    AiCompanionModelSelector_StopMonitorTracking()
    AiCompanionModelSelector_UnbindKeys()
    AiCompanionModelSelector_SafeDestroyGui(g_AiCompanionModelSelectorGui)
    g_AiCompanionModelSelectorGui := false

    companion := g_AiCompanionModelSelectorCompanion
    cfg := AiCompanionModels_Load(companion)
    title := AiCompanionModels_DisplayName(companion)
    fastLabel := (cfg.fast != "") ? cfg.fast : "(not set)"
    deepLabel := (cfg.deep != "") ? cfg.deep : "(not set)"

    g_AiCompanionModelSelectorGui := Gui("+AlwaysOnTop -Caption +ToolWindow +Owner -DPIScale")
    g_AiCompanionModelSelectorGui.BackColor := "1E1E2E"
    g_AiCompanionModelSelectorGui.MarginX := 20
    g_AiCompanionModelSelectorGui.MarginY := 15

    titleText := g_AiCompanionModelSelectorPendingRemove
        ? "🤖 " . title . " models — pick to REMOVE"
            : "🤖 " . title . " models"
    g_AiCompanionModelSelectorGui.SetFont("s14 cCDD6F4 Bold", "Segoe UI")
    g_AiCompanionModelSelectorGui.Add("Text", "w400 Center", titleText)
    g_AiCompanionModelSelectorGui.Add("Text", "w400 h1 Background45475A")

    g_AiCompanionModelSelectorGui.SetFont("s12 cA6E3A1", "Segoe UI")
    g_AiCompanionModelSelectorGui.Add("Text", "w400", "[f] Fast: " . fastLabel)
    g_AiCompanionModelSelectorGui.SetFont("s12 c89B4FA", "Segoe UI")
    g_AiCompanionModelSelectorGui.Add("Text", "w400", "[d] Deep: " . deepLabel)
    g_AiCompanionModelSelectorGui.Add("Text", "w400 h1 Background45475A y+8")

    g_AiCompanionModelSelectorGui.SetFont("s12 cCDD6F4", "Segoe UI")
    maxSlots := AiCompanionModels_MaxSlots()
    if (cfg.models.Length = 0) {
        g_AiCompanionModelSelectorGui.SetFont("s11 c6C7086", "Segoe UI")
        g_AiCompanionModelSelectorGui.Add("Text", "w400 Center", "(no extra models — press a to add)")
    } else {
        for i, name in cfg.models {
            if (i > maxSlots)
                break
            label := AiCompanionModels_LabelForIndex(i)
            g_AiCompanionModelSelectorGui.Add("Text", "w400", "[" . label . "] " . name)
        }
    }

    g_AiCompanionModelSelectorGui.Add("Text", "w400 h1 Background45475A y+10")
    g_AiCompanionModelSelectorGui.SetFont("s9 c6C7086", "Segoe UI")
    footerHint := g_AiCompanionModelSelectorPendingRemove
        ? "Press item key to remove | Esc cancel remove"
            : "1-9 / letters select | a add | r remove | f Fast | d Deep | Esc cancel"
    g_AiCompanionModelSelectorGui.Add("Text", "w400 Center", footerHint)

    try g_AiCompanionModelSelectorGui.OnEvent("Escape", AiCompanionModelSelector_GuiEscape)
    catch {
    }
    AiCompanionModelSelector_PositionGui(g_AiCompanionModelSelectorGui)
    g_AiCompanionModelSelectorLastForegroundMonitorIdx := GetMonitorIndexForForeground_StandardBar()

    count := Min(cfg.models.Length, maxSlots)
    loop count {
        label := AiCompanionModels_LabelForIndex(A_Index)
        Hotkey(label, AiCompanionModelSelector_HandleKey, "On")
    }
    Hotkey("a", AiCompanionModelSelector_AddEntry, "On")
    Hotkey("r", AiCompanionModelSelector_ArmRemove, "On")
    Hotkey("f", AiCompanionModelSelector_SelectFast, "On")
    Hotkey("d", AiCompanionModelSelector_SelectDeep, "On")
    AiCompanionModelSelector_BindRobustEscape()
    if (g_AiCompanionModelSelectorActive)
        SetTimer(AiCompanionModelSelector_TrackActiveMonitorTick, 115)
}

AiCompanionModelSelector_ArmRemove(*) {
    global g_AiCompanionModelSelectorActive, g_AiCompanionModelSelectorPendingRemove
    if (!g_AiCompanionModelSelectorActive)
        return
    g_AiCompanionModelSelectorPendingRemove := true
    AiCompanionModelSelector_Rebuild()
}

AiCompanionModelSelector_AddEntry(*) {
    global g_AiCompanionModelSelectorActive, g_AiCompanionModelSelectorCompanion,
        g_AiCompanionModelSelectorPendingRemove

    if (!g_AiCompanionModelSelectorActive)
        return
    g_AiCompanionModelSelectorPendingRemove := false
    AiCompanionModelSelector_UnbindKeys()
    AiCompanionModelSelector_UnbindRobustEscape()

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

    tagBox := InputBox("Role tag (optional): blank = list only, f = Fast, d = Deep",
        "Add AI model — role", "w440 h120")
    roleTag := ""
    if (tagBox.Result = "OK")
        roleTag := StrLower(Trim(tagBox.Value))

    cfg := AiCompanionModels_Load(companion)
    if (cfg.models.Length >= AiCompanionModels_MaxSlots()) {
        try ShowCenteredOverlay_Utils("❌ Model list is full.", 2500, BANNER_ACCENT_ERROR)
        AiCompanionModelSelector_Rebuild()
        return
    }

    if AiCompanionModels_AddModel(companion, name, roleTag)
        try ShowCenteredOverlay_Utils("✅ Added: " . name, 2000, BANNER_ACCENT_SUCCESS)
    AiCompanionModelSelector_Rebuild()
}

AiCompanionModelSelector_SelectFast(*) {
    global g_AiCompanionModelSelectorActive, g_AiCompanionModelSelectorCompanion,
        g_AiCompanionModelSelectorPendingRemove
    if (!g_AiCompanionModelSelectorActive)
        return
    if (g_AiCompanionModelSelectorPendingRemove) {
        g_AiCompanionModelSelectorPendingRemove := false
        AiCompanionModelSelector_Rebuild()
        return
    }
    companion := g_AiCompanionModelSelectorCompanion
    AiCompanionModelSelector_Close()
    AiCompanionModels_SelectRole(companion, "fast")
}

AiCompanionModelSelector_SelectDeep(*) {
    global g_AiCompanionModelSelectorActive, g_AiCompanionModelSelectorCompanion,
        g_AiCompanionModelSelectorPendingRemove
    if (!g_AiCompanionModelSelectorActive)
        return
    if (g_AiCompanionModelSelectorPendingRemove) {
        g_AiCompanionModelSelectorPendingRemove := false
        AiCompanionModelSelector_Rebuild()
        return
    }
    companion := g_AiCompanionModelSelectorCompanion
    AiCompanionModelSelector_Close()
    AiCompanionModels_SelectRole(companion, "deep")
}

AiCompanionModelSelector_HandleKey(thisHotkey := "") {
    global g_AiCompanionModelSelectorActive, g_AiCompanionModelSelectorCompanion,
        g_AiCompanionModelSelectorPendingRemove

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

    if (g_AiCompanionModelSelectorPendingRemove) {
        removedName := models[idx]
        g_AiCompanionModelSelectorPendingRemove := false
        if AiCompanionModels_RemoveModel(companion, idx)
            try ShowCenteredOverlay_Utils("✅ Removed: " . removedName, 2000, BANNER_ACCENT_SUCCESS)
        AiCompanionModelSelector_Rebuild()
        return
    }

    modelName := models[idx]
    AiCompanionModelSelector_Close()
    ok := AiCompanionModels_Apply(companion, modelName)
    if (!ok)
        try ShowCenteredOverlay_Utils("Could not select " . modelName, 2200, BANNER_ACCENT_ERROR)
}

AiCompanionModelSelector_Cancel(*) {
    global g_AiCompanionModelSelectorActive, g_AiCompanionModelSelectorPendingRemove
    if (g_AiCompanionModelSelectorActive && g_AiCompanionModelSelectorPendingRemove) {
        g_AiCompanionModelSelectorPendingRemove := false
        AiCompanionModelSelector_Rebuild()
        return
    }
    AiCompanionModelSelector_Close()
}

AiCompanionModelSelector_ForceReset() {
    global g_AiCompanionModelSelectorGui, g_AiCompanionModelSelectorActive,
        g_AiCompanionModelSelectorCompanion, g_AiCompanionModelSelectorPendingRemove,
        g_AiCompanionModelSelectorLastForegroundMonitorIdx

    g_AiCompanionModelSelectorActive := false
    g_AiCompanionModelSelectorCompanion := ""
    g_AiCompanionModelSelectorPendingRemove := false
    g_AiCompanionModelSelectorLastForegroundMonitorIdx := 0

    try AiCompanionModelSelector_StopMonitorTracking()
    try AiCompanionModelSelector_UnbindRobustEscape()
    try AiCompanionModelSelector_UnbindKeys()
    try Utils_EnsureGlobalEscapeHotkey()
    catch {
    }
    try AiCompanionModelSelector_SafeDestroyGui(g_AiCompanionModelSelectorGui)
    g_AiCompanionModelSelectorGui := false
}

AiCompanionModelSelector_Close() {
    global g_AiCompanionModelSelectorActive
    if (!g_AiCompanionModelSelectorActive)
        return
    AiCompanionModelSelector_ForceReset()
}
