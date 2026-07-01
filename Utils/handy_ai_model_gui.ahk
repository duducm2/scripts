; =============================================================================
; Utils module: handy_ai_model_gui.ahk
; ShowAiModelSelector GUI
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; ShowAiModelSelector() - Display selection GUI with immediate key capture
; =============================================================================
ShowAiModelSelector() {
    global g_AiModelSelectorGui, g_AiModelSelectorActive, g_HandyAiModels

    ; Don't show if already active
    if (g_AiModelSelectorActive)
        return

    currentSlot := Handy_GetPersistedAiModelSlot()

    ; Create selection GUI
    g_AiModelSelectorGui := Gui("+AlwaysOnTop -Caption +ToolWindow +Owner")
    g_AiModelSelectorGui.BackColor := "1E1E2E"
    g_AiModelSelectorGui.MarginX := 20
    g_AiModelSelectorGui.MarginY := 15

    ; Title
    g_AiModelSelectorGui.SetFont("s14 cCDD6F4 Bold", "Segoe UI")
    g_AiModelSelectorGui.Add("Text", "w280 Center", "🎙️ Select AI Model")
    g_AiModelSelectorGui.Add("Text", "w280 h1 Background45475A")  ; separator

    ; Model options (green row = last saved slot from data\handy_ai_model.ini)
    g_AiModelSelectorGui.SetFont("s12 cCDD6F4", "Segoe UI")
    for num, model in g_HandyAiModels {
        if (num = currentSlot) {
            g_AiModelSelectorGui.SetFont("s12 cA6E3A1 Bold", "Segoe UI")
            g_AiModelSelectorGui.Add("Text", "w280 Background313244", "[" . num . "] " . model.name)
            g_AiModelSelectorGui.SetFont("s9 cA6E3A1", "Segoe UI")
            g_AiModelSelectorGui.Add("Text", "w280 y+2 Background313244", "    " . model.desc)
        } else {
            g_AiModelSelectorGui.SetFont("s12 cCDD6F4", "Segoe UI")
            g_AiModelSelectorGui.Add("Text", "w280", "[" . num . "] " . model.name)
            g_AiModelSelectorGui.SetFont("s9 c6C7086", "Segoe UI")
            g_AiModelSelectorGui.Add("Text", "w280 y+2", "    " . model.desc)
        }
        g_AiModelSelectorGui.SetFont("s12 cCDD6F4", "Segoe UI")
    }

    ; Footer
    g_AiModelSelectorGui.Add("Text", "w280 h1 Background45475A y+10")
    g_AiModelSelectorGui.SetFont("s9 c6C7086", "Segoe UI")
    g_AiModelSelectorGui.Add("Text", "w280 Center", "Green row = last saved model")
    g_AiModelSelectorGui.Add("Text", "w280 Center y+4", "Press 1-3 | Esc to cancel")

    ; Get active window to determine which monitor to center on
    activeWin := 0
    try {
        activeWin := WinGetID("A")
    } catch {
        activeWin := 0
    }

    ; Default to primary monitor work area
    MonitorGetWorkArea(1, &monitorLeft, &monitorTop, &monitorRight, &monitorBottom)
    monitorWidth := monitorRight - monitorLeft
    monitorHeight := monitorBottom - monitorTop

    ; If we have an active window, find which monitor contains its center
    if (activeWin && activeWin != 0) {
        rect := Buffer(16, 0)
        if (DllCall("GetWindowRect", "ptr", activeWin, "ptr", rect)) {
            ; Calculate window center
            winLeft := NumGet(rect, 0, "int")
            winTop := NumGet(rect, 4, "int")
            winRight := NumGet(rect, 8, "int")
            winBottom := NumGet(rect, 12, "int")

            centerX := winLeft + (winRight - winLeft) // 2
            centerY := winTop + (winBottom - winTop) // 2

            ; Find which monitor contains the window center
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

    ; Measure GUI size and center on the active monitor
    g_AiModelSelectorGui.Show("AutoSize Hide")
    g_AiModelSelectorGui.GetPos(&gx, &gy, &gw, &gh)
    cx := monitorLeft + (monitorWidth - gw) // 2
    cy := monitorTop + (monitorHeight - gh) // 2
    g_AiModelSelectorGui.Show("x" . cx . " y" . cy . " NA")

    g_AiModelSelectorActive := true

    ; Enable hotkeys for 1-3 and Escape
    Hotkey("1", AiModelSelector_HandleKey, "On")
    Hotkey("2", AiModelSelector_HandleKey, "On")
    Hotkey("3", AiModelSelector_HandleKey, "On")
    Hotkey("Escape", AiModelSelector_Cancel, "On")
}

; Handle key press in AI model selector
AiModelSelector_HandleKey(key) {
    global g_AiModelSelectorGui, g_AiModelSelectorActive, g_HandyAiModels

    if (!g_AiModelSelectorActive)
        return

    ; Get the selection number
    selection := Integer(key)

    ; Close selector GUI
    AiModelSelector_Close()

    ; Execute the selection
    if (g_HandyAiModels.Has(selection)) {
        ExecuteHandyAiModelSelection(selection)
    }
}

; Cancel AI model selector
AiModelSelector_Cancel(*) {
    AiModelSelector_Close()
}

; Close the selector GUI and disable hotkeys
AiModelSelector_Close() {
    global g_AiModelSelectorGui, g_AiModelSelectorActive

    if (!g_AiModelSelectorActive)
        return

    g_AiModelSelectorActive := false

    ; Disable hotkeys
    try Hotkey("1", "Off")
    try Hotkey("2", "Off")
    try Hotkey("3", "Off")
    try Hotkey("Escape", AiModelSelector_Cancel, "Off")
    Utils_EnsureGlobalEscapeHotkey()

    ; Destroy GUI
    if (IsObject(g_AiModelSelectorGui) && g_AiModelSelectorGui.Hwnd) {
        try g_AiModelSelectorGui.Destroy()
    }
    g_AiModelSelectorGui := false
}
