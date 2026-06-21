; =============================================================================
; Shift keys module: m365_copilot_temp.ahk
; TEMPORARY M365 Copilot auto-continue (^!#n)
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

; =============================================================================
; TEMPORARY — M365 Copilot auto-continue loop (^!#n) — delete this whole section to remove
; Standalone Microsoft 365 Copilot (WebViewHost.exe), not Outlook-embedded Copilot.
; =============================================================================
global g_M365CopilotContinueActive := false
global g_M365CopilotContinueHwnd := 0
global g_M365CopilotContinueRestoreHwnd := 0
global g_M365CopilotContinuePhase := "idle"
global g_M365CopilotContinueSawStop := false
global g_M365CopilotContinueDeadline := 0
global g_M365CopilotContinueChromiumReady := false
global g_M365CopilotContinueSubmitTick := 0
global g_M365CopilotContinueBannerGui := 0
global g_M365CopilotContinueBannerBorderGui := 0

M365COPILOT_CONTINUE_APPEAR_MS := 90000
M365COPILOT_CONTINUE_IDLE_MS := 4000
M365COPILOT_CONTINUE_GENERATE_MS := 600000
M365COPILOT_CONTINUE_POLL_MS := 500

M365Copilot_IsTargetWindow(hwnd) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    try {
        if (StrLower(WinGetProcessName("ahk_id " hwnd)) != "webviewhost.exe")
            return false
        return InStr(WinGetTitle("ahk_id " hwnd), "Microsoft 365 Copilot") > 0
    } catch {
        return false
    }
}

M365Copilot_FindTargetHwnd() {
    hwnd := WinExist("A")
    if (M365Copilot_IsTargetWindow(hwnd))
        return hwnd
    try {
        for h in WinGetList("ahk_exe WebViewHost.exe") {
            if (M365Copilot_IsTargetWindow(h))
                return h
        }
    } catch {
    }
    return WinExist("Microsoft 365 Copilot ahk_exe WebViewHost.exe")
}

M365Copilot_RootFromHwnd(hwnd, activateMs := 500) {
    global g_M365CopilotContinueChromiumReady
    if (!hwnd)
        return ""
    ms := activateMs
    if (g_M365CopilotContinueChromiumReady && activateMs > 0)
        ms := 0
    try {
        root := UIA.ElementFromChromium("ahk_id " hwnd, ms)
        if (root) {
            g_M365CopilotContinueChromiumReady := true
            return root
        }
    } catch {
    }
    try {
        return UIA.ElementFromHandle(hwnd)
    } catch {
    }
    return ""
}

M365Copilot_FindFirstInRoot(root, criteriaList) {
    if (!root)
        return ""
    for criteria in criteriaList {
        try {
            el := root.FindFirst(criteria)
            if (el)
                return el
        } catch {
        }
    }
    return ""
}

M365Copilot_FindStopGenerating(root) {
    if (!root)
        return ""
    el := M365Copilot_FindFirstInRoot(root, [{ Name: "Stop generating", ControlType: "Button" }, { Name: "Stop generating",
        Type: 50000 }])
    if (el)
        return el
    try {
        return root.FindFirst({ Name: "Stop generating", matchmode: "Substring", ControlType: "Button" })
    } catch {
    }
    return ""
}

M365Copilot_FindComposer(root) {
    return M365Copilot_FindFirstInRoot(root, [{ AutomationId: "m365-chat-editor-target-element", ControlType: "Edit" }, { AutomationId: "m365-chat-editor-target-element" }, { Name: "Message Copilot",
        ControlType: "Edit" }])
}

M365Copilot_FocusComposerOnHwnd(hwnd, activateMs := 500) {
    root := M365Copilot_RootFromHwnd(hwnd, activateMs)
    el := M365Copilot_FindComposer(root)
    if (!el)
        return false
    try el.ScrollIntoView()
    try el.SetFocus()
    Sleep 40
    try el.Click()
    catch {
        return true
    }
    return true
}

M365Copilot_TrySubmitChat(root) {
    if (!root)
        return false
    stopBtn := M365Copilot_FindStopGenerating(root)
    if (stopBtn)
        return false
    sendBtn := M365Copilot_FindFirstInRoot(root, [{ Name: "Send ", matchmode: "Substring", ControlType: "Button" }, { ClassName: "fai-SendButton",
        matchmode: "Substring", ControlType: "Button" }])
    if (sendBtn) {
        try {
            if (sendBtn.GetPropertyValue(UIA.Property.IsInvokePatternAvailable)) {
                sendBtn.InvokePattern.Invoke()
                return true
            }
        } catch {
        }
        try {
            sendBtn.Click()
            return true
        } catch {
        }
    }
    SendInput "{Enter}"
    return true
}

M365Copilot_SendContinue(hwnd, restoreHwnd) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    try {
        WinActivate("ahk_id " hwnd)
        if (!WinWaitActive("ahk_id " hwnd, , 2))
            return false
        Sleep 120
        if (!M365Copilot_FocusComposerOnHwnd(hwnd, 500))
            return false
        Sleep 80
        SendInput "^a"
        Sleep 40
        SendText "continue"
        Sleep 60
        root := M365Copilot_RootFromHwnd(hwnd, 0)
        M365Copilot_TrySubmitChat(root)
        Sleep 80
        if (restoreHwnd && WinExist("ahk_id " restoreHwnd)) {
            try WinActivate("ahk_id " restoreHwnd)
        }
        return true
    } catch {
        if (restoreHwnd && WinExist("ahk_id " restoreHwnd)) {
            try WinActivate("ahk_id " restoreHwnd)
        }
        return false
    }
}

M365CopilotContinue_ShowBanner() {
    global g_M365CopilotContinueBannerGui, g_M365CopilotContinueBannerBorderGui
    M365CopilotContinue_HideBanner()
    text := "M365 Copilot: continue loop (Ctrl+Alt+Win+N to stop)"
    target := WinExist("A")
    hasWindow := false
    if (target && WinExist("ahk_id " target)) {
        try {
            WinGetPos(&wx, &wy, &ww, &wh, target)
            hasWindow := (ww > 0 && wh > 0)
        } catch {
            hasWindow := false
        }
    }
    ov := Gui("+AlwaysOnTop -Caption +ToolWindow")
    ov.BackColor := "1E1E2E"
    ov.SetFont("s18 cFFFFFF Bold", "Segoe UI")
    ov.Add("Text", "w560 Center", text)
    ov.Show("AutoSize Hide")
    ov.GetPos(&gx, &gy, &gw, &gh)
    if (hasWindow) {
        cx := wx + (ww - gw) // 2
        cy := wy + (wh - gh) // 2
    } else {
        vx := SysGet(76)
        vy := SysGet(77)
        vw := SysGet(78)
        vh := SysGet(79)
        cx := vx + (vw - gw) // 2
        cy := vy + (vh - gh) // 2
    }
    borderWidth := 6
    borderGui := Gui("+AlwaysOnTop -Caption +ToolWindow")
    borderGui.BackColor := BANNER_ACCENT_INTERMEDIATE
    borderGui.Show("NA x" . (cx - borderWidth) . " y" . (cy - borderWidth) . " w" . (gw + 2 * borderWidth) .
    " h" . (gh +
        2 * borderWidth))
    g_M365CopilotContinueBannerBorderGui := borderGui
    ov.Show("x" . cx . " y" . cy . " NA")
    WinSetTransparent(178, ov)
    g_M365CopilotContinueBannerGui := ov
}

M365CopilotContinue_HideBanner() {
    global g_M365CopilotContinueBannerGui, g_M365CopilotContinueBannerBorderGui
    try {
        if (IsObject(g_M365CopilotContinueBannerBorderGui))
            g_M365CopilotContinueBannerBorderGui.Destroy()
    } catch {
    }
    g_M365CopilotContinueBannerBorderGui := 0
    try {
        if (IsObject(g_M365CopilotContinueBannerGui))
            g_M365CopilotContinueBannerGui.Destroy()
    } catch {
    }
    g_M365CopilotContinueBannerGui := 0
}

M365CopilotContinue_Stop(reason := "") {
    global g_M365CopilotContinueActive, g_M365CopilotContinueHwnd, g_M365CopilotContinueRestoreHwnd
    global g_M365CopilotContinuePhase, g_M365CopilotContinueSawStop, g_M365CopilotContinueDeadline
    global g_M365CopilotContinueChromiumReady, g_M365CopilotContinueSubmitTick
    SetTimer(M365CopilotContinue_Timer, 0)
    SetTimer(M365CopilotContinue_DoSubmit, 0)
    g_M365CopilotContinueActive := false
    g_M365CopilotContinueHwnd := 0
    g_M365CopilotContinueRestoreHwnd := 0
    g_M365CopilotContinuePhase := "idle"
    g_M365CopilotContinueSawStop := false
    g_M365CopilotContinueDeadline := 0
    g_M365CopilotContinueChromiumReady := false
    g_M365CopilotContinueSubmitTick := 0
    M365CopilotContinue_HideBanner()
    if (reason != "")
        ShowCenteredOverlay_Utils(reason, 2200, BANNER_ACCENT_ERROR)
}

M365CopilotContinue_DoSubmit(*) {
    global g_M365CopilotContinueActive, g_M365CopilotContinueHwnd, g_M365CopilotContinueRestoreHwnd
    global g_M365CopilotContinuePhase, g_M365CopilotContinueSawStop, g_M365CopilotContinueDeadline
    if (!g_M365CopilotContinueActive)
        return
    hwnd := g_M365CopilotContinueHwnd
    if (!hwnd || !WinExist("ahk_id " hwnd) || !M365Copilot_IsTargetWindow(hwnd)) {
        M365CopilotContinue_Stop("❌ M365 Copilot window closed — loop stopped")
        return
    }
    if (!M365Copilot_SendContinue(hwnd, g_M365CopilotContinueRestoreHwnd)) {
        M365CopilotContinue_Stop("❌ M365 Copilot: could not send continue")
        return
    }
    g_M365CopilotContinuePhase := "waitAppear"
    g_M365CopilotContinueSawStop := false
    g_M365CopilotContinueSubmitTick := A_TickCount
    g_M365CopilotContinueDeadline := A_TickCount + M365COPILOT_CONTINUE_APPEAR_MS
}

M365CopilotContinue_Timer() {
    global g_M365CopilotContinueActive, g_M365CopilotContinueHwnd, g_M365CopilotContinuePhase
    global g_M365CopilotContinueSawStop, g_M365CopilotContinueDeadline
    if (!g_M365CopilotContinueActive)
        return
    hwnd := g_M365CopilotContinueHwnd
    if (!hwnd || !WinExist("ahk_id " hwnd) || !M365Copilot_IsTargetWindow(hwnd)) {
        M365CopilotContinue_Stop("❌ M365 Copilot window closed — loop stopped")
        return
    }
    if (g_M365CopilotContinuePhase = "waitAppear") {
        root := M365Copilot_RootFromHwnd(hwnd, 0)
        if (M365Copilot_FindStopGenerating(root)) {
            g_M365CopilotContinueSawStop := true
            g_M365CopilotContinuePhase := "waitGone"
            g_M365CopilotContinueDeadline := A_TickCount + M365COPILOT_CONTINUE_GENERATE_MS
            return
        }
        global g_M365CopilotContinueSubmitTick
        if (!g_M365CopilotContinueSawStop && (A_TickCount - g_M365CopilotContinueSubmitTick) >=
        M365COPILOT_CONTINUE_IDLE_MS) {
            SetTimer(M365CopilotContinue_DoSubmit, -1)
            return
        }
        if (A_TickCount >= g_M365CopilotContinueDeadline) {
            SetTimer(M365CopilotContinue_DoSubmit, -1)
            return
        }
        return
    }
    if (g_M365CopilotContinuePhase = "waitGone") {
        root := M365Copilot_RootFromHwnd(hwnd, 0)
        stopBtn := M365Copilot_FindStopGenerating(root)
        if (g_M365CopilotContinueSawStop && !stopBtn) {
            SetTimer(M365CopilotContinue_DoSubmit, -1)
            return
        }
        if (A_TickCount >= g_M365CopilotContinueDeadline) {
            M365CopilotContinue_Stop("❌ M365 Copilot: generation timeout — loop stopped")
        }
    }
}

M365CopilotContinue_Toggle(*) {
    global g_M365CopilotContinueActive, g_M365CopilotContinueHwnd, g_M365CopilotContinueRestoreHwnd
    global g_M365CopilotContinuePhase, g_M365CopilotContinueChromiumReady
    if (g_M365CopilotContinueActive) {
        M365CopilotContinue_Stop()
        ShowCenteredOverlay_Utils("M365 Copilot continue loop stopped", 1600, BANNER_ACCENT_SUCCESS)
        return
    }
    hwnd := M365Copilot_FindTargetHwnd()
    if (!hwnd) {
        ShowCenteredOverlay_Utils("❌ Microsoft 365 Copilot window not found", 2200, BANNER_ACCENT_ERROR)
        return
    }
    restore := WinExist("A")
    if (restore = hwnd)
        restore := 0
    g_M365CopilotContinueHwnd := hwnd
    g_M365CopilotContinueRestoreHwnd := restore
    g_M365CopilotContinueChromiumReady := false
    g_M365CopilotContinueActive := true
    g_M365CopilotContinuePhase := "submit"
    M365CopilotContinue_ShowBanner()
    SetTimer(M365CopilotContinue_Timer, M365COPILOT_CONTINUE_POLL_MS)
    SetTimer(M365CopilotContinue_DoSubmit, -1)
    ShowCenteredOverlay_Utils("M365 Copilot continue loop started", 1400, BANNER_ACCENT_SUCCESS)
}

; TEMPORARY — toggle M365 Copilot auto-continue loop
^!#n:: M365CopilotContinue_Toggle()
