; =============================================================================
; Shift keys module: outlook_helpers_01.ahk
; Outlook helper functions (part 1)
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

;-------------------------------------------------------------------
; Outlook Shortcuts
;-------------------------------------------------------------------

IsOutlookMessageActive() {
    return (WinActive("ahk_exe OUTLOOK.EXE") || WinActive("ahk_exe olk.exe"))
    && RegExMatch(WinGetTitle("A"), "i) - Message \(")
}

IsOutlookAppointmentActive() {
    ; Classic Outlook inspector windows use titles like " - Appointment/Meeting/Event".
    ; New Outlook editors often use titles like "New event - Outlook" and run under OUTLOOK.EXE or olk.exe.
    if !(WinActive("ahk_exe OUTLOOK.EXE") || WinActive("ahk_exe olk.exe"))
        return false

    t := ""
    try t := WinGetTitle("A")
    if RegExMatch(t, "i)(Appointment|Meeting|Event)")
        return true

    ; New Outlook: detect by UIA presence of the title field.
    if IsNewOutlookActive() {
        try {
            root := UIA.ElementFromHandle(WinExist("A"))
            if root.FindFirst({ Name: "Add title", ControlType: "Edit" })
                return true
            if root.FindFirst({ Name: "Add title", Type: 50004 })
                return true
            if root.FindFirst({ AutomationId: "4100" })
                return true
        } catch {
        }
    }

    return false
}

IsOutlookReminderActive() {
    return (WinActive("ahk_exe OUTLOOK.EXE") || WinActive("ahk_exe olk.exe"))
    && RegExMatch(WinGetTitle("A"), "i)Reminders?")
}

IsOutlookComposeActive() {
    ; New Outlook compose runs inside the main window and doesn't match the classic " - Message (" title.
    if !IsNewOutlookActive()
        return false
    try {
        root := UIA.ElementFromHandle(WinExist("A"))
        ; Prefer compose-only anchors seen in outlook-mail.md compose capture.
        if root.FindFirst({ AutomationId: "popoutCompose" })
            return true
        if root.FindFirst({ AutomationId: "discardCompose" })
            return true
        if root.FindFirst({ AutomationId: "splitButton-ram0__primaryActionButton" }) ; Send
            return true
        ; Fallback: presence of the compose Subject edit (MSG_*_SUBJECT) + Message body edit
        if root.FindFirst({ Name: "Subject", ControlType: "Edit" }) && root.FindFirst({ Name: "Message body",
            ControlType: "Edit" })
            return true
    } catch {
    }
    return false
}

IsOutlookMainActive() {
    if !(WinActive("ahk_exe OUTLOOK.EXE") || WinActive("ahk_exe olk.exe"))
        return false
    t := ""
    cls := ""
    try t := WinGetTitle("A")
    try cls := WinGetClass("A")
    ; Exclude inspectors and reminders
    if RegExMatch(t, "i) - Message \(")
        return false
    if RegExMatch(t, "i)(Appointment|Meeting|Event)")
        return false
    if RegExMatch(t, "i)^New event")
        return false
    if RegExMatch(t, "i)Reminder")
        return false
    ; Prefer the New Outlook shell window.
    if (cls != "" && InStr(cls, "Outlook Host"))
        return true
    if (t != "" && InStr(t, " - Outlook"))
        return true
    return true
}

Outlook_ActivateMainWindow() {
    ; Bring the main Outlook shell (Mail/Calendar) to front.
    try {
        wins := WinGetList("ahk_class Outlook Host")
        for hwnd in wins {
            t := ""
            try t := WinGetTitle("ahk_id " hwnd)
            if RegExMatch(t, "i)^(Mail|Calendar) - .* - Outlook") {
                try WinActivate("ahk_id " hwnd)
                try WinWaitActive("ahk_id " hwnd, , 1)
                return hwnd
            }
        }
    } catch {
    }
    ; Fallback: best-effort activate any Outlook process window.
    try {
        if WinExist("ahk_exe olk.exe")
            return WinActivate("ahk_exe olk.exe")
    } catch {
    }
    try {
        if WinExist("ahk_exe OUTLOOK.EXE")
            return WinActivate("ahk_exe OUTLOOK.EXE")
    } catch {
    }
    return 0
}

Outlook_FocusMainSearch() {
    Outlook_ActivateMainWindow()
    return OutlookFocusFirst([{ AutomationId: "topSearchInput", ControlType: "ComboBox" }, { AutomationId: "topSearchInput" }])
}

Outlook_SwitchToMail() {
    Outlook_ActivateMainWindow()
    ; Left-rail toggle (see outlook-mail.md): Chromium tree via OutlookMail_* — not OutlookClickFirst (ElementFromHandle misses WebView2).
    return OutlookMail_ClickFirst([{ AutomationId: "ddea774c-382b-47d7-aab5-adc2139a802b", ControlType: "Button" }, { Name: "Mail",
        ControlType: "Button" }, { Name: "Mail", Type: 50000 }
    ])
}

Outlook_SwitchToCalendar() {
    Outlook_ActivateMainWindow()
    return OutlookMail_ClickFirst([{ AutomationId: "8cbeb86f-83e1-43b5-aaba-cd3514322f0b", ControlType: "Button" }, { Name: "Calendar",
        ControlType: "Button" }, { Name: "Calendar", Type: 50000 }
    ])
}

; Calendar toolbar: "Go to today April 6, 2026" (dynamic date; caledar.md). WebView2 — use Chromium root.
; Loading bar during UIA work: docs/standard_information_display.md (Show → Update → Hide).
OutlookCalendar_ClickGoToToday() {
    try {
        StandardLoadingBar_Show("⏳ Outlook: Go to today…", BANNER_ACCENT_INTERMEDIATE, { passive: false, centerOnHwnd: 0,
            textWidth: 480, fontSize: 17 })
    } catch {
    }
    ok := false
    try {
        Outlook_ActivateMainWindow()
        try StandardLoadingBar_Update("⏳ Outlook: Opening Calendar…")
        if !Outlook_SwitchToCalendar()
            return false
        Sleep 150
        try StandardLoadingBar_Update("⏳ Outlook: Finding Go to today…")
        el := OutlookMail_FindFirst([{ Name: "Go to today", matchmode: "Substring", ControlType: "Button" }, { Name: "Go to today",
            matchmode: "Substring", Type: 50000 }, { Name: "Today", ControlType: "Button" }, { Name: "Today", matchmode: "Substring",
                ControlType: "Button" }
        ])
        if !el
            return false
        try el.ScrollIntoView()
        Sleep 40
        try el.SetFocus()
        Sleep 50
        try el.Click()
        catch Error {
            try el.Invoke()
            catch Error {
                return false
            }
        }
        ok := true
    } finally {
        try StandardLoadingBar_Hide(0)
        catch {
        }
    }
    return ok
}

; Left-rail Copilot app (copilot.md). AutomationId may vary by tenant; name fallback "Copilot".
global OUTLOOK_COPILOT_RAIL_AID := "b5abf2ae-c16b-4310-8f8a-d3bcdb52f162"

; Win+Alt+Shift+L modal: same pattern as Utils ShowAiModelSelector (1–9 + Esc).
global g_OutlookCopilotSelectorGui := false
global g_OutlookCopilotSelectorActive := false
global g_OutlookCopilotEscPollPrev := false
global g_OutlookCopilotShortcuts := [{ name: "Toggle Copilot voice chat", desc: "Voice chat / new voice session" }, { name: "New chat",
    desc: "Start a new Copilot chat" }, { name: "Focus Copilot input", desc: "Message Copilot composer" }, { name: "Open navigation panel",
        desc: "Expand Copilot nav drawer" }, { name: "Work scope", desc: "Toggle Work (grounded) scope" }, { name: "Web scope",
            desc: "Toggle Web scope" }, { name: "Model / mode (Auto…)", desc: "Open mode / model menu" }, { name: "Temporary chat",
                desc: "Temporary chat session" }, { name: "Chats and more", desc: "History and more options" }
]

OutlookCopilot_FindRailButton(root) {
    if !root
        return ""
    searchRoot := root
    try {
        r := root.FindFirst({ Name: "left-rail-appbar", matchmode: "Substring", ControlType: "Group" })
        if !r
            r := root.FindFirst({ Name: "left-rail-appbar", matchmode: "Substring" })
        if r
            searchRoot := r
    } catch {
    }
    el := ""
    try el := searchRoot.FindFirst({ AutomationId: OUTLOOK_COPILOT_RAIL_AID, ControlType: "Button" })
    if !el
        try el := searchRoot.FindFirst({ AutomationId: OUTLOOK_COPILOT_RAIL_AID })
    if !el
        try el := searchRoot.FindFirst({ Name: "Copilot", ControlType: "Button" })
    if !el
        try el := root.FindFirst({ AutomationId: OUTLOOK_COPILOT_RAIL_AID, ControlType: "Button" })
    if !el
        try el := root.FindFirst({ Name: "Copilot", ControlType: "Button" })
    return el
}

; Bounded wait: main Outlook shell is foreground (efficiency-canon; Alt-Tab from other apps).
OutlookCopilot_WaitMainOutlookForeground(timeoutMs := 2200) {
    hwnd := 0
    try {
        for h in WinGetList("ahk_class Outlook Host") {
            t := ""
            try t := WinGetTitle("ahk_id " h)
            if RegExMatch(t, "i)^(Mail|Calendar) - .* - Outlook") {
                hwnd := h
                break
            }
        }
    } catch {
    }
    if !hwnd {
        Outlook_ActivateMainWindow()
        deadline := A_TickCount + timeoutMs
        while (A_TickCount < deadline) {
            try {
                p := WinGetProcessName("A")
                if (p = "OUTLOOK.EXE" || p = "olk.exe")
                    return true
            } catch {
            }
            Outlook_ActivateMainWindow()
            Sleep 55
        }
        return false
    }
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        try {
            if (WinGetID("A") = hwnd)
                return true
        } catch {
        }
        try WinActivate("ahk_id " hwnd)
        try WinWaitActive("ahk_id " hwnd, , 0.35)
        Sleep 45
    }
    return false
}

OutlookCopilot_VoiceChatCriteriaList() {
    return [{ Name: "Start a new voice chat", matchmode: "Substring", ControlType: "Button" }, { Name: "Start a new voice chat",
        matchmode: "Substring", Type: 50000 }, { Name: "voice chat", matchmode: "Substring", ControlType: "Button" }
    ]
}

OutlookCopilot_FindVoiceChatButtonInRoot(root) {
    if !root
        return ""
    for criteria in OutlookCopilot_VoiceChatCriteriaList() {
        try {
            el := root.FindFirst(criteria)
            if el
                return el
        } catch {
        }
    }
    return ""
}

; Quality gate: element exists in UIA tree, is on-screen, enabled, and has non-zero bounds (rendered + clickable).
OutlookCopilot_ElementIsRenderedAndClickable(el) {
    if !IsObject(el)
        return false
    try {
        if el.GetPropertyValue(UIA.Property.IsOffscreen)
            return false
    } catch {
        try {
            if el.IsOffscreen
                return false
        } catch {
        }
    }
    try {
        if !el.GetPropertyValue(UIA.Property.IsEnabled)
            return false
    } catch {
        try {
            if !el.IsEnabled
                return false
        } catch {
            return false
        }
    }
    try {
        rect := el.BoundingRectangle
        l := rect.l, t := rect.t, r := rect.r, b := rect.b
        if (Abs((r - l) * (b - t)) < 4)
            return false
    } catch {
    }
    return true
}

; Wait until voice chat button passes quality gates OR timeout (poll + optional ScrollIntoView when stuck offscreen).
OutlookCopilot_WaitVoiceChatButtonReady(timeoutMs := 12000, pollMs := 120) {
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        root := OutlookMail_RootElement()
        if root {
            el := OutlookCopilot_FindVoiceChatButtonInRoot(root)
            if el {
                if OutlookCopilot_ElementIsRenderedAndClickable(el)
                    return el
                try el.ScrollIntoView()
            }
        }
        Sleep pollMs
    }
    return ""
}

; Ensure left-rail Copilot is selected (WebView2 / Chromium root).
OutlookCopilot_EnsureCopilotRailOn() {
    Outlook_ActivateMainWindow()
    root := OutlookMail_RootElement()
    if !root
        return false
    copBtn := OutlookCopilot_FindRailButton(root)
    if !copBtn
        return false
    if Outlook_RailToggleIsPressed(copBtn)
        return true
    try copBtn.ScrollIntoView()
    Sleep 40
    try copBtn.SetFocus()
    Sleep 50
    try copBtn.Click()
    catch Error {
        try copBtn.Invoke()
        catch Error {
            return false
        }
    }
    return true
}

; Bounded wait until Copilot chat surface or voice control appears (copilot.md mainChat / buttons).
OutlookCopilot_WaitCopilotChatSurface(timeoutMs := 2500) {
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        root := OutlookMail_RootElement()
        if root {
            try {
                if root.FindFirst({ AutomationId: "mainChat" })
                    return true
            } catch {
            }
            try {
                if root.FindFirst({ AutomationId: "iframe:" OUTLOOK_COPILOT_RAIL_AID })
                    return true
            } catch {
            }
            try {
                if root.FindFirst({ Name: "Start a new voice chat", matchmode: "Substring" })
                    return true
            } catch {
            }
            try {
                if root.FindFirst({ Name: "Message Copilot", ControlType: "Edit" })
                    return true
            } catch {
            }
        }
        Sleep 80
    }
    return false
}

OutlookCopilot_ToggleVoiceChat() {
    try {
        StandardLoadingBar_Show("⏳ Outlook Copilot: voice chat…", BANNER_ACCENT_INTERMEDIATE, { passive: false,
            centerOnHwnd: 0, textWidth: 480, fontSize: 17 })
    } catch {
    }
    ok := false
    try {
        try StandardLoadingBar_Update("⏳ Outlook Copilot: activating Outlook…")
        if !OutlookCopilot_WaitMainOutlookForeground(2200) {
            ShowCenteredOverlay_Utils("❌ Outlook Copilot: Outlook window not in foreground", 2000, BANNER_ACCENT_ERROR)
            return false
        }
        try StandardLoadingBar_Update("⏳ Outlook Copilot: opening Copilot…")
        if !OutlookCopilot_EnsureCopilotRailOn() {
            ShowCenteredOverlay_Utils("❌ Outlook Copilot: rail button not found", 1600, BANNER_ACCENT_ERROR)
            return false
        }
        try StandardLoadingBar_Update("⏳ Outlook Copilot: loading surface…")
        OutlookCopilot_WaitCopilotChatSurface(4000)
        try StandardLoadingBar_Update("⏳ Outlook Copilot: waiting for voice button (quality gate)…")
        el := OutlookCopilot_WaitVoiceChatButtonReady(12000, 120)
        if !el {
            ShowCenteredOverlay_Utils("❌ Outlook Copilot: voice chat button not ready (timeout)", 2200,
                BANNER_ACCENT_ERROR)
            return false
        }
        if !OutlookCopilot_ElementIsRenderedAndClickable(el) {
            root := OutlookMail_RootElement()
            el := root ? OutlookCopilot_FindVoiceChatButtonInRoot(root) : ""
            if !el || !OutlookCopilot_ElementIsRenderedAndClickable(el) {
                ShowCenteredOverlay_Utils("❌ Outlook Copilot: voice chat button not actionable", 2000,
                    BANNER_ACCENT_ERROR)
                return false
            }
        }
        try el.ScrollIntoView()
        Sleep 40
        try el.SetFocus()
        Sleep 50
        try el.Click()
        catch Error {
            try el.Invoke()
            catch Error {
                ShowCenteredOverlay_Utils("❌ Outlook Copilot: voice chat click failed", 1600, BANNER_ACCENT_ERROR)
                return false
            }
        }
        ok := true
    } finally {
        try StandardLoadingBar_Hide(0)
        catch {
        }
    }
    return ok
}

OutlookCopilot_FocusComposer() {
    el := OutlookMail_FindFirst([{ AutomationId: "m365-chat-editor-target-element", ControlType: "Edit" }, { AutomationId: "m365-chat-editor-target-element" }, { Name: "Message Copilot",
        ControlType: "Edit" }])
    if !el
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

OutlookCopilot_RunSlot(slot) {
    Outlook_ActivateMainWindow()
    switch slot {
        case 1:
            return OutlookCopilot_ToggleVoiceChat()
        case 2:
            if OutlookMail_ClickFirst([{ AutomationId: "new-chat-button", ControlType: "Button" }, { Name: "New chat",
                matchmode: "Substring", ControlType: "Button" }])
                return true
            ShowCenteredOverlay_Utils("❌ Outlook Copilot: New chat not found", 1600, BANNER_ACCENT_ERROR)
            return false
        case 3:
            if OutlookCopilot_FocusComposer()
                return true
            ShowCenteredOverlay_Utils("❌ Outlook Copilot: composer not found", 1600, BANNER_ACCENT_ERROR)
            return false
        case 4:
            if OutlookMail_ClickFirst([{ AutomationId: "sidepaneExpandButton", ControlType: "Button" }, { Name: "Open navigation panel",
                matchmode: "Substring", ControlType: "Button" }])
                return true
            ShowCenteredOverlay_Utils("❌ Outlook Copilot: navigation panel control not found", 1600,
                BANNER_ACCENT_ERROR)
            return false
        case 5:
            if OutlookMail_ClickFirst([{ AutomationId: "toggle-work", ControlType: "Button" }, { Name: "Work",
                ControlType: "Button" }])
                return true
            ShowCenteredOverlay_Utils("❌ Outlook Copilot: Work scope not found", 1600, BANNER_ACCENT_ERROR)
            return false
        case 6:
            if OutlookMail_ClickFirst([{ AutomationId: "toggle-web", ControlType: "Button" }, { Name: "Web",
                ControlType: "Button" }])
                return true
            ShowCenteredOverlay_Utils("❌ Outlook Copilot: Web scope not found", 1600, BANNER_ACCENT_ERROR)
            return false
        case 7:
            if OutlookMail_ClickFirst([{ AutomationId: "gptModeSwitcher", ControlType: "Button" }, { Name: "Auto",
                matchmode: "Substring", ControlType: "Button" }])
                return true
            ShowCenteredOverlay_Utils("❌ Outlook Copilot: mode menu not found", 1600, BANNER_ACCENT_ERROR)
            return false
        case 8:
            if OutlookMail_ClickFirst([{ AutomationId: "menura", ControlType: "Button" }, { Name: "Temporary chat",
                matchmode: "Substring", ControlType: "Button" }])
                return true
            ShowCenteredOverlay_Utils("❌ Outlook Copilot: Temporary chat not found", 1600, BANNER_ACCENT_ERROR)
            return false
        case 9:
            if OutlookMail_ClickFirst([{ AutomationId: "moreButton", ControlType: "Button" }, { Name: "OpenCopilot chats and more",
                matchmode: "Substring", ControlType: "Button" }, { Name: "chats and more", matchmode: "Substring",
                    ControlType: "Button" }
            ])
                return true
            ShowCenteredOverlay_Utils("❌ Outlook Copilot: More menu not found", 1600, BANNER_ACCENT_ERROR)
            return false
        default:
            return false
    }
}

ShowOutlookCopilotSelector() {
    global g_OutlookCopilotSelectorGui, g_OutlookCopilotSelectorActive, g_OutlookCopilotShortcuts
    if g_OutlookCopilotSelectorActive
        return
    g_OutlookCopilotSelectorGui := Gui("+AlwaysOnTop -Caption +ToolWindow +Owner")
    g_OutlookCopilotSelectorGui.BackColor := "1E1E2E"
    g_OutlookCopilotSelectorGui.MarginX := 20
    g_OutlookCopilotSelectorGui.MarginY := 15
    g_OutlookCopilotSelectorGui.SetFont("s14 cCDD6F4 Bold", "Segoe UI")
    g_OutlookCopilotSelectorGui.Add("Text", "w320 Center", "🤖 Outlook Copilot")
    g_OutlookCopilotSelectorGui.Add("Text", "w320 h1 Background45475A")
    g_OutlookCopilotSelectorGui.SetFont("s12 cCDD6F4", "Segoe UI")
    for num, item in g_OutlookCopilotShortcuts {
        g_OutlookCopilotSelectorGui.Add("Text", "w320", "[" num "] " item.name)
        g_OutlookCopilotSelectorGui.SetFont("s9 c6C7086", "Segoe UI")
        g_OutlookCopilotSelectorGui.Add("Text", "w320 y+2", "    " item.desc)
        g_OutlookCopilotSelectorGui.SetFont("s12 cCDD6F4", "Segoe UI")
    }
    g_OutlookCopilotSelectorGui.Add("Text", "w320 h1 Background45475A y+10")
    g_OutlookCopilotSelectorGui.SetFont("s9 c6C7086", "Segoe UI")
    g_OutlookCopilotSelectorGui.Add("Text", "w320 Center", "Press 1–9 | Esc to cancel")
    try {
        g_OutlookCopilotSelectorGui.OnEvent("Escape", OutlookCopilotSelector_GuiEscape)
    } catch {
    }
    activeWin := 0
    try activeWin := WinGetID("A")
    catch {
        activeWin := 0
    }
    MonitorGetWorkArea(1, &monitorLeft, &monitorTop, &monitorRight, &monitorBottom)
    monitorWidth := monitorRight - monitorLeft
    monitorHeight := monitorBottom - monitorTop
    if activeWin {
        rect := Buffer(16, 0)
        if DllCall("GetWindowRect", "ptr", activeWin, "ptr", rect) {
            winLeft := NumGet(rect, 0, "int")
            winTop := NumGet(rect, 4, "int")
            winRight := NumGet(rect, 8, "int")
            winBottom := NumGet(rect, 12, "int")
            centerX := winLeft + (winRight - winLeft) // 2
            centerY := winTop + (winBottom - winTop) // 2
            loop MonitorGetCount() {
                MonitorGetWorkArea(A_Index, &l, &t, &r, &b)
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
    g_OutlookCopilotSelectorGui.Show("AutoSize Hide")
    g_OutlookCopilotSelectorGui.GetPos(&gx, &gy, &gw, &gh)
    cx := monitorLeft + (monitorWidth - gw) // 2
    cy := monitorTop + (monitorHeight - gh) // 2
    ; Avoid "NA": if focus stays in Outlook/Chromium, Esc is consumed there first. Activate so Esc reaches this Gui.
    g_OutlookCopilotSelectorGui.Show("x" cx " y" cy)
    try WinActivate(g_OutlookCopilotSelectorGui.Hwnd)
    g_OutlookCopilotSelectorActive := true
    loop 9 {
        Hotkey(String(A_Index), OutlookCopilotSelector_HandleKey, "On")
    }
    ; $*Escape: $ forces keyboard hook (ignores synthetic Esc from SendInput in same script). I10 uses g_OnEscapePressed.
    Hotkey("$*Escape", OutlookCopilotSelector_EscapeFromHotkey, "On")
    global g_OnEscapePressed
    g_OnEscapePressed := OutlookCopilotSelector_GlobalEscapeCallback
    Utils_EnsureGlobalEscapeHotkey()
    ; Fallback when hook hotkeys miss Esc (e.g. another AHK process / hook order).
    g_OutlookCopilotEscPollPrev := false
    SetTimer(OutlookCopilotSelector_EscapePoll, 50)
}

; Poll Esc — fallback when $*Escape / g_OnEscapePressed miss. VK_ESCAPE 0x1B via GetAsyncKeyState.
OutlookCopilotSelector_EscapePoll() {
    global g_OutlookCopilotSelectorActive, g_OutlookCopilotEscPollPrev
    if !g_OutlookCopilotSelectorActive {
        SetTimer(OutlookCopilotSelector_EscapePoll, 0)
        return
    }
    escSync := GetKeyState("Escape", "P")
    escAsync := (DllCall("user32\GetAsyncKeyState", "int", 0x1B) & 0x8000) != 0
    escDown := escSync || escAsync
    if escDown {
        if !g_OutlookCopilotEscPollPrev {
            g_OutlookCopilotEscPollPrev := true
            OutlookCopilotSelector_Cancel()
        }
    } else {
        g_OutlookCopilotEscPollPrev := false
    }
}

