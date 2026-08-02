; =============================================================================
; Shift keys module: hotif_editor_01.ahk
; Cursor/VS Code editor hotkeys (part 1)
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

#HotIf IsEditorActive() && WinGetClass("A") != "#32770"

Cursor_ContextMenuSelectByDownAndVerify(targetText, openKey := "{AppsKey}", maxSteps := 28, expectedFileName :=
    "") {
    ; Open context menu.
    Send openKey
    Sleep 240

    ; Try to detect and read currently highlighted item and then walk down.
    step := 0
    while (step <= maxSteps) {
        step += 1
        highlightedEl := Cursor_ContextMenuGetHighlightedElement()
        name := ""
        try name := highlightedEl ? highlightedEl.Name : ""

        if (Mod(step, 3) = 0)
            StandardLoadingBar_Update("⏳ Searching menu item... (" step "/" maxSteps ")")

        if (name = targetText) {
            StandardLoadingBar_Update("⏳ Activating 'Add File to Cursor Chat'...")
            result := Cursor_ContextMenuActivateHighlightedItem(highlightedEl, targetText, 220, 900,
                expectedFileName)
            return result
        }

        Send "{Down}"
        Sleep 55
    }

    return { ok: false, reason: "Menu item not found" }
}

Cursor_ContextMenuSelectByDownAndVerifyAny(targetTexts, openKey := "{AppsKey}", maxSteps := 28,
    expectedFileName := "") {
    ; Open context menu.
    Send openKey
    Sleep 240

    step := 0
    while (step <= maxSteps) {
        step += 1
        highlightedEl := Cursor_ContextMenuGetHighlightedElement()
        name := ""
        try name := highlightedEl ? highlightedEl.Name : ""

        if (Mod(step, 3) = 0)
            StandardLoadingBar_Update("⏳ Searching menu item... (" step "/" maxSteps ")")

        for targetText in targetTexts {
            if (name = targetText) {
                StandardLoadingBar_Update("⏳ Activating '" . targetText . "'...")
                return Cursor_ContextMenuActivateHighlightedItem(highlightedEl, targetText, 220, 900,
                    expectedFileName)
            }
        }

        Send "{Down}"
        Sleep 55
    }

    return { ok: false, reason: "Menu item not found" }
}

Cursor_ContextMenuGetHighlightedElement() {
    ; Strategy A: focused element is a MenuItem
    try {
        fe := UIA.GetFocusedElement()
        if (fe) {
            try {
                if (fe.ControlType = UIA.Type.MenuItem)
                    return fe
            } catch {
            }
        }
    } catch {
    }

    ; Strategy B: find selected MenuItem
    try {
        hwnd := WinExist("ahk_exe Cursor.exe")
        root := UIA.ElementFromHandle(hwnd)
        all := root.FindAll({ Type: 50011 })
        for mi in all {
            try {
                if (mi.GetPropertyValue(UIA.Property.IsSelected)) {
                    return mi
                }
            } catch {
            }
        }
    } catch {
    }

    ; Strategy C: find MenuItem with keyboard focus
    try {
        hwnd := WinExist("ahk_exe Cursor.exe")
        root := UIA.ElementFromHandle(hwnd)
        all := root.FindAll({ Type: 50011 })
        for mi in all {
            try {
                if (mi.GetPropertyValue(UIA.Property.HasKeyboardFocus)) {
                    return mi
                }
            } catch {
            }
        }
    } catch {
    }

    return 0
}

Cursor_ContextMenuActivateHighlightedItem(menuItemEl, targetText, settleMs := 220, verifyTimeoutMs := 900,
    expectedFileName := "") {
    ; Requirement: add a short pause before activation to improve reliability.
    Sleep settleMs

    ; Extra stabilization: ensure the highlighted item stays on the target briefly.
    stableOk := Cursor_WaitForContextMenuHighlightStable(targetText, 260)

    activatedVia := ""
    ok := false

    ; Prefer invoking the element directly (more reliable than raw Enter).
    if (menuItemEl) {
        ; Make sure it really has focus before activation (helps Enter/registering).
        try menuItemEl.SetFocus()
        Sleep 80

        try {
            invAvail := menuItemEl.GetPropertyValue(UIA.Property.IsInvokePatternAvailable)
            if (invAvail = 1) {
                activatedVia := "invokePattern"
                menuItemEl.InvokePattern.Invoke()
                ok := true
            }
        } catch {
        }

        if (!ok) {
            try {
                activatedVia := "click"
                menuItemEl.Click()
                ok := true
            } catch {
                ok := false
            }
        }

        if (!ok) {
            try {
                activatedVia := "focus+enter"
                menuItemEl.SetFocus()
                Sleep 40
                Send "{Enter}"
                ok := true
            } catch {
                ok := false
            }
        }
    } else {
        activatedVia := "enterOnly"
        Send "{Enter}"
        ok := true
    }

    ; Requirement: quality check - verify the menu action actually took effect.
    ; Best available signal: the menu disappears (target menu item no longer present).
    closed := Cursor_WaitForContextMenuItemGone(targetText, verifyTimeoutMs)

    ; Retry once if it didn't close (covers intermittent Enter not registering).
    if (!closed && menuItemEl) {
        try {
            activatedVia .= "+retryEnter"
            menuItemEl.SetFocus()
            Sleep 80
            Send "{Enter}"
        } catch {
        }
        closed := Cursor_WaitForContextMenuItemGone(targetText, verifyTimeoutMs)
    }

    if (!closed) {
        return { ok: false, reason: "Context menu did not close" }
    }

    verifyResult := Cursor_WaitForAddFileChipSuccess(expectedFileName, 1200)
    if (verifyResult.ok)
        return verifyResult

    ; Controlled retry when menu action likely did not trigger composer state.
    if (menuItemEl) {
        try {
            menuItemEl.SetFocus()
            Sleep 70
            Send "{Enter}"
        } catch {
        }
        verifyResult := Cursor_WaitForAddFileChipSuccess(expectedFileName, 1200)
        if (verifyResult.ok)
            return verifyResult
    }

    reason := verifyResult.reason
    if (reason = "")
        reason := Cursor_DetectAddFileFailureSignal(expectedFileName)
    if (reason = "")
        reason := "Chat context signal missing after action"
    return { ok: false, reason: reason }
}

Cursor_DetectAddFileFailureSignal(expectedFileName := "") {
    hwnd := WinExist("ahk_exe Cursor.exe")
    if (!hwnd)
        return "Target window closed"

    try root := UIA.ElementFromHandle(hwnd)
    catch
        return "UIA unreachable"

    if (!root)
        return "UIA unreachable"

    if (expectedFileName != "")
        return "Chat context signal missing for '" . expectedFileName . "'"

    return ""
}

Cursor_WaitForAddFileChipSuccess(expectedFileName := "", timeoutMs := 1800) {
    if (expectedFileName = "")
        return { ok: false, reason: "Selected file name unavailable for verification" }

    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        hwnd := WinExist("ahk_exe Cursor.exe")
        if (!hwnd)
            break

        try root := UIA.ElementFromHandle(hwnd)
        catch {
            Sleep 80
            continue
        }
        if (!root) {
            Sleep 80
            continue
        }

        if (Cursor_IsChatFileContextVisible(root, expectedFileName))
            return { ok: true, reason: "" }

        Sleep 80
    }

    reason := Cursor_DetectAddFileFailureSignal(expectedFileName)
    if (reason != "")
        SafeDebugLog("Cursor_WaitForAddFileChipSuccess failed: " . reason)
    else
        SafeDebugLog("Cursor_WaitForAddFileChipSuccess timed out waiting for chat context chip")
    return { ok: false, reason: reason }
}

Cursor_GetFocusedExplorerItemName() {
    try {
        fe := UIA.GetFocusedElement()
        if (fe) {
            name := ""
            try name := fe.Name
            if (name != "")
                return name
        }
    } catch {
    }

    try {
        hwnd := WinExist("ahk_exe Cursor.exe")
        if (!hwnd)
            return ""
        root := UIA.ElementFromHandle(hwnd)
        if (!root)
            return ""
        items := root.FindAll({ Type: UIA.Type.TreeItem })
        for item in items {
            try {
                if (!item.GetPropertyValue(UIA.Property.IsSelected))
                    continue
                if (item.GetPropertyValue(UIA.Property.IsOffscreen))
                    continue
                nm := item.Name
                if (nm != "")
                    return nm
            } catch {
            }
        }
    } catch {
    }
    return ""
}

Cursor_FindVisibleComposerInput(root) {
    try edits := root.FindAll({ Type: UIA.Type.Edit })
    catch
        return 0
    for editEl in edits {
        className := ""
        try className := editEl.ClassName
        if (!InStr(className, "aislash-editor-input"))
            continue
        try {
            if editEl.GetPropertyValue(UIA.Property.IsOffscreen)
                continue
        } catch {
            continue
        }
        return editEl
    }
    return 0
}

Cursor_IsNameNearComposer(root, controlType, needleName) {
    composer := Cursor_FindVisibleComposerInput(root)
    if (!composer)
        return false
    try compBr := composer.BoundingRectangle
    catch
        return false

    try all := root.FindAll({ Type: controlType })
    catch
        return false

    for el in all {
        nm := ""
        try nm := el.Name
        if (nm = "")
            continue
        if (!InStr(nm, needleName, false))
            continue
        try {
            if el.GetPropertyValue(UIA.Property.IsOffscreen)
                continue
        } catch {
            continue
        }
        try {
            br := el.BoundingRectangle
            ; Keep only elements close to the composer area to avoid Explorer false positives.
            if (Abs(br.t - compBr.t) > 260)
                continue
            if (Abs(br.l - compBr.l) > 520)
                continue
            return true
        } catch {
        }
    }
    return false
}

Cursor_IsChatFileContextVisible(root, expectedFileName) {
    needle := expectedFileName
    if (InStr(needle, "\") || InStr(needle, "/"))
        needle := RegExReplace(needle, "^.*[\\/]")
    needleNoExt := RegExReplace(needle, "\.[^.]+$")

    if (Cursor_IsNameNearComposer(root, UIA.Type.Text, needle))
        return true
    if (Cursor_IsNameNearComposer(root, UIA.Type.Button, needle))
        return true
    if (needleNoExt != "" && needleNoExt != needle) {
        if (Cursor_IsNameNearComposer(root, UIA.Type.Text, needleNoExt))
            return true
        if (Cursor_IsNameNearComposer(root, UIA.Type.Button, needleNoExt))
            return true
    }
    return false
}

Cursor_FallbackAddFileByAtMention(expectedFileName := "") {
    copiedName := Cursor_CopySelectedFileNameFromExplorerRename()
    if (copiedName = "") {
        SafeDebugLog("Cursor_FallbackAddFileByAtMention failed: copy-name phase")
        return { ok: false, reason: "Could not copy selected file name" }
    }

    if (!Cursor_FocusAITextFieldForAddFallback()) {
        SafeDebugLog("Cursor_FallbackAddFileByAtMention failed: focus-input phase")
        return { ok: false, reason: "Could not focus AI text field" }
    }

    SendInput "@" . copiedName
    mentionDropdown := Cursor_WaitForMentionDropdownSignal(copiedName, 1200)
    if (!mentionDropdown.ok) {
        SafeDebugLog("Cursor_FallbackAddFileByAtMention failed: dropdown-missing for " . copiedName)
        return mentionDropdown
    }

    Send "{Enter}"
    verifyResult := Cursor_WaitForAddFileChipSuccess(copiedName, 1400)
    if (verifyResult.ok)
        return verifyResult

    SafeDebugLog("Cursor_FallbackAddFileByAtMention failed: confirm-failed for " . copiedName)
    return { ok: false, reason: "Confirm failed after @filename" }
}

Cursor_CopySelectedFileNameFromExplorerRename() {
    oldClip := A_Clipboard
    copied := ""
    try {
        A_Clipboard := ""
        Send "{F2}"
        Sleep 120
        Send "^a"
        Sleep 40
        Send "^c"
        if (ClipWait(0.8))
            copied := A_Clipboard
        ; Exit rename mode safely without changing file name.
        Send "{Esc}"
        Sleep 70
    } catch {
        try Send "{Esc}"
    } finally {
        A_Clipboard := oldClip
    }

    copied := Trim(StrReplace(StrReplace(copied, "`r", ""), "`n", ""))
    return copied
}

Cursor_FocusAITextFieldForAddFallback() {
    hwnd := WinExist("ahk_exe Cursor.exe")
    if (!hwnd)
        return false
    try WinActivate("ahk_id " hwnd)
    catch
        return false
    Sleep 120

    try root := UIA.ElementFromHandle(hwnd)
    catch
        return false
    if (!root)
        return false

    editEl := Cursor_FindVisibleComposerInput(root)
    if (editEl) {
        try editEl.SetFocus()
        Sleep 60
        try {
            if editEl.HasKeyboardFocus
                return true
        } catch {
        }
        try editEl.Click()
        Sleep 60
        try return editEl.HasKeyboardFocus
        catch
            return false
    }

    ; Pane may be hidden. Open AI pane, then retry composer search.
    Send "^i"
    Sleep 250
    loop 8 {
        try root := UIA.ElementFromHandle(hwnd)
        catch {
            Sleep 120
            continue
        }
        if (!root) {
            Sleep 120
            continue
        }
        editEl := Cursor_FindVisibleComposerInput(root)
        if (editEl) {
            try editEl.SetFocus()
            Sleep 60
            try {
                if editEl.HasKeyboardFocus
                    return true
            } catch {
            }
            try editEl.Click()
            Sleep 60
            try return editEl.HasKeyboardFocus
            catch
                return false
        }
        Sleep 120
    }
    return false
}

Cursor_WaitForMentionDropdownSignal(expectedFileName, timeoutMs := 1200) {
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        hwnd := WinExist("ahk_exe Cursor.exe")
        if (!hwnd)
            break
        try root := UIA.ElementFromHandle(hwnd)
        catch {
            Sleep 80
            continue
        }
        if (!root) {
            Sleep 80
            continue
        }

        if (Cursor_HasMentionCandidateNearComposer(root, expectedFileName))
            return { ok: true, reason: "" }
        Sleep 80
    }
    return { ok: false, reason: "Symbol dropdown not detected for @" . expectedFileName }
}

Cursor_HasMentionCandidateNearComposer(root, expectedFileName) {
    composer := Cursor_FindVisibleComposerInput(root)
    if (!composer)
        return false
    try compBr := composer.BoundingRectangle
    catch
        return false

    needle := expectedFileName
    needleNoExt := RegExReplace(needle, "\.[^.]+$")

    try allText := root.FindAll({ Type: UIA.Type.Text })
    catch
        return false

    for t in allText {
        nm := ""
        try nm := t.Name
        if (nm = "")
            continue
        if (!InStr(nm, needle, false) && !(needleNoExt != "" && InStr(nm, needleNoExt, false)))
            continue
        try {
            if t.GetPropertyValue(UIA.Property.IsOffscreen)
                continue
        } catch {
            continue
        }
        try {
            br := t.BoundingRectangle
            ; Mention dropdown generally appears near/below composer.
            if (Abs(br.l - compBr.l) > 650)
                continue
            if (br.t < compBr.t - 120 || br.t > compBr.t + 520)
                continue
            return true
        } catch {
        }
    }
    return false
}

Cursor_WaitForContextMenuHighlightStable(targetText, timeoutMs := 250, requiredConsecutive := 3) {
    ; Consider highlight stable if we observe the target highlighted N consecutive times.
    deadline := A_TickCount + timeoutMs
    consec := 0
    while (A_TickCount < deadline) {
        el := Cursor_ContextMenuGetHighlightedElement()
        nm := ""
        try nm := el ? el.Name : ""
        if (nm = targetText) {
            consec += 1
            if (consec >= requiredConsecutive)
                return true
        } else {
            consec := 0
        }
        Sleep 40
    }
    return false
}

Cursor_WaitForContextMenuItemGone(itemName, timeoutMs := 800) {
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        try {
            mi := UIA.GetRootElement().FindFirst({ Type: 50011, Name: itemName })
            if (!mi)
                return true
        } catch {
            ; If searching fails (menu destroyed), treat as gone.
            return true
        }
        Sleep 40
    }
    return false
}

CursorShortcutMenu_EscapeFromHotkey(*) {
    CursorShortcutMenu_Cancel()
}

CursorShortcutMenu_GlobalEscapeCallback(*) {
    global g_CursorShortcutMenuActive
    if (!g_CursorShortcutMenuActive)
        return false
    CursorShortcutMenu_Cancel()
    return true
}

CursorShortcutMenu_GuiEscape(*) {
    CursorShortcutMenu_Cancel()
}

CursorShortcutMenu_BindRobustEscape() {
    global g_CursorShortcutMenuGui, g_OnEscapePressed, g_CursorShortcutMenuEscPollPrev
    SetTimer(CursorShortcutMenu_EscapePoll, 0)
    if (!IsObject(g_CursorShortcutMenuGui) || !g_CursorShortcutMenuGui.Hwnd)
        return
    try {
        g_CursorShortcutMenuGui.OnEvent("Escape", CursorShortcutMenu_GuiEscape)
    } catch {
    }
    Hotkey("$*Escape", CursorShortcutMenu_EscapeFromHotkey, "On")
    global g_OnEscapePressed
    g_OnEscapePressed := CursorShortcutMenu_GlobalEscapeCallback
    Utils_EnsureGlobalEscapeHotkey()
    g_CursorShortcutMenuEscPollPrev := false
    SetTimer(CursorShortcutMenu_EscapePoll, 50)
}

CursorShortcutMenu_UnbindRobustEscape() {
    global g_OnEscapePressed, g_CursorShortcutMenuEscPollPrev
    SetTimer(CursorShortcutMenu_EscapePoll, 0)
    g_CursorShortcutMenuEscPollPrev := false
    try Hotkey("Escape", CursorShortcutMenu_Cancel, "Off")
    catch {
    }
    try Hotkey("*Escape", CursorShortcutMenu_Cancel, "Off")
    catch {
    }
    try Hotkey("$*Escape", CursorShortcutMenu_EscapeFromHotkey, "Off")
    catch {
    }
    global g_OnEscapePressed
    if (g_OnEscapePressed = CursorShortcutMenu_GlobalEscapeCallback)
        g_OnEscapePressed := ""
    Utils_EnsureGlobalEscapeHotkey()
}

; Poll Esc — fallback when $*Escape / g_OnEscapePressed miss (StudyTopicSelector_EscapePoll).
CursorShortcutMenu_EscapePoll() {
    global g_CursorShortcutMenuActive, g_CursorShortcutMenuEscPollPrev
    if (!g_CursorShortcutMenuActive) {
        SetTimer(CursorShortcutMenu_EscapePoll, 0)
        return
    }
    escSync := GetKeyState("Escape", "P")
    escAsync := (DllCall("user32\GetAsyncKeyState", "int", 0x1B) & 0x8000) != 0
    escDown := escSync || escAsync
    if (escDown) {
        if (!g_CursorShortcutMenuEscPollPrev) {
            g_CursorShortcutMenuEscPollPrev := true
            CursorShortcutMenu_Cancel()
        }
    } else {
        g_CursorShortcutMenuEscPollPrev := false
    }
}

ShowCursorShortcutMenu() {
    global g_CursorShortcutMenuGui, g_CursorShortcutMenuActive
    if (g_CursorShortcutMenuActive)
        return

    g_CursorShortcutMenuGui := Gui("+AlwaysOnTop -Caption +ToolWindow +Owner")
    g_CursorShortcutMenuGui.BackColor := "1E1E2E"
    g_CursorShortcutMenuGui.MarginX := 20
    g_CursorShortcutMenuGui.MarginY := 15

    g_CursorShortcutMenuGui.SetFont("s14 cCDD6F4 Bold", "Segoe UI")
    g_CursorShortcutMenuGui.Add("Text", "w300 Center", "Select shortcut")
    g_CursorShortcutMenuGui.Add("Text", "w300 h1 Background45475A")

    g_CursorShortcutMenuGui.SetFont("s12 cCDD6F4", "Segoe UI")
    g_CursorShortcutMenuGui.Add("Text", "w300", "[1] hello world one")
    g_CursorShortcutMenuGui.Add("Text", "w300", "[2] hello world two")

    g_CursorShortcutMenuGui.Add("Text", "w300 h1 Background45475A y+8")
    g_CursorShortcutMenuGui.SetFont("s10 c6C7086", "Segoe UI")
    g_CursorShortcutMenuGui.Add("Text", "w300", "Terminal permissions")
    g_CursorShortcutMenuGui.SetFont("s12 cCDD6F4", "Segoe UI")
    g_CursorShortcutMenuGui.Add("Text", "w300", "[R] Run (terminal permission)")
    g_CursorShortcutMenuGui.Add("Text", "w300", "[A] Allowlist (any permission button)")
    g_CursorShortcutMenuGui.Add("Text", "w300", "[F] Mark as fixed")
    g_CursorShortcutMenuGui.Add("Text", "w300", "[P] Proceed")
    g_CursorShortcutMenuGui.Add("Text", "w300", "[E] Fetch")

    g_CursorShortcutMenuGui.Add("Text", "w300 h1 Background45475A y+10")
    g_CursorShortcutMenuGui.SetFont("s9 c6C7086", "Segoe UI")
    g_CursorShortcutMenuGui.Add("Text", "w300 Center", "Press 1–2 | R · A · F · P · E | Esc to cancel")

    ; Center on same monitor as active window (same logic as Utils ShowAiModelSelector)
    activeWin := 0
    try
        activeWin := WinGetID("A")
    catch
        activeWin := 0
    MonitorGetWorkArea(1, &monitorLeft, &monitorTop, &monitorRight, &monitorBottom)
    monitorWidth := monitorRight - monitorLeft
    monitorHeight := monitorBottom - monitorTop
    if (activeWin && activeWin != 0) {
        rect := Buffer(16, 0)
        if (DllCall("GetWindowRect", "ptr", activeWin, "ptr", rect)) {
            centerX := NumGet(rect, 0, "int") + (NumGet(rect, 8, "int") - NumGet(rect, 0, "int")) // 2
            centerY := NumGet(rect, 4, "int") + (NumGet(rect, 12, "int") - NumGet(rect, 4, "int")) // 2
            loop MonitorGetCount() {
                MonitorGetWorkArea(A_Index, &l, &t, &r, &b)
                if (centerX >= l && centerX <= r && centerY >= t && centerY <= b) {
                    monitorLeft := l
                    monitorTop := t
                    monitorWidth := r - l
                    monitorHeight := b - t
                    break
                }
            }
        }
    }
    g_CursorShortcutMenuGui.Show("AutoSize Hide")
    g_CursorShortcutMenuGui.GetPos(&gx, &gy, &gw, &gh)
    cx := monitorLeft + (monitorWidth - gw) // 2
    cy := monitorTop + (monitorHeight - gh) // 2
    global g_CursorShortcutMenuPrevHwnd
    g_CursorShortcutMenuPrevHwnd := activeWin
    ; Avoid "NA": if focus stays in Cursor/Chromium, Esc is consumed there first (ShowAiModelSelector).
    g_CursorShortcutMenuGui.Show("x" . cx . " y" . cy)
    try WinActivate(g_CursorShortcutMenuGui.Hwnd)

    g_CursorShortcutMenuActive := true
    try HotIf()
    catch {
    }
    Hotkey("1", (*) => CursorShortcutMenu_HandleKey("1"), "On")
    Hotkey("2", (*) => CursorShortcutMenu_HandleKey("2"), "On")
    Hotkey("r", (*) => CursorShortcutMenu_HandleKey("r"), "On")
    Hotkey("a", (*) => CursorShortcutMenu_HandleKey("a"), "On")
    Hotkey("f", (*) => CursorShortcutMenu_HandleKey("f"), "On")
    Hotkey("p", (*) => CursorShortcutMenu_HandleKey("p"), "On")
    Hotkey("e", (*) => CursorShortcutMenu_HandleKey("e"), "On")
    CursorShortcutMenu_BindRobustEscape()
}

CursorShortcutMenu_HandleKey(key) {
    global g_CursorShortcutMenuActive
    if (!g_CursorShortcutMenuActive)
        return
    CursorShortcutMenu_Close()
    if (key = "1")
        CursorShortcutMenu_Action1()
    else if (key = "2")
        CursorShortcutMenu_Action2()
    else if (key = "r")
        CursorShortcutMenu_ActionRun()
    else if (key = "a")
        CursorShortcutMenu_ActionAllowlist()
    else if (key = "f")
        CursorShortcutMenu_ActionMarkAsFixed()
    else if (key = "p")
        CursorShortcutMenu_ActionProceed()
    else if (key = "e")
        CursorShortcutMenu_ActionFetch()
}

CursorShortcutMenu_Cancel(*) {
    CursorShortcutMenu_Close()
}

CursorShortcutMenu_Close() {
    global g_CursorShortcutMenuGui, g_CursorShortcutMenuActive, g_CursorShortcutMenuPrevHwnd
    if (!g_CursorShortcutMenuActive)
        return
    g_CursorShortcutMenuActive := false
    try Hotkey("1", "Off")
    try Hotkey("2", "Off")
    try Hotkey("r", "Off")
    try Hotkey("a", "Off")
    try Hotkey("f", "Off")
    try Hotkey("p", "Off")
    try Hotkey("e", "Off")
    CursorShortcutMenu_UnbindRobustEscape()
    if (IsObject(g_CursorShortcutMenuGui) && g_CursorShortcutMenuGui.Hwnd) {
        try g_CursorShortcutMenuGui.Destroy()
    }
    g_CursorShortcutMenuGui := false
    prevHwnd := g_CursorShortcutMenuPrevHwnd
    g_CursorShortcutMenuPrevHwnd := 0
    if (prevHwnd && WinExist("ahk_id " prevHwnd)) {
        try WinActivate("ahk_id " prevHwnd)
        catch {
        }
    }
}

CursorShortcutMenu_Action1(*) {
    ; Replace with your command for "hello world one"
    return
}

CursorShortcutMenu_Action2(*) {
    ; Replace with your command for "hello world two"
    return
}

CursorShortcutMenu_ActionRun(*) {
    Cursor_ClickPermissionLabel("Run")
}

CursorShortcutMenu_ActionAllowlist(*) {
    Cursor_ClickPermissionLabelContains("Allowlist")
}

CursorShortcutMenu_ActionMarkAsFixed(*) {
    Cursor_ClickPermissionLabel("Mark as fixed")
}

CursorShortcutMenu_ActionProceed(*) {
    Cursor_ClickPermissionLabel("Proceed")
}

CursorShortcutMenu_ActionFetch(*) {
    ; UIA: Type 50020 (Text), Name "Fetch", LocalizedType "text"
    Cursor_ClickPermissionLabel("Fetch")
}

; Ctrl + H : Smart navigation - Editor → Explorer, Explorer → Reveal in Explorer
; Works from main editor even when the left Explorer sidebar is closed (opens it first).
^h:: Editor_SmartNavReveal("")

; Alt + I : Same as Ctrl+H; when Windows Explorer opens, copy file and close Explorer
!i:: Editor_SmartNavReveal("copy")

; Alt + H : Same as Ctrl+H; when Windows Explorer opens, open file and close Explorer
!h:: Editor_SmartNavReveal("open")

; Ctrl + 1 : Remove clustering and focus on the code
^1 up::
{
    ; Send ESC two times
    SendEscape()  ; ESC
    Sleep 50
    SendEscape()  ; ESC again
    Sleep 100
    Send "^!n"
    Sleep 100
    Send "^!,"
    Sleep 150
    ClickHidePanelButton()
}

; Ctrl + 2 : Ensure Files Explorer focus, then native Copy Path
$^2:: {
    editorHwnd := WinExist("A")
    if !Editor_FocusIsInFilesExplorer(editorHwnd)
        Editor_EnsureFilesExplorerSidebarFocused(editorHwnd)
    Send "^2"
    ShowCenteredOverlay_Utils("✅ Path copied", 1400, BANNER_ACCENT_SUCCESS)
}

; Ctrl + 5 : Context menu navigation sequence
^5::
{
    Click "Right"         ; 1. Right mouse click
    Sleep 100
    Send "{Down}"         ; 2. Press Down Arrow twice
    Send "{Down}"
    Sleep 50
    Send "{Right}"        ; 3. Press Right Arrow once
    Sleep 50
    Send "{Down}"         ; 4. Press Down Arrow once
    Sleep 50
    Send "{Right}"        ; 5. Press Right Arrow once
    Sleep 50
    Send "{Enter}"        ; 6. Press Enter
}

; Ensure only one Chrome window shows the given PDF: close any Chrome window whose title matches the filename, then open a fresh Chrome window.
EnsureSingleChromePdfInstance(filePath := "", fileNameOnly := "") {
    if (fileNameOnly = "")
        return

    ; Include all Chrome top-level windows (visible, minimized, or hidden)
    prevDetectHidden := A_DetectHiddenWindows
    DetectHiddenWindows true

    ; Collect Chrome window hwnds that appear to be showing this PDF.
    ; Title matching is unreliable because Chrome PDF viewer titles can reflect document content instead of filename.
    toClose := []
    fileNameLower := StrLower(fileNameOnly)
    filePathLower := ""
    if (filePath != "")
        filePathLower := StrLower(StrReplace(filePath, "\", "/"))
    baseNameLower := ""
    try {
        SplitPath fileNameOnly, , , &ext, &baseName
        if (baseName != "")
            baseNameLower := StrLower(baseName)
    }
    catch {
    }
    for hwnd in WinGetList("ahk_exe chrome.exe") {
        try {
            title := WinGetTitle("ahk_id " hwnd)
            if (title = "")
                continue
            titleLower := StrLower(title)
            matched := (fileNameLower != "" && InStr(titleLower, fileNameLower) != 0)
            if (!matched && baseNameLower != "")
                matched := InStr(titleLower, baseNameLower) != 0
            ; UIA-based match: look for a Document element whose Value contains the file path or filename.
            ; This catches Chrome PDF viewer windows whose title doesn't include the filename.
            matchedUia := false
            docValue := ""
            if (!matched) {
                try {
                    root := UIA.ElementFromHandle(hwnd)
                    ; Find a document-like element that exposes the PDF file URL/path.
                    docEl := root.FindFirst({ Type: "Document" })
                    if docEl {
                        try docValue := StrLower(docEl.Value)
                        catch {
                            docValue := ""
                        }
                        if (docValue != "") {
                            docValue := StrReplace(docValue, "\", "/")
                            if (filePathLower != "" && InStr(docValue, filePathLower))
                                matchedUia := true
                            else if (fileNameLower != "" && InStr(docValue, fileNameLower))
                                matchedUia := true
                        }
                    }
                } catch {
                    matchedUia := false
                }
            }

            if (!matched && matchedUia)
                matched := true
            if matched
                toClose.Push(hwnd)
        } catch {
            continue
        }
    }

    ; Restore previous DetectHiddenWindows setting
    DetectHiddenWindows prevDetectHidden

    for hwnd in toClose {
        try {
            WinActivate("ahk_id " hwnd)
            Sleep 150
            Send "!{F4}"
            Sleep 200
        } catch {
            ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
        }
    }
    Sleep 300
    ; Open a fresh, empty Chrome window. Marp will open the PDF itself
    ; in the last activated Chrome window after export completes.
    try Run "chrome.exe --new-window"
}
