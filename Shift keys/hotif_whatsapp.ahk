; =============================================================================
; Shift keys module: hotif_whatsapp.ahk
; WhatsApp desktop hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

;-------------------------------------------------------------------
; WhatsApp Shortcuts
;-------------------------------------------------------------------
; Foreground gate: title often becomes the contact name, so WinActive("WhatsApp") misses
; the Chrome PWA. Check this hwnd only (exe, title, or app id). Cache per hwnd.
global g_WhatsApp_HotkeyHwnd := 0
global g_WhatsApp_HotkeyResult := false
global isRecording := false          ; persists between hotkey presses

IsWhatsAppShiftActive() {
    HookTiming_Begin("IsWhatsAppShiftActive")
    try
        return IsWhatsAppShiftActive_Run()
    finally
        HookTiming_End("IsWhatsAppShiftActive")
}

IsWhatsAppShiftActive_Run() {
    global g_WhatsApp_HotkeyHwnd, g_WhatsApp_HotkeyResult
    hwnd := WinExist("A")
    if (!hwnd)
        return false

    ; Cheap checks every press. Title changes when the open chat changes; hwnd does not.
    try {
        if (WinGetProcessName("ahk_id " hwnd) = "WhatsApp.exe")
            return true
    } catch {
    }
    title := ""
    try title := WinGetTitle("ahk_id " hwnd)
    if (title != "" && InStr(title, "WhatsApp"))
        return true

    ; App-id lookup is the expensive path. Cache it per hwnd.
    if (hwnd = g_WhatsApp_HotkeyHwnd && WinExist("ahk_id " hwnd))
        return g_WhatsApp_HotkeyResult
    result := false
    try result := !!WhatsAppJump_IsWhatsAppChromeAppHwnd(hwnd)
    catch
        result := false
    g_WhatsApp_HotkeyHwnd := hwnd
    g_WhatsApp_HotkeyResult := result
    return result
}

; Loading Indication while the shortcut runs, then a short success or error line.
WhatsApp_Begin(label) {
    try StandardLoadingBar_Show("⏳ WhatsApp: " label, BANNER_ACCENT_INTERMEDIATE, { passive: false })
}

WhatsApp_Ok(label) {
    try StandardLoadingBar_Update("✅ WhatsApp: " label, BANNER_ACCENT_SUCCESS)
    try StandardLoadingBar_Hide(700)
}

WhatsApp_Fail(text) {
    try StandardLoadingBar_Update(text, BANNER_ACCENT_ERROR)
    try StandardLoadingBar_Hide(1600)
}

WhatsApp_FindInTree(uia, condition) {
    if (!uia)
        return 0
    doc := 0
    try doc := uia.GetCurrentDocumentElement()
    if (doc) {
        el := 0
        try el := doc.FindElement(condition)
        if (el)
            return el
    }
    el := 0
    try el := uia.FindElement(condition)
    return el ? el : 0
}

; The visible "Unread" label is a text node. The chip is a parent tab item (same kind as all-filter).
WhatsApp_AsClickableFilter(el) {
    cur := el
    loop 5 {
        if (!cur)
            return 0
        id := ""
        typ := ""
        inv := false
        sel := false
        try id := cur.AutomationId
        try typ := cur.LocalizedControlType
        try inv := !!cur.GetPropertyValue(UIA.Property.IsInvokePatternAvailable)
        try sel := !!cur.GetPropertyValue(UIA.Property.IsSelectionItemPatternAvailable)
        if (id != "" || inv || sel || InStr(typ, "tab"))
            return cur
        parent := 0
        try parent := cur.Parent
        if (!parent || parent = cur)
            return el
        cur := parent
    }
    return el
}

; Page document, not the browser chrome. Id first (EN/PT), then the visible label.
WhatsApp_FindFilter(uia, automationId, namePattern) {
    el := WhatsApp_FindInTree(uia, { AutomationId: automationId })
    if (!el)
        el := WhatsApp_FindInTree(uia, { AutomationId: automationId, Type: "TabItem" })
    if (!el && namePattern != "")
        el := WhatsApp_FindInTree(uia, { Type: "TabItem", Name: namePattern, matchmode: "RegEx" })
    if (!el && namePattern != "")
        el := WhatsApp_FindInTree(uia, { Name: namePattern, matchmode: "RegEx" })
    return el ? WhatsApp_AsClickableFilter(el) : 0
}

WhatsApp_FilterIsOn(el) {
    if (!el)
        return false
    try {
        if (el.GetPropertyValue(UIA.Property.IsSelectionItemPatternAvailable))
            return !!el.SelectionItemPattern.IsSelected
    } catch {
    }
    try {
        if (el.GetPropertyValue(UIA.Property.IsTogglePatternAvailable))
            return el.TogglePattern.ToggleState = UIA.ToggleState.On
    } catch {
    }
    return false
}

WhatsApp_GetActiveUia() {
    hwnd := WinExist("A")
    if (!hwnd)
        return 0
    try return UIA_Browser("ahk_id " hwnd)
    catch
        return 0
}

; FindFirst rejects MatchMode RegEx; FindElement is the one-lookup path that accepts it.
WhatsApp_FindNamedButton(uia, pattern, timeoutMs := 0) {
    if (!uia)
        return 0
    deadline := A_TickCount + timeoutMs
    loop {
        el := 0
        try el := uia.FindElement({ Type: "Button", Name: pattern, matchmode: "RegEx" })
        if (el)
            return el
        if (A_TickCount >= deadline)
            return 0
        Sleep 40
    }
}

WhatsApp_InvokeOrClick(btn) {
    if (!btn)
        return false
    supportsInvoke := false
    try supportsInvoke := btn.GetPropertyValue(UIA.Property.IsInvokePatternAvailable)
    catch
        supportsInvoke := false
    if (supportsInvoke) {
        try {
            btn.Invoke()
            return true
        } catch {
        }
    }
    try {
        btn.Click()
        return true
    } catch {
    }
    return false
}

#HotIf IsWhatsAppShiftActive()

; Shift + V : Toggle voice message - Voice
+v:: ToggleVoiceMessage()

; Shift + S : Search chats - Search
+s::
{
    try {
        uia := UIA_Browser()
        saved := UiElements_TrySaved(uia, "WhatsApp", "Search")
        if (IsObject(saved)) {
            try {
                saved.SetFocus()
                return
            }
        }
    }
    WhatsApp_Begin("Search")
    Send("!k")
    WhatsApp_Ok("Search")
}

; Shift + R : Reply - Reply
+r::
{
    WhatsApp_Begin("Reply")
    Send("!r")
    WhatsApp_Ok("Reply")
}

; Shift + E : Emoji panel - Emoji
+e::
{
    WhatsApp_Begin("Emoji")
    Send("^!s")
    WhatsApp_Ok("Emoji")
}

; Shift + U : Toggle Unread filter - Unread
+u::
{
    WhatsApp_Begin("Unread filter")
    try {
        uia := WhatsApp_GetActiveUia()
        if (!uia) {
            WhatsApp_Fail("❌ WhatsApp: could not attach")
            return
        }

        unreadButton := WhatsApp_FindFilter(uia, "unread-filter", "i)^(Unread|Não lidas|Nao lidas)$")
        allButton := WhatsApp_FindFilter(uia, "all-filter", "i)^(All|Tudo|Todas)$")

        if (unreadButton && allButton) {
            if (WhatsApp_FilterIsOn(unreadButton)) {
                if (WhatsApp_InvokeOrClick(allButton))
                    WhatsApp_Ok("All chats")
                else
                    WhatsApp_Fail("❌ WhatsApp: could not show all chats")
            }
            else if (WhatsApp_InvokeOrClick(unreadButton)) {
                WhatsApp_Ok("Unread only")
            }
            else {
                WhatsApp_Fail("❌ WhatsApp: could not open Unread filter")
            }
        }
        else if (unreadButton && WhatsApp_InvokeOrClick(unreadButton)) {
            WhatsApp_Ok("Unread filter")
        }
        else {
            WhatsApp_Fail("❌ WhatsApp: could not find Unread filter")
        }
    }
    catch Error as e {
        WhatsApp_Fail("❌ WhatsApp: " e.Message)
    }
}

; Shift + F : Focus current chat - Focus
+f::
{
    WhatsApp_Begin("Focus chat")
    try {
        uia := WhatsApp_GetActiveUia()
        if (!uia) {
            WhatsApp_Fail("❌ WhatsApp: could not attach")
            return
        }

        ; Archived / Arquivadas (trailing space on the English name). Anchor, then Tab into the chat list.
        archivedButton := WhatsApp_FindNamedButton(uia, "i)(Archived|Arquivad)")
        if (!archivedButton)
            archivedButton := WhatsApp_FindInTree(uia, { Name: "i)(Archived|Arquivad)", matchmode: "RegEx" })
        if (archivedButton) {
            archivedButton.SetFocus()
            SendInput "{Tab}"
            WhatsApp_Ok("Focus chat")
        }
        else {
            WhatsApp_Fail("❌ WhatsApp: could not find Archived")
        }
    }
    catch Error as e {
        WhatsApp_Fail("❌ WhatsApp: " e.Message)
    }
}

; Shift + M : Mark as read or unread - Mark
+m::
{
    WhatsApp_Begin("Mark read")
    Send "^!+u"
    WhatsApp_Ok("Mark read")
}

; Shift + P : Pin chat or unpin chat - Pin
+p::
{
    WhatsApp_Begin("Pin")
    Send "^!+p"
    WhatsApp_Ok("Pin")
}

; ---------------------------------------------------------------------------
ToggleVoiceMessage() {
    global isRecording

    WhatsApp_Begin("Voice message")
    try {
        chrome := WhatsApp_GetActiveUia()
        if !IsObject(chrome) {
            WhatsApp_Fail("❌ WhatsApp: could not attach")
            return
        }

        ; Exact-name regexes (case-insensitive, anchored ^ $)
        voicePattern := "i)^(Voice message|Record voice message)$"
        sendPattern := "i)^(Send|Stop recording)$"

        if (isRecording) {
            ; Short poll: the send control replaces the mic after recording starts.
            if (btn := WhatsApp_FindNamedButton(chrome, sendPattern, 400)) {
                if (WhatsApp_InvokeOrClick(btn)) {
                    isRecording := false
                    WhatsApp_Ok("Voice sent")
                } else
                    WhatsApp_Fail("❌ WhatsApp: could not send voice message")
            } else {
                ; Assume you clicked Send manually > reset & start new rec
                isRecording := false
                if (btn := WhatsApp_FindNamedButton(chrome, voicePattern, 400)) {
                    if (WhatsApp_InvokeOrClick(btn)) {
                        isRecording := true
                        WhatsApp_Ok("Recording")
                    } else
                        WhatsApp_Fail("❌ WhatsApp: could not start voice message")
                } else
                    WhatsApp_Fail("❌ WhatsApp: voice button missing")
            }
        } else {
            if (btn := WhatsApp_FindNamedButton(chrome, voicePattern, 400)) {
                if (WhatsApp_InvokeOrClick(btn)) {
                    isRecording := true
                    WhatsApp_Ok("Recording")
                } else
                    WhatsApp_Fail("❌ WhatsApp: could not start voice message")
            } else
                WhatsApp_Fail("❌ WhatsApp: voice button missing")
        }
    } catch Error as err {
        WhatsApp_Fail("❌ WhatsApp: " err.Message)
    }
}

; ---------------------------------------------------------------------------
FocusSourceControlViewForCommitGeneration(delayMs := 450) {
    ; Ensure Source Control has time to render commit actions before UIA lookup.
    Send "+d"
    Sleep delayMs
}

; ---------------------------------------------------------------------------
ClickGenerateCommitMessageButton() {
    try {
        ; Ensure Source Control view is focused so the Generate button is visible.
        FocusSourceControlViewForCommitGeneration()

        ; Use UIA_Browser to get the root element (similar to other functions in the script)
        uia := UIA_Browser()
        if !IsObject(uia) {
            ; Fallback: try Ctrl+M shortcut if UIA fails
            Send "^m"
            return true
        }

        ; Find the "Generate Commit Message (Ctrl+M)" button
        ; Try multiple search strategies
        btn := uia.FindFirst({ Name: "Generate Commit Message (Ctrl+M)", ControlType: "Button" })

        ; If not found by exact name, try partial match
        if !btn {
            btn := uia.FindFirst({ Name: "Generate Commit Message", ControlType: "Button" })
        }

        ; If still not found, try by ControlType only (Type: 50000 = Button)
        if !btn {
            ; Get all buttons and find the one with the right name
            buttons := uia.FindAll({ ControlType: "Button" })
            for button in buttons {
                if InStr(button.Name, "Generate Commit Message") {
                    btn := button
                    break
                }
            }
        }

        if btn {
            btn.Click()
            return true
        } else {
            ; Fallback: try Ctrl+M shortcut
            Send "^m"
            return true
        }
    }
    catch Error as e {
        ; Fallback: try Ctrl+M shortcut if UIA fails
        Send "^m"
        return true
    }
}

; ---------------------------------------------------------------------------
; WaitForButton(root, pattern, timeout := 5000)
;   â€¢ Searches all descendant buttons of `root` until Name matches `pattern`
;   â€¢ Returns the UIA element or 0 if none matched within `timeout` ms
; ---------------------------------------------------------------------------
WaitForButton(root, pattern, timeout := 5000) {
    ; #region agent log
    SafeDebugLog Format(
        "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:1727`",`"message`":`"WaitForButton entry`",`"data`":{`"pattern`":`"{4}`",`"timeout`":{5}},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"B`"}`n",
        A_TickCount, Random(1000, 9999), A_TickCount, pattern, timeout)
    ; #endregion
    if !IsObject(root)
        return 0

    deadline := A_TickCount + timeout
    while (A_TickCount < deadline) {
        buttons := root.FindAll({ Type: "Button" })
        ; #region agent log
        SafeDebugLog Format(
            "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:1733`",`"message`":`"Found buttons count`",`"data`":{`"count`":{4}},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"B`"}`n",
            A_TickCount, Random(1000, 9999), A_TickCount, buttons.Length)
        ; #endregion

        ; Collect all matching buttons and their properties
        matchingButtons := []
        for btn in buttons {
            btnName := ""
            try btnName := btn.Name
            ; #region agent log
            if InStr(pattern, "Connect") {
                SafeDebugLog Format(
                    "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:1738`",`"message`":`"Checking button name`",`"data`":{`"name`":`"{4}`",`"pattern`":`"{5}`"},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"B`"}`n",
                    A_TickCount, Random(1000, 9999), A_TickCount, btnName, pattern)
            }
            ; #endregion
            if RegExMatch(btn.Name, pattern) {
                className := ""
                hasClassName := false
                supportsInvoke := false
                parentName := ""
                parentClass := ""

                try {
                    className := btn.ClassName
                    hasClassName := (className != "")
                } catch {
                    ; ClassName property not available or error reading
                    hasClassName := false
                }

                ; Try to capture parent info for better disambiguation (esp. duplicated "Send" buttons)
                try {
                    parent := btn.Parent
                    parentName := parent.Name
                    parentClass := parent.ClassName
                } catch {
                    parentName := ""
                    parentClass := ""
                }

                ; Check if button supports Invoke pattern
                try {
                    supportsInvoke := btn.GetPropertyValue(UIA.Property.IsInvokePatternAvailable)
                } catch {
                    supportsInvoke := false
                }

                ; #region agent log
                SafeDebugLog Format(
                    "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:1770`",`"message`":`"Matching button found`",`"data`":{`"name`":`"{4}`",`"className`":`"{5}`",`"hasClassName`":{6},`"supportsInvoke`":{7}},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"B,C`"}`n",
                    A_TickCount, Random(1000, 9999), A_TickCount, btnName, className, hasClassName, supportsInvoke)
                ; #endregion
                matchingButtons.Push({ btn: btn, hasClassName: hasClassName, className: className, supportsInvoke: supportsInvoke,
                    parentName: parentName, parentClass: parentClass })
            }
        }

        ; If we found matching buttons, prioritize: 1) hasClassName (the actual clickable button), 2) supportsInvoke, 3) first
        ; Note: In WhatsApp, the button WITH ClassName is the actual clickable one, even if it doesn't support Invoke pattern
        if matchingButtons.Length > 0 {
            bestBtn := ""
            bestClassName := ""
            bestScore := 0

            for match in matchingButtons {
                score := 0
                if match.hasClassName && match.className != "" {
                    score += 10  ; Highest priority: has ClassName (the actual clickable button in WhatsApp)
                }
                if match.supportsInvoke {
                    score += 5   ; Second priority: supports Invoke pattern
                }

                ; Additional heuristic for WhatsApp voice "Send" vs text "Send" (H8)
                ; When using the sendPattern, prefer the inner child button whose parent is also "Send"
                if InStr(pattern, "Send|Stop recording") {
                    try {
                        if (match.parentName = "Send") {
                            score += 3
                        }
                    }
                }

                if (score > bestScore) {
                    bestBtn := match.btn
                    bestClassName := match.className
                    bestScore := score
                }
            }

            ; If no button scored (shouldn't happen), use the first one
            if !bestBtn {
                bestBtn := matchingButtons[1].btn
            }

            ; #region agent log
            finalBtnName := "", finalBtnClassName := "", finalBtnType := ""
            try finalBtnName := bestBtn.Name
            try finalBtnClassName := bestBtn.ClassName
            try finalBtnType := bestBtn.ControlType
            SafeDebugLog Format(
                "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:1813`",`"message`":`"WaitForButton returning button`",`"data`":{`"name`":`"{4}`",`"className`":`"{5}`",`"type`":`"{6}`",`"score`":{7}},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"C`"}`n",
                A_TickCount, Random(1000, 9999), A_TickCount, finalBtnName, finalBtnClassName, finalBtnType, bestScore)
            ; #endregion
            return bestBtn
        }
        Sleep 50  ; reduced from 150ms to 50ms for faster retries
    }

    ; #region agent log
    SafeDebugLog Format(
        "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:1818`",`"message`":`"WaitForButton timeout - no button found`",`"data`":{`"pattern`":`"{4}`"},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"B`"}`n",
        A_TickCount, Random(1000, 9999), A_TickCount, pattern)
    ; #endregion
    return 0
}

#HotIf