; =============================================================================
; Utils module: helpers_timing_gemini.ahk
; Timing helpers and FindGeminiPromptField
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; Find the Gemini prompt field via UIA (returns element or 0). Supports EN and PT labels. Used by Gemini.ahk and Utils.ahk.
; Happy path: FindFirst per name only. FindAll({ Type: 50004 }) runs only when those fail (failure path).
FindGeminiPromptField(uia) {
    if (!IsObject(uia))
        return 0

    try {
        saved := UiElements_TrySaved(uia, "Gemini", "Prompt")
        if (saved)
            return saved
    } catch {
    }

    ; Resolve document element if uia is UIA_Browser or top-level window
    doc := uia
    try {
        if (uia.HasMethod("GetCurrentDocumentElement")) {
            d := uia.GetCurrentDocumentElement()
            if (IsObject(d))
                doc := d
        } else if (uia.HasProp("DocumentElement") && IsObject(uia.DocumentElement)) {
            doc := uia.DocumentElement
        }
    } catch {
    }

    ; 1. Fast path: Find by name in document / root
    ; 1. Fast path: Find Edit control by name in document / root (Type 50004 avoids static placeholder text)
    for name in GEMINI_PROMPT_FIELD_NAMES {
        try {
            promptField := doc.FindFirst({ Name: name, matchmode: 2, casesense: false })
            promptField := doc.FindFirst({ Name: name, Type: 50004, matchmode: 2, casesense: false })
            if (promptField)
                return promptField
        } catch {
        }
        try {
            promptField := uia.FindFirst({ Name: name, matchmode: 2, casesense: false })
            promptField := uia.FindFirst({ Name: name, Type: 50004, matchmode: 2, casesense: false })
            if (promptField)
                return promptField
        } catch {
        }
    }

    ; 2. AutomationId & ClassName known for AI companions (ChatGPT, Copilot, Gemini)
    for criteria in [
        { AutomationId: "prompt-textarea" },
        { AutomationId: "agent-search-prosemirror-editor" },
        { AutomationId: "m365-chat-editor-target-element" },
        { ClassName: "ql-editor", matchmode: 2 },
        { ClassName: "rich-textarea", matchmode: 2 },
        { ClassName: "ProseMirror", matchmode: 2 }
    ] {
        try {
            el := doc.FindFirst(criteria)
            if (el)
                return el
        } catch {
        }
        try {
            el := uia.FindFirst(criteria)
            if (el)
                return el
        } catch {
        }
    }

    ; 2b. Focusable element by name (in case a companion uses a non-50004 editable role)
    for name in GEMINI_PROMPT_FIELD_NAMES {
        try {
            promptField := doc.FindFirst({ Name: name, IsKeyboardFocusable: true, matchmode: 2, casesense: false })
            if (promptField)
                return promptField
        } catch {
        }
    }

    ; 3. Fallback: Find all edits in document and match name / class
    try {
        allEdits := doc.FindAll({ Type: 50004 })
        for edit in allEdits {
            try {
                if (InStr(edit.ClassName, "ql-editor") || InStr(edit.ClassName, "new-input-ui") || InStr(edit.ClassName, "rich-textarea"))
                    return edit
                for name in GEMINI_PROMPT_FIELD_NAMES {
                    if InStr(edit.Name, name) || InStr(edit.Name, "prompt") || InStr(edit.Name, "gemini")
                        return edit
                }
            } catch {
                continue
            }
        }
    } catch {
    }

    ; 4. Final fallback on top-level uia if distinct
    if (doc != uia) {
        try {
            allEdits := uia.FindAll({ Type: 50004 })
            for edit in allEdits {
                try {
                    if (InStr(edit.ClassName, "ql-editor") || InStr(edit.ClassName, "rich-textarea"))
                        return edit
                } catch {
                    continue
                }
            }
        } catch {
        }
    }

    return 0
}

; Play audible focused chime for AI companions (16-bit PCM cue).
Utils_PlayGeminiFocusedChime(minIntervalMs := 400) {
    static lastChimeTick := 0
    if (!IsSoundEnabled())
        return false
    now := A_TickCount
    if (lastChimeTick && (now - lastChimeTick) < minIntervalMs)
        return false
    lastChimeTick := now
    soundPath := A_ScriptDir . "\assets\sounds\ai-companion-focused.wav"
    if (!FileExist(soundPath))
        soundPath := A_ScriptDir . "\..\assets\sounds\ai-companion-focused.wav"
    if (!FileExist(soundPath))
        soundPath := A_ScriptDir . "\assets\sounds\into-cursor-textfield.wav"
    if (!FileExist(soundPath))
        soundPath := A_ScriptDir . "\assets\sounds\gemini-focused.wav"
    ScriptSoundPlay(soundPath)
    return true
}

; Play audible restored chime when returning focus from AI companion to the previous window.
Utils_PlayAICompanionRestoredChime(minIntervalMs := 400) {
    static lastRestoredTick := 0
    if (!IsSoundEnabled())
        return false
    now := A_TickCount
    if (lastRestoredTick && (now - lastRestoredTick) < minIntervalMs)
        return false
    lastRestoredTick := now
    soundPath := A_ScriptDir . "\assets\sounds\ai-companion-restored.wav"
    if (!FileExist(soundPath))
        soundPath := A_ScriptDir . "\..\assets\sounds\ai-companion-restored.wav"
    if (!FileExist(soundPath))
        return false
    ScriptSoundPlay(soundPath)
    return true
}

Gemini_PollPromptKeyboardFocus(promptField, timeoutMs := 0, pollMs := 0) {
    if (timeoutMs <= 0)
        timeoutMs := GEMINI_PROMPT_FOCUS_TIMEOUT_MS
    if (pollMs <= 0)
        pollMs := GEMINI_PROMPT_FOCUS_POLL_MS
    start := A_TickCount
    while (A_TickCount - start < timeoutMs) {
        try {
            if (promptField.HasKeyboardFocus)
                return true
        } catch {
        }
        Sleep pollMs
    }
    try
        return promptField.HasKeyboardFocus
    catch
        return false
}

Gemini_FindUploadAnchorButton(uia) {
    anchorButton := 0
    try {
        anchorButton := UiElements_TrySaved(uia, "Gemini", "Upload")
        if (anchorButton)
            return anchorButton
    } catch {
    }
    try {
        anchorButton := uia.FindFirst({ Type: UIA_ControlType_Button, Name: "Open upload file menu", ControlType: "Button" })
        if (!anchorButton)
            anchorButton := uia.FindFirst({ Type: UIA_ControlType_Button, Name: "Open upload file menu", cs: false })
    } catch {
    }
    if (!anchorButton) {
        try {
            allButtons := uia.FindAll({ Type: UIA_ControlType_Button })
            for button in allButtons {
                try {
                    if (InStr(button.Name, "Open upload file menu", false)) {
                        anchorButton := button
                        break
                    }
                } catch {
                    continue
                }
            }
        } catch {
        }
    }
    return anchorButton
}

Gemini_ApplyAnchorBacktrack(anchorButton) {
    try {
        anchorButton.SetFocus()
        Sleep GEMINI_PROMPT_FOCUS_POLL_MS
        SendInput "+{Tab}"
        Sleep GEMINI_PROMPT_FOCUS_POLL_MS
        return true
    } catch {
        return false
    }
}

; Focus Gemini prompt: direct FindFirst → bounded poll → anchor Shift+Tab fallback. Optional ready chime.
Gemini_FocusPromptWithChime(uia, options := "", &outPhase := "") {
    outPhase := "focus_failed"
    if (!IsObject(uia))
        return false

    playChime := true
    useAnchorFallback := true
    if (IsObject(options)) {
        if (options.HasProp("playChime"))
            playChime := options.playChime
        if (options.HasProp("useAnchorFallback"))
            useAnchorFallback := options.useAnchorFallback
    }

    promptField := 0
    try
        promptField := FindGeminiPromptField(uia)
    catch {
    }

    if (promptField) {
        try {
            if (promptField.HasKeyboardFocus) {
                outPhase := "fast_already_focused"
                if (playChime)
                    Utils_PlayGeminiFocusedChime()
                return promptField
            }
        } catch {
        }

        try
            promptField.ScrollIntoView()
        catch {
        }

        try
            promptField.SetFocus()
        catch {
        }

        if (Gemini_PollPromptKeyboardFocus(promptField)) {
        if (Gemini_PollPromptKeyboardFocus(promptField, 150)) {
            outPhase := "direct_focus"
            if (playChime)
                Utils_PlayGeminiFocusedChime()
            return promptField
        }

        try
            promptField.Click()
        catch {
        clicked := false
        try {
            loc := promptField.Location
            if (IsObject(loc) && loc.w > 0 && loc.h > 0) {
                promptField.Click("left")
                clicked := true
            }
        } catch {
        }
        if (!clicked) {
            try {
                promptField.Click()
                clicked := true
            } catch {
            }
        }

        if (Gemini_PollPromptKeyboardFocus(promptField, Max(300, GEMINI_PROMPT_FOCUS_TIMEOUT_MS))) {
            outPhase := "direct_focus"
            if (playChime)
                Utils_PlayGeminiFocusedChime()
            return promptField
        }

        ; Element was found and clicked/focused; even if UIA HasKeyboardFocus lags in Chromium, focus is set
        outPhase := "direct_focus"
        if (playChime)
            Utils_PlayGeminiFocusedChime()
        return promptField
        ; If UIA still reports no keyboard focus and anchor fallback is enabled, try Shift+Tab from upload anchor
        if (useAnchorFallback) {
            anchorButton := Gemini_FindUploadAnchorButton(uia)
            if (anchorButton && Gemini_ApplyAnchorBacktrack(anchorButton)) {
                if (playChime)
                    Utils_PlayGeminiFocusedChime()
                outPhase := "anchor_fallback"
                return promptField
            }
        }

        if (clicked) {
            outPhase := "direct_focus"
            if (playChime)
                Utils_PlayGeminiFocusedChime()
            return promptField
        }
        return false
    }

    ; Anchor fallback if promptField was not found directly
    if (useAnchorFallback) {
        anchorButton := Gemini_FindUploadAnchorButton(uia)
        if (anchorButton && Gemini_ApplyAnchorBacktrack(anchorButton)) {
            try
                promptField := FindGeminiPromptField(uia)
            catch {
            }
            if (promptField) {
                try
                    promptField.SetFocus()
                catch {
                }
                if (playChime)
                    Utils_PlayGeminiFocusedChime()
                outPhase := "anchor_fallback"
                return promptField
            }
            if (playChime)
                Utils_PlayGeminiFocusedChime()
            outPhase := "anchor_fallback"
            return anchorButton
        }
    }

    return false
}

; Shared by Gemini.ahk #!+i, Shift keys Fast Copy, and async flows. Resolves UIA when omitted.
Gemini_FocusPromptSameAsOpenHotkey(uia, playChime := true) {
    if (!IsObject(uia)) {
        try
            uia := UIA_Browser()
        catch
            return false
    }
    return Gemini_FocusPromptWithChime(uia, { playChime: playChime, useAnchorFallback: true })
}

Gemini_WaitForPromptFieldDiscoverable(uia, timeoutMs := 0) {
    if (timeoutMs <= 0)
        timeoutMs := GEMINI_PROMPT_FOCUS_TIMEOUT_MS + 200
    start := A_TickCount
    while (A_TickCount - start < timeoutMs) {
        try {
            if (FindGeminiPromptField(uia))
                return true
        } catch {
        }
        Sleep GEMINI_PROMPT_FOCUS_POLL_MS
    }
    return false
}

Gemini_ShowDeferredTabBanner(uia) {
    ; Tab indicator banner permanently removed per user request.
    return
}

FocusGeminiAskFieldForHwnd(geminiHwnd, playChime := false) {
    if (!geminiHwnd)
        return false
    try {
        if (!WinActive("ahk_id " geminiHwnd))
            return false
    } catch {
        return false
    }
    try {
        uia := UIA_Browser("ahk_id " geminiHwnd)
        result := Gemini_FocusPromptWithChime(uia, { playChime: playChime, useAnchorFallback: true })
        return !!result
    } catch {
    }
    return false
}

; 1-based active Chrome tab index via UIA (tab bar). Returns { index, count } or 0. Shared with Gemini.ahk.
GetChromeActiveTabIndex(uia) {
    try {
        uia.GetCurrentMainPaneElement()
        tabs := uia.GetTabs()
        if (!tabs.Length)
            return 0
        current := uia.GetTab("")
        if (!current)
            return 0
        rid := current.RuntimeId
        for i, tab in tabs {
            try {
                if (tab.RuntimeId = rid)
                    return { index: i, count: tabs.Length }
            } catch {
                continue
            }
        }
    } catch {
    }
    return 0
}
