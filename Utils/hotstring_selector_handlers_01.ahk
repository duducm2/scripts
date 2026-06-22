; =============================================================================
; Utils module: hotstring_selector_handlers_01.ahk
; HandleHotstringChar and Gemini paste helpers
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; HandleHotstringChar(char)
; =============================================================================
; PURPOSE: Processes character key press events when hotstring selector is active.
;          Executes the action (text expansion, file open, or macro) associated with the character.
;
; PARAMETERS:
;   char: String - Single character that was pressed (e.g., "a", "1", ",")
;
; EXECUTION ORDER:
;   1. Special cases: Characters "9" and "0" trigger Miro board activation (hardcoded)
;   2. Quick-open files: Check g_QuickOpenFileCharMap for file path mappings
;   3. Macros: Check g_MacroCharMap for executable macro functions
;   4. Hotstrings: Check g_HotstringCharMap for text expansion mappings
;
; BEHAVIOR:
;   - Performs case-insensitive lookup (tries both original and lowercase)
;   - Closes selector GUI before executing action
;   - For hotstrings: Inserts text using InsertText() after 150ms delay
;   - For files: Opens file based on extension (Power BI files use special handler)
;   - For macros: Executes macro function directly
;
; RETURNS: None (void function)
; =============================================================================

; Navigate to Gemini, focus the prompt field, then paste first clipboard snippet (same as Win+Alt+Shift+1).
; Reference: "order called snippets" - Clip Angel top item via ClipAngel_SendNativeTopItemKeys (Win+Alt+Shift+1 sequence).
; If optionalPromptText is non-empty, inserts that text into the prompt field instead (same as Win+Alt+Shift+U then L, prompt char).
; switchToFirstTab: when true (default), send Ctrl+1 and show tab-1 banner (AI Text Optimizer / ^!#4). When false, use currently active Gemini tab if any, else first Gemini window, without changing tab (delay-submit flow).
GeminiNavigateFocusAndPasteFirstSnippet(optionalPromptText := "", switchToFirstTab := true) {
    SetTitleMatchMode(2)
    geminiHwnd := 0
    if (!switchToFirstTab) {
        ; Prefer the currently active window if it is already a Gemini tab (do not switch tabs).
        try {
            activeHwnd := WinExist("A")
            if (activeHwnd && WinGetProcessName("ahk_id " activeHwnd) = "chrome.exe" && InStr(WinGetTitle("ahk_id " activeHwnd
            ), "gemini", false))
                geminiHwnd := activeHwnd
        } catch {
        }
    }
    if (!geminiHwnd) {
        try {
            for hwnd in WinGetList("ahk_exe chrome.exe") {
                try {
                    if InStr(WinGetTitle("ahk_id " hwnd), "gemini", false) {
                        geminiHwnd := hwnd
                        break
                    }
                } catch {
                }
            }
        } catch {
        }
    }

    if (!geminiHwnd) {
        ; Navigate to Gemini website
        Run "chrome.exe --new-window https://gemini.google.com/"
        if !WinWaitActive("ahk_exe chrome.exe", , 5)
            return
        Sleep 2500  ; Allow page to load
        geminiHwnd := WinExist("A")
    }

    if (geminiHwnd) {
        WinActivate("ahk_id " geminiHwnd)
        WinWaitActive("ahk_id " geminiHwnd, , 2)
        Sleep 350  ; Let UI finish processing the window transition before focus/paste
    } else {
        WinActivate("ahk_exe chrome.exe")
        WinWaitActive("ahk_exe chrome.exe", , 2)
        Sleep 350
    }

    if (switchToFirstTab) {
        ; Switch to first Gemini tab (Ctrl+1) and show number-one banner (consistent with Gemini.ahk)
        Send("^1")
        Sleep 280
        ShowSingleCharTabBanner_Utils(1)
    }

    ; Focus the Gemini prompt field (Anchor & Backtrack strategy)
    try {
        uia := UIA_Browser()
        Sleep 80
        anchorButton := 0
        try {
            anchorButton := uia.FindFirst({ Type: "50000", Name: "Open upload file menu", ControlType: "Button" })
            if (!anchorButton)
                anchorButton := uia.FindFirst({ Type: "50000", Name: "Open upload file menu", cs: false })
        } catch {
        }
        if (!anchorButton) {
            try {
                allButtons := uia.FindAll({ Type: "50000" })
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
        if (anchorButton) {
            try {
                anchorButton.SetFocus()
                Sleep 25
                SendInput "+{Tab}"
                Sleep 15
            } catch {
                try {
                    promptField := FindGeminiPromptField(uia)
                    if (promptField)
                        promptField.SetFocus()
                } catch as e {
                }
            }
        } else {
            try {
                promptField := FindGeminiPromptField(uia)
                if (promptField)
                    promptField.SetFocus()
            } catch as e {
            }
        }
    } catch {
    }

    ; Explicitly target Gemini window again before paste so paste goes to Gemini, not the trigger window
    if (geminiHwnd && WinExist("ahk_id " geminiHwnd)) {
        WinActivate("ahk_id " geminiHwnd)
        WinWaitActive("ahk_id " geminiHwnd, , 2)
        Sleep 150
    }

    if (optionalPromptText != "") {
        InsertText(optionalPromptText)
    } else {
        ; Paste first clipboard snippet (same as Win+Alt+Shift+1: order called snippets)
        ClipAngel_SendTopListItem(geminiHwnd)
    }
    ; Brief delay so paste is received and UI/character limits register before any submit or focus change
    Sleep 250
    ; Same sound as when opening Gemini (focus/paste feedback)
    ScriptSoundPlay(A_ScriptDir . "\assets\sounds\gemini-focused.wav")
}

; Returns grammar preset (from assets/prompt/grammar.txt). Matches InitHotstringsCheatSheet for :o:cgrammar.
GetGrammarPromptText() {
    return GetPromptText("grammar")
}

; Returns the AI Text Optimizer prompt text (from assets/prompt/aiopt.txt). Used by Ctrl+Alt+Win+4 and L+4 flow.
GetAioptPromptText() {
    return GetPromptText("aiopt")
}

; Delayed submit flow: show 4s banner, allow N to cancel auto-submit; then navigate+paste and optionally send Enter.
; Tell Gemini.ahk to start background completion monitor (must match WM_START_DELAYED_SUBMIT_MONITOR in Gemini.ahk).
GeminiDelayedSubmitMonitorStartFromUtils(originalHwnd, geminiChromeHwnd) {
    WM_START_DELAYED_SUBMIT_MONITOR := 0x8002
    targetHwnd := GetGeminiScriptMsgTargetHwnd()
    if (targetHwnd) {
        try SendMessage(WM_START_DELAYED_SUBMIT_MONITOR, originalHwnd, geminiChromeHwnd, , "ahk_id " targetHwnd)
    }
}

; Tell Gemini.ahk to stop the delayed-submit monitor so "Copy response?" is not shown (when user chose S or N at 5s).
GeminiDelayedSubmitMonitorStopFromUtils() {
    WM_STOP_DELAYED_SUBMIT_MONITOR := 0x8003
    targetHwnd := GetGeminiScriptMsgTargetHwnd()
    if (targetHwnd) {
        try SendMessage(WM_STOP_DELAYED_SUBMIT_MONITOR, 0, 0, , "ahk_id " targetHwnd)
    }
}

; Paste transcription to Gemini prompt only (no Enter, no 4s banner). Used when user presses S at 5s dictation confirm.
DEPRECATED_GeminiDictationPasteOnlyFlow() {
    restoreHwnd := WinExist("A")
    GeminiNavigateFocusAndPasteFirstSnippet("", false)
    if (restoreHwnd && WinExist("ahk_id " restoreHwnd))
        WinActivate("ahk_id " restoreHwnd)
}

DEPRECATED_GeminiDelayedSubmitFlow() {
    global g_HotstringGeminiAutoSubmit, g_HotstringGeminiRestoreHwnd
    g_HotstringGeminiRestoreHwnd := WinExist("A")
    g_HotstringGeminiAutoSubmit := true
    GeminiFinalizeSubmit()
}

; Delay (ms) after paste and before Send Enter in Gemini delayed-submit flow. Prevents premature send and lets the UI register paste + character limits.
global g_GeminiDelayedSubmit_PreEnterDelayMs := 1000
; Max ms to wait for prompt field to show content before sending Enter (guarantee layer). Poll interval = 200 ms.
global g_GeminiDelayedSubmit_WaitContentMaxMs := 5000

; Returns non-empty trimmed text if Gemini prompt field has content (Value or TextPattern); "" on failure or empty. Used to guarantee message is present before submit.
GeminiPromptFieldGetText() {
    try {
        uia := UIA_Browser()
        pf := FindGeminiPromptField(uia)
        if (!pf)
            return ""
        try {
            text := Trim(pf.Value)
            if (text != "")
                return text
        } catch {
        }
        try {
            text := Trim(pf.TextPattern.DocumentRange.GetText(-1))
            if (text != "")
                return text
        } catch {
        }
    } catch {
    }
    return ""
}

GeminiFinalizeSubmit() {
    global g_HotstringGeminiAutoSubmit, g_HotstringGeminiRestoreHwnd, g_GeminiDelayedSubmit_PreEnterDelayMs,
        g_GeminiDelayedSubmit_WaitContentMaxMs

    try Hotkey("n", "Off")
    try Hotkey("N", "Off")
    try Hotkey("y", "Off")
    try Hotkey("Y", "Off")
    HotstringGeminiBanner_Hide()

    ; Delay-submit flow: do not switch tabs; paste to currently active Gemini tab
    GeminiNavigateFocusAndPasteFirstSnippet("", false)

    didAutoSubmit := false
    geminiChromeHwnd := 0
    if (g_HotstringGeminiAutoSubmit) {
        ; Execution delay so paste is fully received and UI/character limits register before submit
        Sleep (g_GeminiDelayedSubmit_PreEnterDelayMs)
        ; Guarantee layer: wait until prompt field has content (or timeout) so we don't send Enter prematurely
        pollIntervalMs := 200
        endTick := A_TickCount + g_GeminiDelayedSubmit_WaitContentMaxMs
        contentFound := false
        while (A_TickCount < endTick) {
            if (GeminiPromptFieldGetText() != "") {
                contentFound := true
                break
            }
            Sleep pollIntervalMs
        }
        Send("{Enter}")
        geminiChromeHwnd := WinExist("A")
        didAutoSubmit := true
    }

    g_HotstringGeminiAutoSubmit := true

    ; Return focus to the window the user had before paste (whether Enter was sent or not)
    if (g_HotstringGeminiRestoreHwnd && WinExist("ahk_id " g_HotstringGeminiRestoreHwnd)) {
        WinActivate("ahk_id " g_HotstringGeminiRestoreHwnd)
    }

    ; If we auto-submitted (user did not cancel), ask Gemini.ahk to monitor for completion and show "Copy? [N] [R]" when done
    if (didAutoSubmit && geminiChromeHwnd)
        GeminiDelayedSubmitMonitorStartFromUtils(g_HotstringGeminiRestoreHwnd, geminiChromeHwnd)
}

HandleHotstringChar(char) {
    global g_HotstringSelectorActive, g_HotstringCharMap, g_QuickOpenFileCharMap, g_MacroCharMap
    global g_HotstringPromptCharMap, g_HotstringGeminiArmed
    global g_UtilitySelectorMode, g_UtilityTopCategoryById

    ; Only process if selector is active
    if (!g_HotstringSelectorActive) {
        return
    }

    ; Top-level category selection (R/P/M/G/L/H) or Context browser (C)
    if (g_UtilitySelectorMode = "top") {
        ch := StrLower(char)
        if (ch = "c") {
            CleanupHotstringSelector()
            ShowContextBrowser()
            return
        }
        if (g_UtilityTopCategoryById.Has(ch)) {
            category := g_UtilityTopCategoryById[ch]
            if (category = "Context")
                return
            UtilitySelector_SwitchToCategory(category)
        }
        return
    }

    ; L key: first press = arm Gemini mode (show banner); second press (double-tap) = navigate to Gemini, focus field, paste first snippet.
    if (char = "l" || char = "L") {
        if (g_HotstringGeminiArmed) {
            ; Double-tap L: delayed submit flow (paste + Enter to Gemini).
            CleanupHotstringSelector()
            D2C_FlowManager.GetInstance().StartFromHotstring()
            g_HotstringGeminiArmed := false
            return
        }
        g_HotstringGeminiArmed := true
        ; Show banner when entering Gemini mode (same pattern as Project Selector "Entering Selection Mode").
        HotstringGeminiBanner_Show("⌨ Entering Gemini Mode - Select prompt")
        SetTimer(HotstringGeminiBanner_Hide, -1500)  ; Hide banner after 1.5 s
        SetTimer(DisarmHotstringGeminiMode, -4000)
        return
    }

    ; Consume the armed state on the next selection (any selection), but only redirect Prompts.
    useGemini := false
    if (g_HotstringGeminiArmed) {
        useGemini := g_HotstringPromptCharMap.Has(StrLower(char)) || g_HotstringPromptCharMap.Has(char)
        g_HotstringGeminiArmed := false
    }

    ; Special handling for Miro boards (characters "9" and "0")
    ; Gate to Links category so other views can't trigger it.
    global g_UtilitySelectorCategory
    if (g_UtilitySelectorCategory = "Links") {
        if (char = "9") {
            CleanupHotstringSelector()
            FindAndActivateMiroWindow("https://miro.com/app/board/uXjVJdbNFkA=/", "CIP & UX Integration")
            return
        } else if (char = "0") {
            CleanupHotstringSelector()
            FindAndActivateMiroWindow("https://miro.com/app/board/uXjVJVZSXvk=/", "CIP Dashboard")
            return
        }
    }

    ; Category-scoped dispatch (prevents cross-menu execution when chars overlap)
    global g_UtilityHotstringCharMapByCategory, g_UtilitySelectorCategory

    ResolveExpansion() {
        exp := ""
        try {
            if (IsObject(g_UtilityHotstringCharMapByCategory) && g_UtilityHotstringCharMapByCategory.Has(
                g_UtilitySelectorCategory)) {
                exp := g_UtilityHotstringCharMapByCategory[g_UtilitySelectorCategory].Get(char, "")
                if (exp = "")
                    exp := g_UtilityHotstringCharMapByCategory[g_UtilitySelectorCategory].Get(StrLower(char), "")
            }
        } catch {
            exp := ""
        }
        if (exp = "") {
            exp := g_HotstringCharMap.Get(char, "")
            if (exp = "")
                exp := g_HotstringCharMap.Get(StrLower(char), "")
        }
        return exp
    }

    ResolveFilePath() {
        fp := g_QuickOpenFileCharMap.Get(char, "")
        if (fp = "")
            fp := g_QuickOpenFileCharMap.Get(StrLower(char), "")
        return fp
    }

    ResolveMacro() {
        fn := g_MacroCharMap.Get(char, "")
        if (fn = "")
            fn := g_MacroCharMap.Get(StrLower(char), "")
        return fn
    }

    TryRunFile(fp) {
        if (fp = "")
            return false
        CleanupHotstringSelector()
        SplitPath(fp, , , &ext)
        ext := StrLower(ext)
        if (ext = "pbix") {
            FindAndActivatePowerBIFile(fp)
        } else {
            fpTrim := Trim(fp)
            if (SubStr(fpTrim, 1, 8) = "https://" || SubStr(fpTrim, 1, 7) = "http://") {
                StudyLink_OpenUrlInChrome(fpTrim, true)
            } else {
                try Run(fp)
                catch {
                }
            }
        }
        return true
    }

    TryRunMacro(fn) {
        if (fn = "")
            return false
        CleanupHotstringSelector()
        try fn()
        catch {
        }
        return true
    }

    expansion := ""
    filePath := ""
    macroFunc := ""

    if (g_UtilitySelectorCategory = "Links") {
        filePath := ResolveFilePath()
        if (TryRunFile(filePath))
            return
        ; fallback for unexpected collisions
        expansion := ResolveExpansion()
    } else if (g_UtilitySelectorCategory = "Macros") {
        macroFunc := ResolveMacro()
        if (TryRunMacro(macroFunc))
            return
        expansion := ResolveExpansion()
    } else {
        ; Projects / Prompts / Hotstrings / General: hotstring-first
        expansion := ResolveExpansion()
        if (expansion = "") {
            macroFunc := ResolveMacro()
            if (TryRunMacro(macroFunc))
                return
            filePath := ResolveFilePath()
            if (TryRunFile(filePath))
                return
        }
    }

    if (expansion != "") {
        ; Cleanup first (closes GUI, disables hotkeys)
        CleanupHotstringSelector()

        if (useGemini) {
            ; L+Prompt selection: redirect to Gemini (focus prompt field, paste, do NOT submit).
            HotstringGeminiBanner_Show("📤 Gemini: inserting prompt...")
            try {
                SetTitleMatchMode(2)
                geminiHwnd := 0
                try {
                    for hwnd in WinGetList("ahk_exe chrome.exe") {
                        try {
                            if InStr(WinGetTitle("ahk_id " hwnd), "gemini", false) {
                                geminiHwnd := hwnd
                                break
                            }
                        } catch {
                            ; Skip invalid windows
                        }
                    }
                } catch {
                    ; Ignore WinGetList errors
                }

                if (geminiHwnd) {
                    WinActivate("ahk_id " geminiHwnd)
                    WinWaitActive("ahk_id " geminiHwnd, , 2)
                } else {
                    ; Per your preference: fallback to any Chrome window if Gemini isn't found.
                    WinActivate("ahk_exe chrome.exe")
                    WinWaitActive("ahk_exe chrome.exe", , 2)
                }

                ; Explicitly target Gemini tabs based on character:
                ; - L+1/2/3  -> Tab 2 (right Gemini tab, temporary prompts)
                ; - L+4/5/Q/W/E/R/T/A -> Tab 1 (left Gemini tab, main workflow)
                if (char = "1" || char = "2" || char = "3") {
                    ; Chrome convention: Ctrl+2 selects the second tab in the window.
                    Send("^2")
                    Sleep 120
                    ShowSingleCharTabBanner_Utils(2)
                } else if (char = "4" || char = "5"
                    || char = "q" || char = "Q"
                    || char = "w" || char = "W"
                    || char = "e" || char = "E"
                    || char = "r" || char = "R"
                    || char = "t" || char = "T"
                    || char = "a" || char = "A") {
                    ; Ensure Tab 1 is active before inserting the prompt.
                    Send("^1")
                    Sleep 120
                    ShowSingleCharTabBanner_Utils(1)
                }

                ; Focus the Gemini prompt field (shared helper; no chime — paste path plays its own sound).
                try {
                    uia := UIA_Browser()
                    Gemini_FocusPromptWithChime(uia, { playChime: false, useAnchorFallback: true })
                } catch {
                    ; If focus fails, we still attempt to paste (user can click manually).
                }

                ; Paste the text (do NOT send Enter)
                InsertText(expansion)
                ; Same sound as when opening Gemini (focus/paste feedback)
                ScriptSoundPlay(A_ScriptDir . "\assets\sounds\gemini-focused.wav")
            } finally {
                HotstringGeminiBanner_Hide()
            }
            return
        }

        ; Standard behavior: paste into current active text field.
        ; Small delay to ensure target window has focus before pasting
        Sleep 150
        InsertText(expansion)
    }
}

; =============================================================================
; CreateHotstringCharHandler(char)
; =============================================================================
; PURPOSE: Factory function that creates a hotkey handler function with proper closure over character value.
;          Required because AutoHotkey hotkey handlers need unique function instances per character.
;
; PARAMETERS:
;   char: String - Character to create handler for
;
; RETURNS: Function object that calls HandleHotstringChar(char) when invoked
; =============================================================================
DisarmHotstringGeminiMode(*) {
    global g_HotstringGeminiArmed
    g_HotstringGeminiArmed := false
}

CreateHotstringCharHandler(char) {
    ; Return a function that captures the char value at creation time via closure
    return (*) => HandleHotstringChar(char)
}

; =============================================================================
; HandleHotstringEscape(*)
; =============================================================================
; PURPOSE: Handles Escape key press to close hotstring selector without executing any action.
;
; PARAMETERS: None (varargs function signature for hotkey compatibility)
; RETURNS: None (void function)
; =============================================================================
HandleHotstringEscape(*) {
    global g_HotstringSelectorActive
    if (g_HotstringSelectorActive) {
        CleanupHotstringSelector()
    }
}
