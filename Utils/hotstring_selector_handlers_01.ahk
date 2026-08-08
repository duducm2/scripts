; =============================================================================
; Utils module: hotstring_selector_handlers_01.ahk
; HandleHotstringChar and selector key handlers
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
        HotstringGeminiBanner_Show("âŒ¨ Entering " . GetGlobalAIProviderLabel() . " Mode - Select prompt")
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
        ; Capture before Cleanup clears category (R→Prompts path strips after paste).
        wasPromptsCategory := (g_UtilitySelectorCategory = "Prompts")
        ; Cleanup first (closes GUI, disables hotkeys)
        CleanupHotstringSelector()

        if (useGemini) {
            companion := ResolveGlobalAICompanion()
            aiLabel := GetGlobalAIProviderLabel()
            HotstringGeminiBanner_Show("ðŸ“¤ " . aiLabel . ": inserting prompt...")
            try {
                if (companion = "enterprise") {
                    GeminiEnterprise_NavigateFocusAndPaste(expansion, false)
                    try ReplaceComposerWithStrippedReminders(expansion)
                    catch {
                    }
                } else if (companion = "copilot") {
                    CopilotWeb_NavigateFocusAndPaste(expansion, false)
                    try ReplaceComposerWithStrippedReminders(expansion)
                    catch {
                    }
                } else {
                    SetTitleMatchMode(2)
                    geminiHwnd := 0
                    try {
                        for hwnd in WinGetList("ahk_exe chrome.exe") {
                            try {
                                if IsConsumerGeminiChromeTitle(WinGetTitle("ahk_id " hwnd)) {
                                    geminiHwnd := hwnd
                                    break
                                }
                            } catch {
                            }
                        }
                    } catch {
                    }

                    if (geminiHwnd) {
                        WinActivate("ahk_id " geminiHwnd)
                        WinWaitActive("ahk_id " geminiHwnd, , 2)
                    } else {
                        WinActivate("ahk_exe chrome.exe")
                        WinWaitActive("ahk_exe chrome.exe", , 2)
                    }

                    if (char = "1" || char = "2" || char = "3") {
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
                        Send("^1")
                        Sleep 120
                        ShowSingleCharTabBanner_Utils(1)
                    }

                    try {
                        uia := UIA_Browser()
                        Gemini_FocusPromptWithChime(uia, { playChime: false, useAnchorFallback: true })
                    } catch {
                    }

                    InsertText(expansion)
                    ScriptSoundPlay(A_ScriptDir . "\assets\sounds\gemini-focused.wav")
                    try ReplaceComposerWithStrippedReminders(expansion)
                    catch {
                    }
                }
            } finally {
                HotstringGeminiBanner_Hide()
            }
            return
        }

        ; Standard behavior: paste into current active text field.
        ; Small delay to ensure target window has focus before pasting
        Sleep 150
        if (wasPromptsCategory)
            PasteStrippedPromptOfferReminders(expansion)
        else
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
