; =============================================================================
; Utils module: hotstring_selector_handlers_01.ahk
; HandleHotstringChar, Gemini paste, Escape
; =============================================================================

HandleHotstringChar(char) {
    global g_HotstringSelectorActive, g_HotstringPromptCharMap, g_HotstringGeminiArmed
    global g_UtilitySelectorMode, g_UtilityTopCategoryById, g_UtilitySelectorCategory
    global g_MacroCharMap

    if (!g_HotstringSelectorActive)
        return

    if (g_UtilitySelectorMode = "top") {
        ch := StrLower(char)
        if (g_UtilityTopCategoryById.Has(ch))
            UtilitySelector_SwitchToCategory(g_UtilityTopCategoryById[ch])
        return
    }

    if (g_UtilitySelectorCategory = "Prompts" && (char = "l" || char = "L")) {
        if (g_HotstringGeminiArmed) {
            CleanupHotstringSelector()
            D2C_FlowManager.GetInstance().StartFromHotstring()
            g_HotstringGeminiArmed := false
            return
        }
        g_HotstringGeminiArmed := true
        HotstringGeminiBanner_Show("⌨ Entering " . GetGlobalAIProviderLabel() . " Mode - Select prompt")
        SetTimer(HotstringGeminiBanner_Hide, -1500)
        SetTimer(DisarmHotstringGeminiMode, -4000)
        return
    }

    useGemini := false
    if (g_HotstringGeminiArmed) {
        useGemini := (g_UtilitySelectorCategory = "Prompts") && (g_HotstringPromptCharMap.Has(StrLower(char)) ||
        g_HotstringPromptCharMap.Has(char))
        g_HotstringGeminiArmed := false
    }

    ch := StrLower(char)
    if (g_UtilitySelectorCategory = "Prompts") {
        prompt := PromptData_FindByChar(ch)
        if (!IsObject(prompt))
            return
        UtilitySelector_InsertPrompt(prompt, useGemini)
        return
    }

    if (g_UtilitySelectorCategory = "Projects") {
        project := ""
        for row in UtilitySelector_ProjectRows() {
            if (row.HasProp("char") && row.char = ch) {
                project := row
                break
            }
        }
        if (!IsObject(project) || project.name = "")
            return
        CleanupHotstringSelector()
        UtilitySelector_RestorePreviousHwnd()
        Sleep 150
        InsertText(project.name)
        return
    }

    if (g_UtilitySelectorCategory = "Hotstrings") {
        item := HotstringData_FindByChar(ch)
        if (!IsObject(item))
            return
        CleanupHotstringSelector()
        UtilitySelector_RestorePreviousHwnd()
        Sleep 150
        InsertText(item.text)
        return
    }

    if (g_UtilitySelectorCategory = "Macros") {
        if (!IsObject(g_MacroCharMap) || g_MacroCharMap.Count = 0)
            BuildMacroCharMap()
        fn := g_MacroCharMap.Get(ch, "")
        if (fn = "")
            fn := g_MacroCharMap.Get(char, "")
        if (fn = "")
            return
        CleanupHotstringSelector()
        try fn()
        catch {
        }
    }
}

UtilitySelector_InsertPrompt(prompt, useGemini := false) {
    if (!IsObject(prompt))
        return
    body := PromptData_ReadBody(prompt)
    CleanupHotstringSelector()
    if (useGemini) {
        UtilitySelector_PastePromptToGemini(body, prompt)
        return
    }
    UtilitySelector_RestorePreviousHwnd()
    Sleep 150
    UtilitySelector_AttachPromptContextFiles(prompt)
    PasteStrippedPromptOfferReminders(body)
}

UtilitySelector_AttachPromptContextFiles(prompt) {
    if (!IsObject(prompt))
        return
    paths := PromptData_ContextFilesForCurrentEnv(prompt)
    if (paths.Length = 0)
        return
    existing := []
    missing := []
    for p in paths {
        if (Clipboard_PathIsExistingFile(p))
            existing.Push(p)
        else
            missing.Push(p)
    }
    if (missing.Length > 0) {
        label := missing.Length = 1 ? missing[1] : (missing.Length . " context files")
        ShowCenteredOverlay_Utils("⚠ Missing context file(s): " . label, 2200, BANNER_ACCENT_ERROR)
    }
    if (existing.Length = 0)
        return
    if !InsertFiles(existing)
        ShowCenteredOverlay_Utils("⚠ Could not attach context files", 2200, BANNER_ACCENT_ERROR)
}

UtilitySelector_PastePromptToGemini(expansion, prompt := false) {
    companion := ResolveGlobalAICompanion()
    aiLabel := GetGlobalAIProviderLabel()
    HotstringGeminiBanner_Show("📤 " . aiLabel . ": inserting prompt...")
    try {
        if (companion = "enterprise") {
            GeminiEnterprise_OpenOrFocus()
            UtilitySelector_AttachPromptContextFiles(prompt)
            InsertText(expansion)
            try ReplaceComposerWithStrippedReminders(expansion)
            catch {
            }
        } else if (companion = "copilot") {
            CopilotWeb_OpenOrFocus()
            UtilitySelector_AttachPromptContextFiles(prompt)
            InsertText(expansion)
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

            try {
                uia := UIA_Browser()
                Gemini_FocusPromptWithChime(uia, { playChime: false, useAnchorFallback: true })
            } catch {
            }

            UtilitySelector_AttachPromptContextFiles(prompt)
            InsertText(expansion)
            ScriptSoundPlay(A_ScriptDir . "\assets\sounds\gemini-focused.wav")
            try ReplaceComposerWithStrippedReminders(expansion)
            catch {
            }
        }
    } finally {
        HotstringGeminiBanner_Hide()
    }
}

UtilitySelector_RestorePreviousHwnd() {
    global g_UtilitySelectorRestoreHwnd
    hwnd := g_UtilitySelectorRestoreHwnd
    if (!hwnd)
        return
    try {
        if (DllCall("IsWindow", "ptr", hwnd)) {
            WinActivate("ahk_id " hwnd)
            WinWaitActive("ahk_id " hwnd, , 1)
        }
    } catch {
    }
}

DisarmHotstringGeminiMode(*) {
    global g_HotstringGeminiArmed
    g_HotstringGeminiArmed := false
}

CreateHotstringCharHandler(char) {
    return (*) => HandleHotstringChar(char)
}

HandleHotstringEscape(*) {
    global g_HotstringSelectorActive
    if (g_HotstringSelectorActive) {
        CleanupHotstringSelector()
        return true
    }
    return false
}
