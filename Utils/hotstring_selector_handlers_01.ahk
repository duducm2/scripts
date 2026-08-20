; =============================================================================
; Utils module: hotstring_selector_handlers_01.ahk
; HandleHotstringChar, Gemini paste, Escape
; =============================================================================

global g_PromptContextTempFiles := []

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
    body := PromptRender_Prepare(prompt)
    if (body = "")
        return
    PromptUsage_Log(prompt, useGemini ? "gemini" : "direct")
    mode := PromptData_PasteMode(prompt)
    doAttach := (mode = "default" || mode = "body_attach_clipboard" || mode = "attach_only")
    doPasteBody := (mode = "default" || mode = "body_only" || mode = "body_plus_clipboard" || mode =
        "body_attach_clipboard"
        || mode = "auto_send")
    doAppendClipboard := (mode = "body_plus_clipboard")
    doAutoSend := (mode = "auto_send")
    CleanupHotstringSelector()
    if (useGemini) {
        UtilitySelector_PastePromptToGemini(body, prompt, doAttach, doPasteBody, doAutoSend)
        return
    }
    UtilitySelector_RestorePreviousHwnd()
    Sleep 150
    if (doAttach)
        UtilitySelector_AttachPromptContextFiles(prompt)
    if (doPasteBody)
        PasteStrippedPromptOfferReminders(body)
    if (doAppendClipboard) {
        clip := ""
        try clip := A_Clipboard
        catch {
        }
        if (clip != "")
            InsertText(clip)
    }
}

UtilitySelector_AttachPromptContextFiles(prompt) {
    if (!IsObject(prompt))
        return
    entries := PromptData_ContextEntriesForCurrentEnv(prompt)
    if (entries.Length = 0)
        return
    existing := []
    missing := []
    for e in entries {
        p := PromptData_ContextEntryPath(e)
        if (Clipboard_PathIsExistingFile(p))
            existing.Push(e)
        else
            missing.Push(p)
    }
    if (missing.Length > 0) {
        label := missing.Length = 1 ? missing[1] : (missing.Length . " context files")
        ShowCenteredOverlay_Utils("⚠ Missing context file(s): " . label, 2200, BANNER_ACCENT_ERROR)
    }
    if (existing.Length = 0)
        return
    attachPaths := PromptContext_ResolveAttachPaths(existing)
    if (attachPaths.Length = 0)
        return
    if !InsertFiles(attachPaths) {
        ShowCenteredOverlay_Utils("⚠ Could not attach context files", 2200, BANNER_ACCENT_ERROR)
        return
    }
    PromptContext_WaitForAttachUploadIdle(attachPaths.Length)
}

PromptContext_CmdQuote(s) {
    return '"' . s . '"'
}

PromptContext_ScriptPath() {
    return A_ScriptDir "\infra\python\context_compact.py"
}

PromptContext_TempDir() {
    dir := A_Temp "\prompt-context\" A_TickCount
    try DirCreate(dir)
    catch {
    }
    return dir
}

PromptContext_ScheduleTempCleanup(paths) {
    global g_PromptContextTempFiles
    if (!IsObject(g_PromptContextTempFiles))
        g_PromptContextTempFiles := []
    for p in paths
        g_PromptContextTempFiles.Push(p)
    ; Keep staged copies until Gemini/Enterprise finishes multi-file upload.
    try SetTimer(PromptContext_CleanupTemps, -60000)
    catch {
    }
}

PromptContext_CleanupTemps(*) {
    global g_PromptContextTempFiles
    if (!IsObject(g_PromptContextTempFiles))
        return
    for p in g_PromptContextTempFiles {
        try FileDelete(p)
        catch {
        }
        SplitPath p, , &dir
        if (dir != "" && InStr(dir, "\prompt-context\")) {
            try DirDelete(dir)
            catch {
            }
        }
    }
    g_PromptContextTempFiles := []
}

PromptContext_RunCompact(src, dst, compact, csvFrom, csvTo) {
    script := PromptContext_ScriptPath()
    if (!FileExist(script))
        return false
    cmd := "python " . PromptContext_CmdQuote(script) . " --in " . PromptContext_CmdQuote(src) . " --out " .
    PromptContext_CmdQuote(dst)
    if (compact)
        cmd .= " --compact"
    if (csvFrom >= 1 && csvTo >= 1)
        cmd .= " --csv-keep " . csvFrom . " " . csvTo
    exitCode := 1
    try exitCode := RunWait(cmd, A_ScriptDir, "Hide")
    catch {
        return false
    }
    return (exitCode = 0 && FileExist(dst))
}

; Basename for staged attach: keep name; .ini → .txt for Gemini upload compatibility.
PromptContext_StagedAttachName(path, usedMap) {
    SplitPath path, &name, , &ext, &nameNoExt
    if (nameNoExt = "")
        nameNoExt := (name != "") ? name : "file"
    if (StrLower(ext) = "ini")
        name := nameNoExt ".txt"
    else if (ext != "")
        name := nameNoExt "." ext
    else
        name := nameNoExt
    key := StrLower(name)
    if (!usedMap.Has(key)) {
        usedMap[key] := true
        return name
    }
    i := 2
    loop {
        if (StrLower(ext) = "ini" || ext = "")
            cand := nameNoExt "-" i ".txt"
        else
            cand := nameNoExt "-" i "." ext
        ck := StrLower(cand)
        if (!usedMap.Has(ck)) {
            usedMap[ck] := true
            return cand
        }
        i += 1
    }
}

; Copy to local temp (Drive-safe). Returns staged path or "" on failure.
PromptContext_StageLocalCopy(src, tempDir, usedMap) {
    if (src = "" || !Clipboard_PathIsExistingFile(src))
        return ""
    if (tempDir = "")
        return ""
    outName := PromptContext_StagedAttachName(src, usedMap)
    dst := tempDir "\" outName
    if (StrLower(dst) = StrLower(src))
        return src
    try {
        FileCopy(src, dst, 1)
    } catch {
        return ""
    }
    if !Clipboard_PathIsExistingFile(dst)
        return ""
    return dst
}

PromptContext_ResolveAttachPaths(entries) {
    paths := []
    temps := []
    usedNames := Map()
    tempDir := PromptContext_TempDir()
    failedCompact := 0
    failedStage := 0
    tempPrefix := StrLower(A_Temp "\prompt-context\")
    for e in entries {
        src := PromptData_ContextEntryPath(e)
        workSrc := src
        if PromptData_ContextEntryNeedsTransform(e) {
            outName := PromptData_UniqueCompactedName(src, usedNames)
            dst := tempDir "\" outName
            compact := (e.HasProp("compact") && e.compact) ? 1 : 0
            csvFrom := 0
            csvTo := 0
            if (PromptData_IsCsvPath(src)) {
                csvFrom := e.HasProp("csvKeepFrom") ? e.csvKeepFrom : 0
                csvTo := e.HasProp("csvKeepTo") ? e.csvKeepTo : 0
            }
            if PromptContext_RunCompact(src, dst, compact, csvFrom, csvTo) {
                workSrc := dst
                temps.Push(dst)
            } else {
                failedCompact += 1
                workSrc := src
            }
        }
        SplitPath workSrc, , , &workExt
        alreadyLocal := (InStr(StrLower(workSrc), tempPrefix) = 1)
        if (alreadyLocal && StrLower(workExt) != "ini") {
            paths.Push(workSrc)
            continue
        }
        staged := PromptContext_StageLocalCopy(workSrc, tempDir, usedNames)
        if (staged != "") {
            paths.Push(staged)
            if (staged != workSrc)
                temps.Push(staged)
        } else {
            failedStage += 1
            paths.Push(workSrc)
        }
    }
    if (failedCompact > 0)
        ShowCenteredOverlay_Utils("⚠ Compact failed for " . failedCompact . " file(s); attaching original", 2200,
            BANNER_ACCENT_ERROR)
    if (failedStage > 0)
        ShowCenteredOverlay_Utils("⚠ Local stage failed for " . failedStage . " file(s); attaching source", 2200,
            BANNER_ACCENT_ERROR)
    if (temps.Length > 0)
        PromptContext_ScheduleTempCleanup(temps)
    return paths
}

; After CF_HDROP paste: wait until Gemini/Enterprise upload UI settles before prompt body paste.
PromptContext_WaitForAttachUploadIdle(fileCount := 1) {
    if !InsertFiles_IsAiChatForeground()
        return
    minMs := (fileCount >= 3) ? 2000 : 1500
    uia := ""
    try uia := Gemini_GetUiaForActiveGeminiChrome()
    catch {
        uia := ""
    }
    if (!IsObject(uia)) {
        try {
            hwnd := WinGetID("A")
            if (hwnd)
                uia := UIA_Browser("ahk_id " hwnd)
        } catch {
            uia := ""
        }
    }
    if (IsObject(uia)) {
        try Gemini_WaitForUploadIdleWithRefocus(uia, 8000, minMs)
        catch {
            Sleep minMs
        }
    } else {
        Sleep minMs
    }
}

UtilitySelector_PastePromptToGemini(expansion, prompt := false, doAttach := true, doPasteBody := true, doAutoSend :=
    false) {
    companion := ResolveGlobalAICompanion()
    aiLabel := GetGlobalAIProviderLabel()
    HotstringGeminiBanner_Show("📤 " . aiLabel . ": inserting prompt...")
    try {
        if (companion = "enterprise") {
            GeminiEnterprise_OpenOrFocus()
            if (doAttach)
                UtilitySelector_AttachPromptContextFiles(prompt)
            if (doPasteBody) {
                InsertText(expansion)
                try ReplaceComposerWithStrippedReminders(expansion)
                catch {
                }
            }
            if (doAutoSend)
                Send "{Enter}"
        } else if (companion = "copilot") {
            CopilotWeb_OpenOrFocus()
            if (doAttach)
                UtilitySelector_AttachPromptContextFiles(prompt)
            if (doPasteBody) {
                InsertText(expansion)
                try ReplaceComposerWithStrippedReminders(expansion)
                catch {
                }
            }
            if (doAutoSend)
                Send "{Enter}"
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

            if (doAttach)
                UtilitySelector_AttachPromptContextFiles(prompt)
            if (doPasteBody) {
                InsertText(expansion)
                ScriptSoundPlay(A_ScriptDir . "\assets\sounds\gemini-focused.wav")
                try ReplaceComposerWithStrippedReminders(expansion)
                catch {
                }
            }
            if (doAutoSend)
                Send "{Enter}"
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
