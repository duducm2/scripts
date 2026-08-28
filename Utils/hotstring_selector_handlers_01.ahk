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

UtilitySelector_InsertPrompt(prompt, useGemini := false, appendClipboard := false) {
    global g_lastExpansion
    if (!IsObject(prompt))
        return
    body := PromptRender_Prepare(prompt)
    if (body = "")
        return
    mode := PromptData_PasteMode(prompt)
    doAttach := (mode = "default" || mode = "body_attach_clipboard" || mode = "attach_only")
    doPasteBody := (mode = "default" || mode = "body_only" || mode = "body_plus_clipboard" || mode =
        "body_attach_clipboard"
        || mode = "auto_send")
    doAppendClipboard := (mode = "body_plus_clipboard") || appendClipboard
    pasteChoice := ""
    if (doPasteBody) {
        CleanupHotstringSelector()
        pasteChoice := PromptPaste_ShowOptionsAndWait()
    }
    contextEntries := ""
    pickedCount := 0
    if (doAttach) {
        resolved := UtilitySelector_ResolveContextEntries(prompt)
        if (resolved = false)
            return
        contextEntries := resolved.entries
        pickedCount := resolved.pickedCount
    }
    PromptUsage_Log(prompt, useGemini ? "gemini" : "direct", pickedCount)
    clip := ""
    if (doAppendClipboard) {
        try clip := A_Clipboard
        catch {
        }
    }
    if (!doPasteBody)
        CleanupHotstringSelector()
    if (useGemini) {
        UtilitySelector_PastePromptToGemini(body, prompt, doAttach, doPasteBody, clip, contextEntries, pasteChoice)
        return
    }
    UtilitySelector_RestorePreviousHwnd()
    Sleep 150
    if (doAttach)
        UtilitySelector_AttachPromptContextFiles(prompt, contextEntries)
    if (doPasteBody) {
        onAfter := ""
        if (doAppendClipboard && clip != "") {
            clipCopy := clip
            onAfter := (*) => (g_lastExpansion := 0, InsertText(clipCopy))
        }
        global g_UtilitySelectorRestoreHwnd
        attachCount := (doAttach && contextEntries.Length) ? contextEntries.Length : 0
        companionId := (attachCount > 0 || InsertFiles_IsAiChatForeground()) ? ResolveGlobalAICompanion() : ""
        submitOpts := { hwnd: g_UtilitySelectorRestoreHwnd, companionId: companionId, attachCount: attachCount }
        PromptPaste_ApplyChoice(pasteChoice, body, onAfter, UtilitySelector_RestorePreviousHwnd, submitOpts)
    } else if (doAppendClipboard && clip != "") {
        g_lastExpansion := 0
        InsertText(clip)
    }
}

UtilitySelector_ResolveContextEntries(prompt) {
    if (!IsObject(prompt))
        return { entries: [], pickedCount: 0 }
    staticEntries := PromptData_ContextEntriesForCurrentEnv(prompt)
    pool := PromptContextPicker_BuildPool(prompt)
    if (pool.Length = 0)
        return { entries: staticEntries, pickedCount: 0 }
    picked := PromptContextPicker_ShowPool(pool)
    if (picked = false)
        return false
    return {
        entries: PromptContext_MergeEntries(staticEntries, picked),
        pickedCount: IsObject(picked) ? picked.Length : 0
    }
}

UtilitySelector_AttachPromptContextFiles(prompt, entries := "") {
    if (!IsObject(prompt))
        return
    if (entries = "")
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
    asTxt := PromptData_AttachAsTxt(prompt)
    attachPaths := PromptContext_ResolveAttachPaths(existing, asTxt)
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

; Basename for staged attach. asTxt (or .ini): force .txt for Gemini upload compatibility.
PromptContext_StagedAttachName(path, usedMap, asTxt := false) {
    SplitPath path, &name, , &ext, &nameNoExt
    if (nameNoExt = "")
        nameNoExt := (name != "") ? name : "file"
    forceTxt := asTxt || (StrLower(ext) = "ini")
    if (forceTxt)
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
        if (forceTxt || ext = "")
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
PromptContext_StageLocalCopy(src, tempDir, usedMap, asTxt := false) {
    if (src = "" || !Clipboard_PathIsExistingFile(src))
        return ""
    if (tempDir = "")
        return ""
    outName := PromptContext_StagedAttachName(src, usedMap, asTxt)
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

PromptContext_ResolveAttachPaths(entries, asTxt := false) {
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
        needsTxtRename := asTxt || (StrLower(workExt) = "ini")
        if (alreadyLocal && !needsTxtRename) {
            paths.Push(workSrc)
            continue
        }
        if (alreadyLocal && asTxt && StrLower(workExt) = "txt") {
            paths.Push(workSrc)
            continue
        }
        staged := PromptContext_StageLocalCopy(workSrc, tempDir, usedNames, asTxt)
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
; Self-contained (no Shift-keys helpers) so Utils #Warn stays clean when those symbols are absent.
PromptContext_IsUploading(uia) {
    if (!IsObject(uia))
        return false
    try {
        texts := uia.FindAll({ Type: 50020 }) ; Text
        for t in texts {
            name := ""
            try name := t.Name
            catch {
                continue
            }
            if (!name)
                continue
            low := StrLower(name)
            if (InStr(low, "open upload file menu"))
                continue
            if (InStr(low, "upload") || InStr(low, "sending") || InStr(low, "carreg") || InStr(low, "enviando"))
                return true
        }
    } catch {
    }
    return false
}

PromptContext_WaitForAttachUploadIdle(fileCount := 1) {
    if !InsertFiles_IsAiChatForeground()
        return
    minMs := (fileCount >= 3) ? 2000 : 1500
    ; Base 8s + 2s per file above 2, cap 25s (multi-file finance attach).
    extra := (fileCount > 2) ? ((fileCount - 2) * 2000) : 0
    timeoutMs := Min(8000 + extra, 25000)
    uia := ""
    try {
        hwnd := WinGetID("A")
        if (hwnd)
            uia := UIA_Browser("ahk_id " hwnd)
    } catch {
        uia := ""
    }
    if (!IsObject(uia)) {
        Sleep minMs
        return
    }
    tStart := A_TickCount
    sawUploading := false
    while ((A_TickCount - tStart) < timeoutMs) {
        up := PromptContext_IsUploading(uia)
        if (up)
            sawUploading := true
        if (!up && (sawUploading || (A_TickCount - tStart) >= minMs))
            return
        Sleep 150
    }
}

; After multi-file attach + body paste: wait until companion Send/Submit is enabled.
; Returns true if ready, false on timeout (caller may still attempt submit).
PromptContext_WaitForSendReady(hwnd, companionId := "", timeoutMs := 45000) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    companionId := StrLower(Trim(companionId))
    tStart := A_TickCount
    while ((A_TickCount - tStart) < timeoutMs) {
        if (!WinExist("ahk_id " hwnd))
            return false
        uia := 0
        try uia := UIA_Browser("ahk_id " hwnd)
        catch {
            uia := 0
        }
        if (IsObject(uia) && !PromptContext_IsUploading(uia)) {
            sendBtn := 0
            hasText := false
            try {
                if (companionId = "enterprise") {
                    sendBtn := GeminiEnterprise_FindSubmitButton(uia)
                    hasText := (GeminiEnterprise_ComposerGetTextViaUia(hwnd) != "")
                } else if (companionId = "copilot") {
                    sendBtn := CopilotWeb_FindSendButton(uia)
                    hasText := (CopilotWeb_ComposerGetText(hwnd) != "")
                } else {
                    sendBtn := Gemini_FindSendButton(uia)
                    hasText := (GeminiPromptFieldGetTextFromUia(uia) != "")
                }
            } catch {
                sendBtn := 0
                hasText := false
            }
            if (sendBtn && hasText) {
                enabled := false
                try enabled := !!sendBtn.GetPropertyValue(UIA.Property.IsEnabled)
                catch {
                    try enabled := !!sendBtn.IsEnabled
                    catch {
                        enabled := true ; control found; treat as ready if IsEnabled unavailable
                    }
                }
                if (enabled)
                    return true
            }
        }
        Sleep 200
    }
    return false
}

; [S] send: wait for upload idle + enabled Send per companion, then submit (mirrors D2C_FlowManager).
PromptPaste_SubmitWhenReady(hwnd := 0, companionId := "", attachCount := 0) {
    companionId := StrLower(Trim(companionId))
    if (companionId = "" && (attachCount > 0 || InsertFiles_IsAiChatForeground()))
        companionId := ResolveGlobalAICompanion()
    if (companionId != "" && (!hwnd || !WinExist("ahk_id " hwnd))) {
        if (companionId = "enterprise") {
            try hwnd := GetGeminiEnterpriseWindowHwnd()
            catch {
            }
        } else if (companionId = "copilot") {
            try hwnd := GetCopilotWebWindowHwnd()
            catch {
            }
        } else {
            try hwnd := FindGeminiChromeHwnd()
            catch {
            }
        }
    }
    if (!hwnd)
        hwnd := WinExist("A")
    if (!hwnd)
        return false

    if (attachCount > 0 || companionId != "") {
        ready := false
        try ready := PromptContext_WaitForSendReady(hwnd, companionId, 45000)
        catch {
        }
        if (!ready && attachCount > 0) {
            try ShowCenteredOverlay_Utils("⚠ Send not ready — submitting anyway", 2200, BANNER_ACCENT_ERROR)
            catch {
            }
        }
    }

    if (companionId = "enterprise") {
        Sleep 1000
        endTick := A_TickCount + 5000
        while (A_TickCount < endTick) {
            if (GeminiEnterprise_ComposerGetText(hwnd) != "")
                break
            Sleep 200
        }
        try {
            uia := UIA_Browser("ahk_id " hwnd)
            return GeminiEnterprise_TrySubmit(uia)
        } catch {
            Send "{Enter}"
            return true
        }
    } else if (companionId = "copilot") {
        Sleep 1000
        endTick := A_TickCount + 5000
        while (A_TickCount < endTick) {
            if (CopilotWeb_ComposerGetText(hwnd) != "")
                break
            Sleep 200
        }
        try {
            uia := UIA_Browser("ahk_id " hwnd)
            return CopilotWeb_TrySubmit(uia)
        } catch {
            Send "{Enter}"
            return true
        }
    } else if (companionId = "gemini") {
        return Gemini_WaitForPromptContentAndSubmit(hwnd)
    }
    if (attachCount > 0)
        Sleep 400
    Send "{Enter}"
    return true
}

UtilitySelector_RestoreConsumerGeminiFocus(*) {
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
}

UtilitySelector_PastePromptToGemini(expansion, prompt := false, doAttach := true, doPasteBody := true, appendClip := "",
    contextEntries := "", pasteChoice := "") {
    global g_lastExpansion
    companion := ResolveGlobalAICompanion()
    aiLabel := GetGlobalAIProviderLabel()
    HotstringGeminiBanner_Show("📤 " . aiLabel . ": inserting prompt...")
    restoreFocus := ""
    playGeminiChime := false
    companionHwnd := 0
    try {
        if (companion = "enterprise") {
            GeminiEnterprise_OpenOrFocus()
            restoreFocus := GeminiEnterprise_OpenOrFocus
            try companionHwnd := GetGeminiEnterpriseWindowHwnd()
            catch {
            }
        } else if (companion = "copilot") {
            CopilotWeb_OpenOrFocus()
            restoreFocus := CopilotWeb_OpenOrFocus
            try companionHwnd := GetCopilotWebWindowHwnd()
            catch {
            }
        } else {
            UtilitySelector_RestoreConsumerGeminiFocus()
            restoreFocus := UtilitySelector_RestoreConsumerGeminiFocus
            playGeminiChime := true
            try companionHwnd := FindGeminiChromeHwnd()
            catch {
            }
            if (!companionHwnd)
                companionHwnd := WinExist("A")
        }
        if (doAttach)
            UtilitySelector_AttachPromptContextFiles(prompt, contextEntries)
    } finally {
        HotstringGeminiBanner_Hide()
    }
    attachCount := IsObject(contextEntries) ? contextEntries.Length : 0
    submitOpts := { hwnd: companionHwnd, companionId: companion, attachCount: attachCount }
    if (doPasteBody) {
        onAfter := ""
        if (appendClip != "") {
            clipCopy := appendClip
            if (playGeminiChime) {
                onAfter := (*) => (ScriptSoundPlay(A_ScriptDir . "\assets\sounds\gemini-focused.wav"),
                g_lastExpansion := 0, InsertText(clipCopy))
            } else {
                onAfter := (*) => (g_lastExpansion := 0, InsertText(clipCopy))
            }
        } else if (playGeminiChime) {
            onAfter := (*) => ScriptSoundPlay(A_ScriptDir . "\assets\sounds\gemini-focused.wav")
        }
        PromptPaste_ApplyChoice(pasteChoice, expansion, onAfter, restoreFocus, submitOpts)
    } else if (appendClip != "") {
        g_lastExpansion := 0
        InsertText(appendClip)
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
