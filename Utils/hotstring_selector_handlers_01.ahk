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
    fp := StrLower(Trim(prompt.HasProp("filePath") ? prompt.filePath : ""))
    if (InStr(fp, "mnemonic-atoms-import")) {
        studyId := Trim(Palace_Setting("General", "LastStudyId", ""))
        pack := Palace_BuildPromptStudyPack(studyId)
        if (pack.Length)
            staticEntries := PromptContext_MergeEntries(staticEntries, pack)
    }
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
; Scoped to companion main pane (Gemini_GetSearchRoot) — avoid Shift-keys helpers for Utils #Warn.
; Prompt Manager [Y] auto-send: stable Send-enabled + upload-idle (efficiency canon §13 / stable polls).
PROMPT_PASTE_USE_STABLE_SEND_READY := true
PROMPT_PASTE_USE_CHIP_READY := true
PROMPT_PASTE_SEND_READY_STABLE_POLLS := 2
PROMPT_PASTE_SEND_READY_POLL_MS := 200
PROMPT_PASTE_SEND_MIN_NO_INDICATOR_MS := 600
; Total cap for Prompt Manager [Y] auto-send (wait + submit + confirm). Efficiency canon: bounded waits.
PROMPT_PASTE_AUTO_SEND_CAP_MS := 10000
PROMPT_PASTE_SEND_MIN_SUBMIT_MS := 4000

PromptContext_UploadSearchRoot(uia, companionId := "") {
    if (!IsObject(uia))
        return 0
    companionId := StrLower(Trim(companionId))
    try {
        if (companionId = "enterprise" || companionId = "copilot")
            return uia
        root := Gemini_GetSearchRoot(uia)
        if (root)
            return root
    } catch {
    }
    return uia
}

PromptContext_IsUploading(uia, companionId := "") {
    if (!IsObject(uia))
        return false
    root := PromptContext_UploadSearchRoot(uia, companionId)
    if (!IsObject(root))
        root := uia
    try {
        ; Text under search root only (not full Chrome tree). Trade-off: misses upload UI outside main pane.
        texts := root.FindAll({ Type: 50020 }) ; Text
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

PromptContext_SendButtonIsEnabled(sendBtn) {
    if (!sendBtn)
        return false
    try {
        return !!sendBtn.GetPropertyValue(UIA.Property.IsEnabled)
    } catch {
    }
    try {
        return !!sendBtn.IsEnabled
    } catch {
    }
    return true ; control found; treat as ready if IsEnabled unavailable
}

PromptContext_ProbeSendReady(hwnd, uia, companionId) {
    companionId := StrLower(Trim(companionId))
    uploadIdle := !PromptContext_IsUploading(uia, companionId)
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
        return { uploadIdle: uploadIdle, sendEnabled: false, hasText: false }
    }
    return {
        uploadIdle: uploadIdle,
        sendEnabled: PromptContext_SendButtonIsEnabled(sendBtn),
        hasText: hasText
    }
}

; Secondary signal (orthogonal to Send-enabled + upload text): composer chips and no ProgressBar.
; Scoped FindAll under search root; stop once chipNeed is met. Trade-off: other "remove" buttons in pane can inflate count.
PromptContext_CountFileChips(uia, companionId := "", chipNeed := 0) {
    n := 0
    root := PromptContext_UploadSearchRoot(uia, companionId)
    if (!IsObject(root))
        return 0
    try {
        buttons := root.FindAll({ Type: 50000 }) ; Button
        for btn in buttons {
            name := ""
            try name := btn.Name
            catch {
                continue
            }
            if (!name)
                continue
            low := StrLower(name)
            if (InStr(low, "open upload file menu"))
                continue
            if (InStr(low, "remove") || InStr(low, "remover") || InStr(low, "excluir")
            || InStr(low, "delete file") || InStr(low, "close file") || InStr(low, "fechar arquivo")) {
                n += 1
                if (chipNeed > 0 && n >= chipNeed)
                    return n
            }
        }
    } catch {
    }
    return n
}

PromptContext_HasProgressBar(uia, companionId := "") {
    root := PromptContext_UploadSearchRoot(uia, companionId)
    if (!IsObject(root))
        return false
    try {
        el := root.FindFirst({ Type: 50012 }) ; ProgressBar
        return !!el
    } catch {
    }
    return false
}

; Gate B: chips (when attaching) + no ProgressBar + composer text. Does not use Send.IsEnabled or upload labels.
PromptContext_ProbeChipReady(hwnd, uia, companionId, attachCount := 0) {
    companionId := StrLower(Trim(companionId))
    hasText := false
    try {
        if (companionId = "enterprise")
            hasText := (GeminiEnterprise_ComposerGetTextViaUia(hwnd) != "")
        else if (companionId = "copilot")
            hasText := (CopilotWeb_ComposerGetText(hwnd) != "")
        else
            hasText := (GeminiPromptFieldGetTextFromUia(uia) != "")
    } catch {
        hasText := false
    }
    if (!hasText)
        return false
    if (PromptContext_HasProgressBar(uia, companionId))
        return false
    if (attachCount > 0)
        return PromptContext_CountFileChips(uia, companionId, attachCount) >= attachCount
    return true
}

PromptContext_WaitForAttachUploadIdle(fileCount := 1) {
    if !InsertFiles_IsAiChatForeground()
        return
    minMs := (fileCount >= 3) ? 2000 : 1500
    ; Base 8s + 2s per file above 2, cap 25s (multi-file finance attach).
    extra := (fileCount > 2) ? ((fileCount - 2) * 2000) : 0
    timeoutMs := Min(8000 + extra, 25000)
    companionId := ""
    try companionId := ResolveGlobalAICompanion()
    catch {
    }
    uia := ""
    try {
        hwnd := WinGetID("A")
        if (hwnd)
            uia := PromptPaste_UiaForCompanion(hwnd, companionId)
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
        up := PromptContext_IsUploading(uia, companionId)
        if (up)
            sawUploading := true
        if (!up && (sawUploading || (A_TickCount - tStart) >= minMs))
            return
        Sleep 150
    }
}

; After multi-file attach + body paste: wait until companion Send/Submit is enabled.
; Returns true if ready, false on timeout (caller may still attempt submit).
; Two independent streaks in one poll loop (AHK has no worker threads); first to N wins:
;   A — upload-idle + Send enabled + text (PROMPT_PASTE_USE_STABLE_SEND_READY)
;   B — chips + no ProgressBar + text (PROMPT_PASTE_USE_CHIP_READY)
PromptContext_WaitForSendReady(hwnd, companionId := "", timeoutMs := 45000, attachCount := 0, updateBanner := false) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    companionId := StrLower(Trim(companionId))
    if (!PROMPT_PASTE_USE_STABLE_SEND_READY && !PROMPT_PASTE_USE_CHIP_READY)
        return PromptContext_WaitForSendReadyLegacy(hwnd, companionId, timeoutMs)

    pollMs := Max(50, PROMPT_PASTE_SEND_READY_POLL_MS)
    needStable := Max(1, PROMPT_PASTE_SEND_READY_STABLE_POLLS)
    minNoInd := (attachCount > 0) ? Max(0, PROMPT_PASTE_SEND_MIN_NO_INDICATOR_MS) : 0
    tStart := A_TickCount
    sawUploading := false
    stableA := 0
    stableB := 0
    bannerPhase := ""
    uia := 0
    pollIndex := 0

    while ((A_TickCount - tStart) < timeoutMs) {
        if (!WinExist("ahk_id " hwnd))
            return false
        pollIndex += 1
        ; Refresh UIA periodically; stale COM after long uploads can miss Send enablement.
        if (!IsObject(uia) || Mod(pollIndex, 5) = 1) {
            try uia := PromptPaste_UiaForCompanion(hwnd, companionId)
            catch {
                uia := 0
            }
        }
        if (!IsObject(uia)) {
            stableA := 0
            stableB := 0
            Sleep pollMs
            continue
        }

        probe := PromptContext_ProbeSendReady(hwnd, uia, companionId)
        if (!probe.uploadIdle)
            sawUploading := true

        if (updateBanner) {
            phase := !probe.uploadIdle ? "uploads" : "send"
            if (phase != bannerPhase) {
                bannerPhase := phase
                try {
                    if (phase = "uploads")
                        StandardLoadingBar_Update("⏳ Waiting for uploads…", BANNER_ACCENT_INTERMEDIATE)
                    else
                        StandardLoadingBar_Update("⏳ Waiting for Send…", BANNER_ACCENT_INTERMEDIATE)
                } catch {
                }
            }
        }

        if (PROMPT_PASTE_USE_STABLE_SEND_READY) {
            idleOk := probe.uploadIdle && (sawUploading || (A_TickCount - tStart) >= minNoInd)
            if (idleOk && probe.sendEnabled && probe.hasText) {
                stableA += 1
                if (stableA >= needStable)
                    return true
            } else {
                stableA := 0
            }
        }

        if (PROMPT_PASTE_USE_CHIP_READY) {
            chipOk := false
            try chipOk := PromptContext_ProbeChipReady(hwnd, uia, companionId, attachCount)
            catch {
                chipOk := false
            }
            if (chipOk) {
                stableB += 1
                if (stableB >= needStable)
                    return true
            } else {
                stableB := 0
            }
        }

        Sleep pollMs
    }
    return false
}

PromptContext_WaitForSendReadyLegacy(hwnd, companionId := "", timeoutMs := 45000) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    companionId := StrLower(Trim(companionId))
    tStart := A_TickCount
    while ((A_TickCount - tStart) < timeoutMs) {
        if (!WinExist("ahk_id " hwnd))
            return false
        uia := 0
        try uia := PromptPaste_UiaForCompanion(hwnd, companionId)
        catch {
            uia := 0
        }
        if (IsObject(uia)) {
            probe := PromptContext_ProbeSendReady(hwnd, uia, companionId)
            if (probe.uploadIdle && probe.sendEnabled && probe.hasText)
                return true
        }
        Sleep 200
    }
    return false
}

PromptPaste_SendRemainingMs(tDeadline) {
    if (!tDeadline)
        return PROMPT_PASTE_AUTO_SEND_CAP_MS
    return Max(0, tDeadline - A_TickCount)
}

; Reserve time at end of cap for focus + Enter/submit + generation confirm (D2C-style).
PromptPaste_SendWaitBudget(tDeadline) {
    return Max(0, PromptPaste_SendRemainingMs(tDeadline) - PROMPT_PASTE_SEND_MIN_SUBMIT_MS)
}

PromptPaste_FocusCompanionComposer(hwnd, companionId) {
    companionId := StrLower(Trim(companionId))
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    if (!WinActive("ahk_id " hwnd)) {
        WinActivate("ahk_id " hwnd)
        if (!WinWaitActive("ahk_id " hwnd, , 2))
            return false
    }
    try {
        if (companionId = "enterprise") {
            root := GeminiEnterprise_ReadRootFromHwnd(hwnd)
            return IsObject(root) && GeminiEnterprise_FocusComposer(root, false)
        }
        if (companionId = "copilot")
            return CopilotWeb_FocusComposerForHwnd(hwnd, false)
        if (companionId = "gemini") {
            uia := UIA_Browser("ahk_id " hwnd)
            return !!Gemini_FocusPromptWithChime(uia, { playChime: false, useAnchorFallback: true })
        }
    } catch {
    }
    return WinActive("ahk_id " hwnd)
}

PromptPaste_UiaForCompanion(hwnd, companionId) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return 0
    companionId := StrLower(Trim(companionId))
    try {
        if (companionId = "enterprise") {
            root := GeminiEnterprise_ReadRootFromHwnd(hwnd)
            if (root)
                return root
        } else if (companionId = "copilot") {
            root := CopilotWeb_ReadRootFromHwnd(hwnd)
            if (root)
                return root
        }
        return UIA_Browser("ahk_id " hwnd)
    } catch {
        return 0
    }
}

PromptPaste_CompanionIsGenerating(hwnd, companionId) {
    companionId := StrLower(Trim(companionId))
    if (!hwnd || !WinExist("ahk_id " hwnd) || companionId = "")
        return false
    uia := PromptPaste_UiaForCompanion(hwnd, companionId)
    if (!IsObject(uia))
        return false
    try {
        if (companionId = "enterprise")
            return !!GeminiEnterprise_FindStopButton(uia)
        if (companionId = "copilot")
            return !!CopilotWeb_FindStopGenerating(uia)
        if (companionId = "gemini")
            return Gemini_HasGeneratingStopButtonForUia(uia)
    } catch {
    }
    return false
}

PromptPaste_WaitForGenerationStarted(hwnd, companionId, timeoutMs := 5000) {
    tStart := A_TickCount
    while ((A_TickCount - tStart) < timeoutMs) {
        if (PromptPaste_CompanionIsGenerating(hwnd, companionId))
            return true
        Sleep 200
    }
    return false
}

PromptPaste_SubmitCompanion(hwnd, companionId, tDeadline := 0) {
    global g_GeminiDelayedSubmit_WaitContentMaxMs
    companionId := StrLower(Trim(companionId))
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    if (PromptPaste_CompanionIsGenerating(hwnd, companionId))
        return true
    if (PromptPaste_SendRemainingMs(tDeadline) <= 0)
        return false
    prevHwnd := WinExist("A")
    result := false
    try {
        if (companionId != "") {
            if (!PromptPaste_FocusCompanionComposer(hwnd, companionId))
                return false
            Sleep 200
        } else if (!WinActive("ahk_id " hwnd)) {
            WinActivate("ahk_id " hwnd)
            WinWaitActive("ahk_id " hwnd, , 1)
        }
        if (companionId = "enterprise") {
            try {
                uia := UIA_Browser("ahk_id " hwnd)
                result := GeminiEnterprise_TrySubmit(uia)
            } catch {
                result := false
            }
        } else if (companionId = "copilot") {
            try {
                uia := UIA_Browser("ahk_id " hwnd)
                result := CopilotWeb_TrySubmit(uia)
            } catch {
                result := false
            }
        } else if (companionId = "gemini") {
            uia := 0
            try uia := UIA_Browser("ahk_id " hwnd)
            catch {
                uia := 0
            }
            if (IsObject(uia)) {
                contentMs := Min(g_GeminiDelayedSubmit_WaitContentMaxMs, PromptPaste_SendRemainingMs(tDeadline))
                if (contentMs > 0 && Gemini_WaitForPromptContent(uia, contentMs))
                    result := Gemini_TrySubmit(hwnd, uia)
            }
        } else {
            Send "{Enter}"
            result := true
        }
    } finally {
        if (prevHwnd && prevHwnd != hwnd && WinExist("ahk_id " prevHwnd))
            WinActivate("ahk_id " prevHwnd)
    }
    return result
}

; [Y] send: non-blocking loading bar, wait for upload idle, UIA submit, confirm generation started.
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

    showBar := (attachCount > 0 || companionId != "")
    ok := false
    tDeadline := A_TickCount + PROMPT_PASTE_AUTO_SEND_CAP_MS
    try {
        if (showBar) {
            try StandardLoadingBar_Show("⏳ Waiting to send…", BANNER_ACCENT_INTERMEDIATE, {
                passive: false,
                centerOnHwnd: hwnd,
                trackActiveMonitor: true
            })
            catch {
            }
        }

        if (attachCount > 0 || companionId != "") {
            try StandardLoadingBar_Update(
                (attachCount > 0) ? "⏳ Waiting for uploads…" : "⏳ Waiting for Send…",
                BANNER_ACCENT_INTERMEDIATE)
            catch {
            }
            ready := false
            waitMs := PromptPaste_SendWaitBudget(tDeadline)
            if (waitMs > 0) {
                try ready := PromptContext_WaitForSendReady(hwnd, companionId, waitMs, attachCount, true)
                catch {
                }
            }
            if (!ready && attachCount > 0 && PromptPaste_SendRemainingMs(tDeadline) > PROMPT_PASTE_SEND_MIN_SUBMIT_MS) {
                try ShowCenteredOverlay_Utils("⚠ Send not ready — submitting anyway", 2200, BANNER_ACCENT_ERROR)
                catch {
                }
            }
        }

        if (PromptPaste_SendRemainingMs(tDeadline) <= PROMPT_PASTE_SEND_MIN_SUBMIT_MS / 2) {
            ok := false
        } else if (companionId != "" && PromptPaste_CompanionIsGenerating(hwnd, companionId)) {
            ok := true
        } else {
            if (showBar) {
                try StandardLoadingBar_Update("⏳ Sending…", BANNER_ACCENT_INTERMEDIATE)
                catch {
                }
            }
            submitted := PromptPaste_SubmitCompanion(hwnd, companionId, tDeadline)
            if (companionId != "") {
                if (showBar) {
                    try StandardLoadingBar_Update("⏳ Confirming…", BANNER_ACCENT_INTERMEDIATE)
                    catch {
                    }
                }
                confirmMs := PromptPaste_SendRemainingMs(tDeadline)
                ok := (confirmMs > 0 && PromptPaste_WaitForGenerationStarted(hwnd, companionId, confirmMs))
                if (!ok && submitted)
                    ok := PromptPaste_CompanionIsGenerating(hwnd, companionId)
            } else {
                ok := submitted
            }
        }
    } finally {
        if (showBar) {
            try StandardLoadingBar_Hide(0)
            catch {
            }
        }
    }

    if (companionId != "" || attachCount > 0) {
        if (ok) {
            try ShowCenteredOverlay_Utils("✅ Sent — AI is working", 1800, BANNER_ACCENT_SUCCESS)
            catch {
            }
        } else {
            msg := (A_TickCount >= tDeadline) ? "⚠ Send timed out (10s)" : "⚠ Send may not have started"
            try ShowCenteredOverlay_Utils(msg, 2200, BANNER_ACCENT_ERROR)
            catch {
            }
        }
    }
    return ok
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
