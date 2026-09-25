; =============================================================================
; Utils module: pack_pipeline.ahk
; Auto bridge: pack prompt sent → wait generation → extract (code then message)
; → validate → write canonical Desktop file → domain import (confirm GUI kept)
; → on structure fail: AI fix paste+submit and retry.
; Owner: AppLaunchers.ahk only. Coordinates with import_watcher busy/seen.
; Agent docs: docs/prompt-data-output-and-finance-packs.md
; =============================================================================

global g_PackPipeline := false
global g_PackPipelineMonitorTimer := ""
global g_PackPipelineMonitorRetry := 0
global g_PackPipelineMonitorMaxRetries := 600   ; ~5 min at 500ms
global g_PackPipelineMonitorButtonSeen := false
global g_PackPipelineMaxAttempts := 2

WM_COPY_LAST_GEMINI := 0x8001
WM_COPY_LAST_COPILOT := 0x8005
WM_COPY_LAST_GEMINI_CODE := 0x8007
WM_COPY_LAST_COPILOT_CODE := 0x8008

PackPipeline_IsOwnerProcess() {
    return A_ScriptName = "AppLaunchers.ahk"
}

; Catalog entries: keyed by Utility prompt char and D2C presetMode aliases.
PackPipeline_Catalog() {
    return Map(
        "d", Map("canonical", "FINANCE_DAILY.txt", "label", "Finance daily",
            "run", Finance_ImportDaily, "fixKind", "finance_daily"),
        "m", Map("canonical", "FINANCE_MONTHLY.txt", "label", "Finance monthly",
            "run", Finance_ImportMonthly, "fixKind", "finance_monthly"),
        "4", Map("canonical", "PALACE_PACK.txt", "label", "Palace pack",
            "run", Palace_ImportMnemonicsFromDesktop, "fixKind", "palace"),
        "a", Map("canonical", "PALACE_PACK.txt", "label", "Palace pack",
            "run", Palace_ImportMnemonicsFromDesktop, "fixKind", "palace"),
        "n", Map("canonical", "PLAN_PACK.txt", "label", "Study plan pack",
            "run", Palace_ImportPlanPackFromDesktop, "fixKind", "plan"),
        "k", Map("canonical", "TASK_PACK.txt", "label", "Task pack",
            "run", Task_ImportPackFromDesktop, "fixKind", "task"),
        "finance_daily", Map("canonical", "FINANCE_DAILY.txt", "label", "Finance daily",
            "run", Finance_ImportDaily, "fixKind", "finance_daily"),
        "task_pack", Map("canonical", "TASK_PACK.txt", "label", "Task pack",
            "run", Task_ImportPackFromDesktop, "fixKind", "task")
    )
}

PackPipeline_Lookup(key) {
    key := StrLower(Trim(key))
    if (key = "")
        return false
    cat := PackPipeline_Catalog()
    if (!cat.Has(key))
        return false
    return cat[key]
}

PackPipeline_Reset() {
    global g_PackPipeline, g_PackPipelineMonitorTimer, g_PackPipelineMonitorRetry,
        g_PackPipelineMonitorButtonSeen
    PackPipeline_StopMonitor()
    g_PackPipeline := false
    g_PackPipelineMonitorRetry := 0
    g_PackPipelineMonitorButtonSeen := false
}

PackPipeline_StopMonitor() {
    global g_PackPipelineMonitorTimer
    if (IsObject(g_PackPipelineMonitorTimer)) {
        try SetTimer(g_PackPipelineMonitorTimer, 0)
        catch {
        }
    }
    try SetTimer(PackPipeline_MonitorTick, 0)
    catch {
    }
    g_PackPipelineMonitorTimer := ""
}

PackPipeline_IsActive() {
    global g_PackPipeline
    return IsObject(g_PackPipeline) && g_PackPipeline.HasProp("active") && g_PackPipeline.active
}

; Capture foreground hwnd when it is not the companion (call before intentional focus steal).
PackPipeline_CaptureUserHwnd() {
    global g_PackPipeline
    if (!IsObject(g_PackPipeline))
        return 0
    companionHwnd := g_PackPipeline.HasProp("hwnd") ? g_PackPipeline.hwnd : 0
    try {
        fg := WinGetID("A")
        if (fg && (!companionHwnd || fg != companionHwnd)) {
            g_PackPipeline.userHwnd := fg
            return fg
        }
    } catch {
    }
    return g_PackPipeline.HasProp("userHwnd") ? g_PackPipeline.userHwnd : 0
}

; Restore the user's window after brief companion activate (extract / fix-submit).
PackPipeline_RestoreUserHwnd() {
    global g_PackPipeline
    if (!IsObject(g_PackPipeline) || !g_PackPipeline.HasProp("userHwnd"))
        return false
    userHwnd := g_PackPipeline.userHwnd
    if (!userHwnd || !WinExist("ahk_id " userHwnd))
        return false
    companionHwnd := g_PackPipeline.HasProp("hwnd") ? g_PackPipeline.hwnd : 0
    if (companionHwnd && userHwnd = companionHwnd)
        return false
    try {
        if (!WinActive("ahk_id " userHwnd)) {
            WinActivate("ahk_id " userHwnd)
            WinWaitActive("ahk_id " userHwnd, , 1)
        }
        return !!WinActive("ahk_id " userHwnd)
    } catch {
    }
    return false
}

; Information Only: user may leave the companion — pack wait is hwnd-only in background.
PackPipeline_NotifyUserFree() {
    global g_PackPipeline
    if (!PackPipeline_IsActive())
        return
    label := g_PackPipeline.HasProp("label") ? g_PackPipeline.label : "pack"
    msg := "ℹ Free to work — watching " . label . " in background"
    try ShowCenteredOverlay_Utils(msg, 4500, BANNER_ACCENT_INFO)
    catch {
    }
}

; Brief activate companion for copy/paste; captures userHwnd first.
PackPipeline_ActivateCompanionBriefly(hwnd) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    PackPipeline_CaptureUserHwnd()
    try {
        if (!WinActive("ahk_id " hwnd)) {
            WinActivate("ahk_id " hwnd)
            if (!WinWaitActive("ahk_id " hwnd, , 1.5))
                return false
        }
        return true
    } catch {
    }
    return false
}

; Arm from Utility prompt object after Gemini send.
; userHwnd: origin window before companion focus (dictation-style restore target).
PackPipeline_ArmFromPrompt(prompt, companionId := "", hwnd := 0, userHwnd := 0) {
    if (!PackPipeline_IsOwnerProcess())
        return false
    if (!IsObject(prompt))
        return false
    ch := StrLower(Trim(prompt.HasProp("char") ? prompt.char : ""))
    item := PackPipeline_Lookup(ch)
    if (!IsObject(item))
        return false
    return PackPipeline_Arm(ch, item, companionId, hwnd, userHwnd)
}

; True when prompt char is a pack-pipeline catalog entry (Finance/Palace/Plan/Task).
PackPipeline_IsCatalogPrompt(prompt) {
    if (!IsObject(prompt))
        return false
    ch := StrLower(Trim(prompt.HasProp("char") ? prompt.char : ""))
    return IsObject(PackPipeline_Lookup(ch))
}

; Arm after Prompt Manager / companion send when pasteChoice is "send" and prompt is catalog.
PackPipeline_MaybeArmAfterSend(prompt, pasteChoice, companionId := "", hwnd := 0, userHwnd := 0) {
    if (pasteChoice != "send")
        return false
    if (!PackPipeline_IsCatalogPrompt(prompt))
        return false
    return PackPipeline_ArmFromPrompt(prompt, companionId, hwnd, userHwnd)
}

; Arm from D2C presetMode ("finance_daily" | "task_pack").
PackPipeline_ArmFromPreset(presetMode, companionId := "", hwnd := 0, userHwnd := 0) {
    if (!PackPipeline_IsOwnerProcess())
        return false
    item := PackPipeline_Lookup(presetMode)
    if (!IsObject(item))
        return false
    return PackPipeline_Arm(presetMode, item, companionId, hwnd, userHwnd)
}

; Override restore target after arm (e.g. origin captured before companion focus).
PackPipeline_SetUserHwnd(userHwnd) {
    global g_PackPipeline
    if (!IsObject(g_PackPipeline))
        return false
    if (!userHwnd || !WinExist("ahk_id " userHwnd))
        return false
    companionHwnd := g_PackPipeline.HasProp("hwnd") ? g_PackPipeline.hwnd : 0
    if (companionHwnd && userHwnd = companionHwnd)
        return false
    g_PackPipeline.userHwnd := userHwnd
    return true
}

PackPipeline_Arm(key, item, companionId := "", hwnd := 0, userHwnd := 0) {
    global g_PackPipeline, g_PackPipelineMaxAttempts, g_PackPipelineMonitorRetry,
        g_PackPipelineMonitorButtonSeen
    if (!IsObject(item))
        return false
    if (companionId = "") {
        try companionId := ResolveGlobalAICompanion()
        catch {
            companionId := "gemini"
        }
    }
    if (!hwnd) {
        try {
            switch companionId {
                case "enterprise":
                    hwnd := GetGeminiEnterpriseWindowHwnd()
                case "copilot":
                    hwnd := GetCopilotWebWindowHwnd()
                default:
                    hwnd := FindGeminiChromeHwnd()
            }
        } catch {
            hwnd := 0
        }
    }
    PackPipeline_StopMonitor()
    ; Prefer caller-supplied origin (pre-companion); else foreground if not companion.
    if (!userHwnd || (hwnd && userHwnd = hwnd) || !WinExist("ahk_id " userHwnd)) {
        userHwnd := 0
        try {
            fg := WinGetID("A")
            if (fg && (!hwnd || fg != hwnd))
                userHwnd := fg
        } catch {
        }
    }
    g_PackPipeline := {
        active: true,
        key: key,
        companionId: companionId,
        hwnd: hwnd,
        userHwnd: userHwnd,
        canonical: item["canonical"],
        label: item["label"],
        run: item["run"],
        fixKind: item["fixKind"],
        attempt: 0,
        maxAttempts: g_PackPipelineMaxAttempts
    }
    g_PackPipelineMonitorRetry := 0
    g_PackPipelineMonitorButtonSeen := false
    PackPipeline_StartMonitor()
    PackPipeline_NotifyUserFree()
    return true
}

PackPipeline_StartMonitor() {
    global g_PackPipelineMonitorTimer, g_PackPipelineMonitorRetry, g_PackPipelineMonitorButtonSeen
    PackPipeline_StopMonitor()
    g_PackPipelineMonitorRetry := 0
    g_PackPipelineMonitorButtonSeen := false
    g_PackPipelineMonitorTimer := PackPipeline_MonitorTick
    SetTimer(PackPipeline_MonitorTick, 500)
}

; Hwnd-scoped Stop-button poll only — never WinActivate, never bare UIA_Browser().
PackPipeline_CompanionIsGenerating(hwnd, companionId) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    companionId := StrLower(Trim(companionId))
    try {
        if (companionId = "copilot") {
            root := CopilotWeb_ReadRootFromHwnd(hwnd)
            return !!(root && CopilotWeb_FindStopGenerating(root))
        }
        if (companionId = "enterprise") {
            root := GeminiEnterprise_ReadRootFromHwnd(hwnd)
            return !!(root && GeminiEnterprise_FindStopButton(root))
        }
        ; Consumer Gemini (default): hwnd-bound UIA only.
        try {
            uia := UIA_Browser("ahk_id " hwnd)
            if (IsObject(uia) && Gemini_HasGeneratingStopButtonForUia(uia))
                return true
        } catch {
        }
        try {
            root := UIA.ElementFromHandle(hwnd)
            if (IsObject(root) && Gemini_HasGeneratingStopButtonForUia(root))
                return true
            for n in ["Stop streaming", "Interromper transmissão", "Stop response"] {
                try {
                    if (root.FindElement({ Name: n, Type: "Button" }))
                        return true
                } catch {
                }
            }
        } catch {
        }
    } catch {
    }
    return false
}

PackPipeline_MonitorTick(*) {
    global g_PackPipeline, g_PackPipelineMonitorRetry, g_PackPipelineMonitorMaxRetries,
        g_PackPipelineMonitorButtonSeen
    if (!PackPipeline_IsActive()) {
        PackPipeline_StopMonitor()
        return
    }
    g_PackPipelineMonitorRetry += 1
    if (g_PackPipelineMonitorRetry > g_PackPipelineMonitorMaxRetries) {
        PackPipeline_Fail("Timed out waiting for " . g_PackPipeline.label . " response")
        return
    }
    hwnd := g_PackPipeline.hwnd
    companionId := g_PackPipeline.companionId
    if (!hwnd || !WinExist("ahk_id " hwnd)) {
        try {
            switch companionId {
                case "enterprise":
                    hwnd := GetGeminiEnterpriseWindowHwnd()
                case "copilot":
                    hwnd := GetCopilotWebWindowHwnd()
                default:
                    hwnd := FindGeminiChromeHwnd()
            }
            g_PackPipeline.hwnd := hwnd
        } catch {
        }
        if (!hwnd)
            return
    }
    ; Passive status every ~15s — no focus steal; remind user they can keep working.
    if (Mod(g_PackPipelineMonitorRetry, 30) = 0) {
        PackPipeline_NotifyUserFree()
    }
    generating := PackPipeline_CompanionIsGenerating(hwnd, companionId)
    if (generating) {
        g_PackPipelineMonitorButtonSeen := true
        return
    }
    if (!g_PackPipelineMonitorButtonSeen) {
        ; Not started yet — keep waiting for Stop to appear.
        if (g_PackPipelineMonitorRetry < 20)
            return
        ; After ~10s with no Stop, assume already complete (fast response / missed start).
        g_PackPipelineMonitorButtonSeen := true
    }
    ; Verify Stop stayed gone.
    PackPipeline_StopMonitor()
    isTrulyGone := true
    loop 4 {
        Sleep 200
        if (PackPipeline_CompanionIsGenerating(g_PackPipeline.hwnd, companionId)) {
            isTrulyGone := false
            break
        }
    }
    if (!isTrulyGone) {
        PackPipeline_StartMonitor()
        return
    }
    try ScriptSoundPlay(A_ScriptDir . "\assets\sounds\gemini-completion.wav")
    catch {
    }
    SetTimer(PackPipeline_OnGenerationComplete, -1)
}

PackPipeline_OnGenerationComplete(*) {
    if (!PackPipeline_IsActive())
        return
    try StandardLoadingBar_Show("⏳ Extracting pack…", BANNER_ACCENT_INTERMEDIATE, { passive: false })
    catch {
    }
    text := PackPipeline_Extract()
    try StandardLoadingBar_Hide(0)
    catch {
    }
    if (Trim(text) = "") {
        PackPipeline_HandleInvalid("Empty extraction (no code block or message)")
        return
    }
    if (!PackPipeline_ValidateStructure(text)) {
        PackPipeline_HandleInvalid("Pack structure invalid (missing PREVIEW/FILE markers or CSV header)")
        return
    }
    PackPipeline_WriteAndImport(text)
}

; --- Extraction (code primary, message fallback) ---

PackPipeline_Extract() {
    global g_PackPipeline
    if (!PackPipeline_IsActive())
        return ""
    hwnd := g_PackPipeline.hwnd
    companionId := g_PackPipeline.companionId
    text := PackPipeline_CopyCode(companionId, hwnd)
    if (Trim(text) != "" && StrLen(Trim(text)) >= 20)
        return text
    text := PackPipeline_CopyMessage(companionId, hwnd)
    return text
}

PackPipeline_CopyResultPath() {
    return A_ScriptDir "\.cursor\gemini_copy_result.txt"
}

PackPipeline_IpcCopy(wm, hwnd) {
    ; Capture before Gemini IPC may WinActivate the companion.
    PackPipeline_CaptureUserHwnd()
    targetHwnd := 0
    try targetHwnd := GetGeminiScriptMsgTargetHwnd()
    catch {
        targetHwnd := 0
    }
    if (!targetHwnd)
        return ""
    copyResultPath := PackPipeline_CopyResultPath()
    clipBefore := A_Clipboard
    seqBefore := 0
    try seqBefore := Clipboard_GetSequenceNumber()
    catch {
    }
    try {
        if (FileExist(copyResultPath))
            FileDelete(copyResultPath)
        FileAppend("0", copyResultPath)
    } catch {
    }
    prevDH := A_DetectHiddenWindows
    DetectHiddenWindows true
    sendOk := false
    try {
        SendMessage(wm, 0, hwnd, , "ahk_id " targetHwnd, , , , 20000)
        sendOk := true
    } catch {
    } finally {
        DetectHiddenWindows prevDH
        ; Bridge may have activated Chrome — restore user focus immediately.
        PackPipeline_RestoreUserHwnd()
    }
    if (!sendOk)
        return ""
    changed := true
    try changed := Clipboard_WaitForSequenceChange(seqBefore, 2000, 850)
    catch {
        Sleep 400
    }
    resultOk := false
    try resultOk := FileExist(copyResultPath) && Trim(FileRead(copyResultPath)) = "1"
    catch {
        resultOk := false
    }
    clipNow := Trim(A_Clipboard)
    if (!resultOk || clipNow = "" || StrLen(clipNow) < 10)
        return ""
    if (clipNow = Trim(clipBefore) && !changed)
        return ""
    return clipNow
}

PackPipeline_CopyCode(companionId, hwnd) {
    ; Activate briefly so alreadyActive:true binds UIA to companion (not user's fg window).
    PackPipeline_ActivateCompanionBriefly(hwnd)
    opts := { restoreWindow: false, playChimeAndNotify: false, alreadyActive: true }
    text := ""
    try {
        if (companionId = "copilot") {
            ok := false
            try ok := CopilotWeb_CopyLastCodeSnippetToClipboard(opts, hwnd)
            catch {
                ok := false
            }
            text := ok ? Trim(A_Clipboard) : ""
        } else if (companionId = "enterprise") {
            ok := false
            try ok := GeminiEnterprise_CopyLastCodeSnippetToClipboard(opts, hwnd)
            catch {
                ok := false
            }
            text := ok ? Trim(A_Clipboard) : ""
        } else {
            ; Consumer Gemini via IPC (code lives in Gemini.ahk process).
            ; IpcCopy also restores after SendMessage; finally covers early IPC failures.
            text := PackPipeline_IpcCopy(WM_COPY_LAST_GEMINI_CODE, hwnd)
        }
    } finally {
        PackPipeline_RestoreUserHwnd()
    }
    return text
}

PackPipeline_CopyMessage(companionId, hwnd) {
    PackPipeline_ActivateCompanionBriefly(hwnd)
    opts := { restoreWindow: false, playChimeAndNotify: false, alreadyActive: true }
    text := ""
    try {
        if (companionId = "copilot") {
            ok := false
            try ok := CopilotWeb_CopyLastMessageWithRetry(opts, hwnd)
            catch {
                try ok := CopilotWeb_CopyLastMessageToClipboard(opts, hwnd)
                catch {
                    ok := false
                }
            }
            text := ok ? Trim(A_Clipboard) : ""
        } else if (companionId = "enterprise") {
            ok := false
            try ok := GeminiEnterprise_CopyLastMessageWithRetry(opts, hwnd)
            catch {
                try ok := GeminiEnterprise_CopyLastMessageToClipboard(opts, hwnd)
                catch {
                    ok := false
                }
            }
            text := ok ? Trim(A_Clipboard) : ""
        } else {
            text := PackPipeline_IpcCopy(WM_COPY_LAST_GEMINI, hwnd)
        }
    } finally {
        PackPipeline_RestoreUserHwnd()
    }
    return text
}

; --- Validation ---

PackPipeline_ValidateStructure(text) {
    t := Trim(text)
    if (t = "" || StrLen(t) < 20)
        return false
    lower := StrLower(t)
    if (InStr(lower, "===preview===") || InStr(lower, "---preview---"))
        return true
    if (InStr(lower, "===file:") || InStr(lower, "---file:"))
        return true
    if (InStr(lower, "description,amount") || InStr(lower, "entity_type,entity_id"))
        return true
    if (InStr(lower, "palace_number") || InStr(lower, "beast_name") || InStr(lower, "peg_code"))
        return true
    if (InStr(lower, "plan_id") || InStr(lower, "===file: plans.csv"))
        return true
    if (InStr(lower, "filter,") && (InStr(lower, "title") || InStr(lower, "project")))
        return true
    ; Markdown fence containing CSV-looking content
    fence := Chr(96) . Chr(96) . Chr(96)
    if (InStr(t, fence) && (InStr(lower, "description") || InStr(lower, "entity_type")))
        return true
    return false
}

PackPipeline_BuildStructureFix(errorMsg) {
    global g_PackPipeline
    if (!PackPipeline_IsActive())
        return ""
    kind := g_PackPipeline.fixKind
    canonical := g_PackPipeline.canonical
    path := ""
    try {
        switch kind {
            case "finance_daily":
                path := Finance_WriteAiCompanionImportError(errorMsg, "daily")
            case "finance_monthly":
                path := Finance_WriteAiCompanionImportError(errorMsg, "monthly")
            case "palace":
                path := Palace_WriteAiCompanionImportError(errorMsg)
            case "task":
                ; Tasks normally write via Python; synthesize a Desktop fix file.
                path := PackPipeline_WriteGenericTaskFix(errorMsg)
            case "plan":
                path := PackPipeline_WriteGenericPlanFix(errorMsg)
            default:
                path := PackPipeline_WriteGenericFix(errorMsg, canonical)
        }
    } catch {
        path := PackPipeline_WriteGenericFix(errorMsg, canonical)
    }
    if (path = "" || !FileExist(path))
        return PackPipeline_GenericFixBody(errorMsg, canonical)
    body := ""
    try {
        f := FileOpen(path, "r", "UTF-8")
        if (f) {
            body := f.Read()
            f.Close()
            if (SubStr(body, 1, 1) = Chr(0xFEFF))
                body := SubStr(body, 2)
        }
    } catch {
        body := ""
    }
    return body != "" ? body : PackPipeline_GenericFixBody(errorMsg, canonical)
}

PackPipeline_GenericFixBody(errorMsg, canonical) {
    return "The last pack delivery failed validation. Fix and re-deliver.`r`n`r`n"
    . "IMPORT ERROR`r`n" . errorMsg . "`r`n`r`n"
        . "WHAT YOU MUST DO`r`n"
        . "- Re-emit one complete " . canonical .
        " with ===PREVIEW=== … ===END_PREVIEW=== and ===FILE: …=== sections.`r`n"
        . "- Prefer download chip; else one marked code fence.`r`n"
        . "- Never claim a Desktop/disk save.`r`n"
        . "- Use the exact canonical filename (" . canonical . ").`r`n"
}

PackPipeline_WriteGenericFix(errorMsg, canonical) {
    path := A_Desktop . "\PACK_AI_FIX.txt"
    body := PackPipeline_GenericFixBody(errorMsg, canonical)
    try {
        f := FileOpen(path, "w", "UTF-8")
        if (f) {
            f.Write(body)
            f.Close()
            return path
        }
    } catch {
    }
    return ""
}

PackPipeline_WriteGenericTaskFix(errorMsg) {
    path := A_Desktop . "\TASK_AI_FIX.txt"
    body := PackPipeline_GenericFixBody(errorMsg, "TASK_PACK.txt")
    try {
        f := FileOpen(path, "w", "UTF-8")
        if (f) {
            f.Write(body)
            f.Close()
            return path
        }
    } catch {
    }
    return ""
}

PackPipeline_WriteGenericPlanFix(errorMsg) {
    path := A_Desktop . "\PALACE_AI_FIX.txt"
    body := PackPipeline_GenericFixBody(errorMsg, "PLAN_PACK.txt")
    try {
        f := FileOpen(path, "w", "UTF-8")
        if (f) {
            f.Write(body)
            f.Close()
            return path
        }
    } catch {
    }
    return ""
}

PackPipeline_HandleInvalid(errorMsg, fixText := "") {
    global g_PackPipeline
    if (!PackPipeline_IsActive())
        return
    g_PackPipeline.attempt += 1
    if (g_PackPipeline.attempt > g_PackPipeline.maxAttempts) {
        PackPipeline_Fail("Pack still invalid after " . g_PackPipeline.maxAttempts . " fix attempt(s): " . errorMsg)
        return
    }
    if (Trim(fixText) = "")
        fixText := PackPipeline_BuildStructureFix(errorMsg)
    try ShowCenteredOverlay_Utils("⚠ Pack invalid — sending AI fix (" . g_PackPipeline.attempt . "/"
        . g_PackPipeline.maxAttempts . ")", 2500, BANNER_ACCENT_ERROR)
    catch {
    }
    if (!PackPipeline_SendFixAndSubmit(fixText)) {
        PackPipeline_Fail("Could not paste/submit AI fix")
        return
    }
    ; Re-enter wait for the corrected generation.
    global g_PackPipelineMonitorButtonSeen, g_PackPipelineMonitorRetry
    g_PackPipelineMonitorButtonSeen := false
    g_PackPipelineMonitorRetry := 0
    PackPipeline_StartMonitor()
}

; After domain import: if a fresh *_AI_FIX.txt was written, feed companion + re-monitor.
; If no fresh fix (user cancelled confirm), end the session.
PackPipeline_HandleImportOutcome(importOk, importStartStamp, errMsg := "") {
    if (!PackPipeline_IsActive())
        return
    if (importOk) {
        PackPipeline_Reset()
        return
    }
    fixPath := ""
    try fixPath := ImportWatcher_CompanionNewestAiFixSince(importStartStamp)
    catch {
        fixPath := ""
    }
    if (fixPath = "") {
        ; Cancel or soft fail without AI-fix file — stop pipeline, do not loop.
        PackPipeline_Reset()
        return
    }
    fixText := ""
    try fixText := ImportWatcher_CompanionReadUtf8(fixPath)
    catch {
        fixText := ""
    }
    if (Trim(fixMsg) = "")
        errMsg := "Import validation failed"
    PackPipeline_HandleInvalid(errMsg, fixText)
}

PackPipeline_SendFixAndSubmit(fixText) {
    global g_PackPipeline
    fixText := Trim(fixText)
    if (fixText = "" || !PackPipeline_IsActive())
        return false
    companionId := g_PackPipeline.companionId
    PackPipeline_CaptureUserHwnd()
    hwnd := 0
    try {
        switch companionId {
            case "enterprise":
                hwnd := GeminiEnterprise_NavigateFocusAndPaste(fixText, true)
            case "copilot":
                hwnd := CopilotWeb_NavigateFocusAndPaste(fixText, true)
            default:
                ; GeminiNavigateFocusAndPasteFirstSnippet second arg is switchToFirstTab, not submit.
                hwnd := GeminiNavigateFocusAndPasteFirstSnippet(fixText, false)
                if (hwnd)
                    try PromptPaste_SubmitWhenReady(hwnd, "gemini", 0)
                    catch {
                        try Gemini_WaitForPromptContentAndSubmit(hwnd)
                        catch {
                        }
                    }
        }
    } catch as e {
        try ShowCenteredOverlay_Utils("❌ AI fix send failed: " . e.Message, 2800, BANNER_ACCENT_ERROR)
        catch {
        }
        PackPipeline_RestoreUserHwnd()
        return false
    }
    if (hwnd)
        g_PackPipeline.hwnd := hwnd
    ; Restore user focus before re-entering background wait.
    PackPipeline_RestoreUserHwnd()
    PackPipeline_NotifyUserFree()
    return true
}

; --- Write Desktop + import ---

PackPipeline_WriteUtf8(path, text) {
    try {
        f := FileOpen(path, "w", "UTF-8")
        if (!f)
            return false
        f.Write(text)
        f.Close()
        return true
    } catch {
        return false
    }
}

PackPipeline_MarkSeenForWatcher(path) {
    global g_ImportWatcherSeen, g_ImportWatcherPending, g_ImportWatcherBusy
    if (path = "" || !FileExist(path))
        return
    key := ""
    stamp := ""
    try {
        key := StrLower(RTrim(path, "\"))
        mtime := FileGetTime(path, "M")
        size := FileGetSize(path)
        stamp := mtime . "|" . size
    } catch {
        return
    }
    try {
        if (!IsObject(g_ImportWatcherSeen))
            g_ImportWatcherSeen := Map()
        g_ImportWatcherSeen[key] := stamp
        if (IsObject(g_ImportWatcherPending) && g_ImportWatcherPending.Has(key))
            g_ImportWatcherPending.Delete(key)
    } catch {
    }
}

PackPipeline_WriteAndImport(text) {
    global g_PackPipeline, g_ImportWatcherBusy
    if (!PackPipeline_IsActive())
        return
    canonical := g_PackPipeline.canonical
    label := g_PackPipeline.label
    runFn := g_PackPipeline.run
    path := ""
    try path := PackImport_CanonicalDesktopPath(canonical)
    catch {
        path := A_Desktop . "\" . canonical
    }
    try StandardLoadingBar_Show("⏳ Saving " . canonical . "…", BANNER_ACCENT_INTERMEDIATE, { passive: false })
    catch {
    }
    ; Prevent Desktop watcher from double-importing this write.
    prevBusy := false
    try prevBusy := g_ImportWatcherBusy
    catch {
    }
    try g_ImportWatcherBusy := true
    catch {
    }
    ok := PackPipeline_WriteUtf8(path, text)
    if (!ok) {
        try g_ImportWatcherBusy := prevBusy
        catch {
        }
        try StandardLoadingBar_Hide(0)
        catch {
        }
        PackPipeline_Fail("Failed to write " . canonical)
        return
    }
    PackPipeline_MarkSeenForWatcher(path)
    try StandardLoadingBar_Update("✅ " . label . " saved — opening import confirm…", BANNER_ACCENT_SUCCESS)
    catch {
    }
    Sleep 400
    try StandardLoadingBar_Hide(0)
    catch {
    }
    ; Keep session alive through confirm so import-fail can re-arm the fix loop.
    importStartStamp := FormatTime(, "yyyyMMddHHmmss")
    importOk := false
    importErr := ""
    try {
        if (runFn) {
            result := runFn.Call()
            ; Strict == so empty/unset success returns are not treated as false.
            if (result == false)
                importOk := false
            else
                importOk := true
        } else {
            importOk := true
        }
    } catch as e {
        importOk := false
        importErr := e.Message
        try ShowCenteredOverlay_Utils("Import failed: " . e.Message, 2800, BANNER_ACCENT_ERROR)
        catch {
            TrayTip("Pack pipeline", "Import failed")
        }
    } finally {
        try g_ImportWatcherBusy := false
        catch {
        }
    }
    PackPipeline_HandleImportOutcome(importOk, importStartStamp, importErr)
}

PackPipeline_Fail(msg) {
    try StandardLoadingBar_Hide(0)
    catch {
    }
    try ShowCenteredOverlay_Utils("❌ Pack pipeline: " . msg, 4000, BANNER_ACCENT_ERROR)
    catch {
        TrayTip("Pack pipeline", msg)
    }
    PackPipeline_Reset()
}
