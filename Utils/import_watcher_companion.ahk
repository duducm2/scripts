; =============================================================================
; Utils module: import_watcher_companion.ahk
; Route Desktop AI-fix text into the active (or resolved) AI companion prompt.
; Watcher after-import: submit once if this run has not already sent that file.
; Never paste the same fix into the composer again after a successful send.
; Hub / manual path: paste+submit when PackPipeline is inactive
; (pipeline-active fixes are submitted by PackPipeline).
; Agent docs: docs/prompt-data-output-and-finance-packs.md
; =============================================================================

; Desktop path of the AI-fix file submitted during the current watcher import.
global g_ImportWatcherAiFixSubmittedPath := ""
; One confirmed AI-fix send. While this is in flight, no import pastes another fix.
global g_AiFixGateActive := false
global g_AiFixGateFp := ""
global g_AiFixGateHwnd := 0
global g_AiFixGateCompanion := ""
global g_AiFixGateAt := 0
global g_AiFixGateSawStop := false
global g_AiFixGateMisses := 0

ImportWatcher_CompanionFixFingerprint(fixText) {
    fixText := StrReplace(StrReplace(Trim(fixText), "`r`n", "`n"), "`r", "`n")
    return fixText
}

ImportWatcher_CompanionFindHwnd(companionId) {
    companionId := StrLower(Trim(companionId))
    hwnd := 0
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
    return hwnd
}

ImportWatcher_CompanionReleaseAiFixGate() {
    global g_AiFixGateActive, g_AiFixGateFp, g_AiFixGateHwnd, g_AiFixGateCompanion
    global g_AiFixGateAt, g_AiFixGateSawStop, g_AiFixGateMisses
    g_AiFixGateActive := false
    g_AiFixGateFp := ""
    g_AiFixGateHwnd := 0
    g_AiFixGateCompanion := ""
    g_AiFixGateAt := 0
    g_AiFixGateSawStop := false
    g_AiFixGateMisses := 0
}

ImportWatcher_CompanionNoteAiFixSent(fixText, hwnd, companionId) {
    global g_AiFixGateActive, g_AiFixGateFp, g_AiFixGateHwnd, g_AiFixGateCompanion
    global g_AiFixGateAt, g_AiFixGateSawStop, g_AiFixGateMisses
    g_AiFixGateActive := true
    g_AiFixGateFp := ImportWatcher_CompanionFixFingerprint(fixText)
    g_AiFixGateHwnd := hwnd
    g_AiFixGateCompanion := StrLower(Trim(companionId))
    g_AiFixGateAt := A_TickCount
    g_AiFixGateSawStop := true
    g_AiFixGateMisses := 0
}

; True while a fix we already handed the companion is still being answered.
ImportWatcher_CompanionAiFixInFlight() {
    global g_AiFixGateActive, g_AiFixGateHwnd, g_AiFixGateCompanion, g_AiFixGateAt
    global g_AiFixGateSawStop, g_AiFixGateMisses
    if (!g_AiFixGateActive)
        return false
    hwnd := g_AiFixGateHwnd
    companionId := g_AiFixGateCompanion
    if ((!hwnd || !WinExist("ahk_id " hwnd)) && companionId != "")
        hwnd := ImportWatcher_CompanionFindHwnd(companionId)
    generating := false
    if (hwnd && companionId != "") {
        try generating := !!PromptPaste_CompanionIsGenerating(hwnd, companionId)
        catch {
            generating := false
        }
    }
    if (generating) {
        g_AiFixGateHwnd := hwnd
        g_AiFixGateSawStop := true
        g_AiFixGateMisses := 0
        return true
    }
    ; Stop can blink off, and a short reply can finish before the next poll. Hold the gate
    ; so a second import does not paste another fix while this one is still the active send.
    if (g_AiFixGateAt && (A_TickCount - g_AiFixGateAt) < 15000)
        return true
    if (!g_AiFixGateSawStop) {
        ImportWatcher_CompanionReleaseAiFixGate()
        return false
    }
    g_AiFixGateMisses += 1
    if (g_AiFixGateMisses < 2)
        return true
    ImportWatcher_CompanionReleaseAiFixGate()
    return false
}

; Hold the gate before the paste so a second caller cannot send while this one is still confirming.
ImportWatcher_CompanionClaimAiFixSend(fixText) {
    global g_AiFixGateActive, g_AiFixGateFp, g_AiFixGateAt, g_AiFixGateSawStop, g_AiFixGateMisses
    fp := ImportWatcher_CompanionFixFingerprint(fixText)
    if (fp = "")
        return false
    if (g_AiFixGateActive)
        return false
    g_AiFixGateActive := true
    g_AiFixGateFp := fp
    g_AiFixGateAt := A_TickCount
    g_AiFixGateSawStop := false
    g_AiFixGateMisses := 0
    return true
}

ImportWatcher_CompanionClearAiFixSubmitted() {
    global g_ImportWatcherAiFixSubmittedPath
    g_ImportWatcherAiFixSubmittedPath := ""
}

ImportWatcher_CompanionMarkAiFixSubmitted(path) {
    global g_ImportWatcherAiFixSubmittedPath
    if (path != "")
        g_ImportWatcherAiFixSubmittedPath := path
}

; Remember the newest Desktop fix file whose body matches the text just sent.
ImportWatcher_CompanionMarkAiFixSubmittedText(fixText) {
    fixText := Trim(fixText)
    if (fixText = "")
        return
    bestPath := ""
    bestStamp := ""
    for path in ImportWatcher_CompanionAiFixPaths() {
        if (path = "" || !FileExist(path))
            continue
        body := Trim(ImportWatcher_CompanionReadUtf8(path))
        if (body = "" || body != fixText)
            continue
        stamp := ""
        try stamp := FileGetTime(path, "M")
        catch {
            continue
        }
        if (bestStamp = "" || stamp > bestStamp) {
            bestStamp := stamp
            bestPath := path
        }
    }
    ImportWatcher_CompanionMarkAiFixSubmitted(bestPath)
}

ImportWatcher_CompanionAiFixPaths() {
    return [
        A_Desktop . "\FINANCE_AI_FIX.txt",
        A_Desktop . "\TASK_AI_FIX.txt",
        A_Desktop . "\PALACE_AI_FIX.txt",
        A_Desktop . "\PACK_AI_FIX.txt",
        A_Desktop . "\COLLECTIBLE_AI_FIX.txt"
    ]
}

; Return newest AI-fix Desktop path with mtime > sinceTick (AHK YYYYMMDDHH24MISS), or "".
ImportWatcher_CompanionNewestAiFixSince(sinceStamp) {
    bestPath := ""
    bestStamp := ""
    for path in ImportWatcher_CompanionAiFixPaths() {
        if (path = "" || !FileExist(path))
            continue
        stamp := ""
        try stamp := FileGetTime(path, "M")
        catch {
            continue
        }
        if (sinceStamp != "" && stamp < sinceStamp)
            continue
        if (bestStamp = "" || stamp > bestStamp) {
            bestStamp := stamp
            bestPath := path
        }
    }
    return bestPath
}

ImportWatcher_CompanionReadUtf8(path) {
    if (path = "" || !FileExist(path))
        return ""
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
        return ""
    }
    return body
}

; Prefer the foreground companion window when it is Gemini / Copilot / Enterprise.
ImportWatcher_CompanionDetectForeground() {
    hwnd := 0
    try hwnd := WinGetID("A")
    catch {
        hwnd := 0
    }
    if (!hwnd)
        return ""
    proc := ""
    title := ""
    try proc := StrLower(WinGetProcessName("ahk_id " hwnd))
    catch {
        proc := ""
    }
    if (proc != "chrome.exe")
        return ""
    try title := WinGetTitle("ahk_id " hwnd)
    catch {
        title := ""
    }
    if (title = "")
        return ""
    try {
        if (GeminiEnterprise_TitleMatches(title))
            return "enterprise"
    } catch {
    }
    try {
        if (CopilotWeb_TitleMatchesCopilot(title))
            return "copilot"
    } catch {
    }
    try {
        if (IsConsumerGeminiChromeTitle(title))
            return "gemini"
    } catch {
    }
    return ""
}

ImportWatcher_CompanionResolveTarget() {
    fg := ImportWatcher_CompanionDetectForeground()
    if (fg != "")
        return fg
    try return ResolveGlobalAICompanion()
    catch {
        return "gemini"
    }
}

ImportWatcher_CompanionLabel(companionId) {
    switch companionId {
        case "enterprise":
            return "Gemini Enterprise"
        case "copilot":
            return "Copilot"
        default:
            return "Gemini"
    }
}

; Paste fixText into the resolved companion prompt. Does not submit.
; Returns true when paste was attempted on a focused companion.
ImportWatcher_CompanionPasteFixText(fixText) {
    fixText := Trim(fixText)
    if (fixText = "")
        return false
    companionId := ImportWatcher_CompanionResolveTarget()
    label := ImportWatcher_CompanionLabel(companionId)
    hwnd := 0
    try {
        switch companionId {
            case "enterprise":
                hwnd := GeminiEnterprise_NavigateFocusAndPaste(fixText, false)
            case "copilot":
                hwnd := CopilotWeb_NavigateFocusAndPaste(fixText, false)
            default:
                hwnd := GeminiNavigateFocusAndPasteFirstSnippet(fixText, false)
        }
    } catch as e {
        try ShowCenteredOverlay_Utils("❌ AI fix paste failed: " . e.Message, 2800, BANNER_ACCENT_ERROR)
        catch {
            TrayTip("Import", "AI fix paste failed")
        }
        return false
    }
    msg := "AI fix pasted to " . label . " — review before sending"
    try ShowCenteredOverlay_Utils(msg, 3500, BANNER_ACCENT_INFO)
    catch {
        TrayTip("Import", msg)
    }
    return !!hwnd
}

; Paste + submit one AI fix. A second call while that fix is still with the companion does not paste again.
ImportWatcher_CompanionSubmitAiFixOnce(fixText) {
    fixText := Trim(fixText)
    if (fixText = "")
        return false
    if (ImportWatcher_CompanionAiFixInFlight() || !ImportWatcher_CompanionClaimAiFixSend(fixText))
        return true
    companionId := ImportWatcher_CompanionResolveTarget()
    label := ImportWatcher_CompanionLabel(companionId)
    hwnd := 0
    submitted := false
    try {
        switch companionId {
            case "enterprise":
                hwnd := GeminiEnterprise_NavigateFocusAndPaste(fixText, true)
                submitted := !!hwnd
                if (!submitted) {
                    hwnd := ImportWatcher_CompanionFindHwnd("enterprise")
                    if (hwnd)
                        submitted := !!PromptPaste_WaitForGenerationStarted(hwnd, "enterprise", 8000)
                }
            case "copilot":
                hwnd := CopilotWeb_NavigateFocusAndPaste(fixText, true)
                submitted := !!hwnd
                if (!submitted) {
                    hwnd := ImportWatcher_CompanionFindHwnd("copilot")
                    if (hwnd)
                        submitted := !!PromptPaste_WaitForGenerationStarted(hwnd, "copilot", 8000)
                }
            default:
                hwnd := GeminiNavigateFocusAndPasteFirstSnippet(fixText, false)
                if (hwnd) {
                    try submitted := !!PromptPaste_SubmitWhenReady(hwnd, "gemini", 0)
                    catch {
                        try submitted := !!Gemini_WaitForPromptContentAndSubmit(hwnd)
                        catch
                            submitted := false
                    }
                    if (!submitted)
                        submitted := !!PromptPaste_WaitForGenerationStarted(hwnd, "gemini", 8000)
                }
        }
    } catch as e {
        ImportWatcher_CompanionReleaseAiFixGate()
        try ShowCenteredOverlay_Utils("❌ AI fix send failed: " . e.Message, 2800, BANNER_ACCENT_ERROR)
        catch {
            TrayTip("Import", "AI fix send failed")
        }
        return false
    }
    if (!submitted) {
        ImportWatcher_CompanionReleaseAiFixGate()
        return false
    }
    ImportWatcher_CompanionNoteAiFixSent(fixText, hwnd, companionId)
    msg := "AI fix sent to " . label
    try ShowCenteredOverlay_Utils(msg, 2800, BANNER_ACCENT_INFO)
    catch {
        TrayTip("Import", msg)
    }
    return true
}

; Paste + submit AI-fix text (hub / non-pipeline path). When PackPipeline is active,
; prefer PackPipeline_SendFixAndSubmit so userHwnd restore + re-monitor stay consistent.
; pathOrText: Desktop fix path or raw fix body.
ImportWatcher_CompanionPasteAndSubmitFix(pathOrText) {
    fixText := ""
    submittedPath := ""
    if (pathOrText != "" && FileExist(pathOrText)) {
        submittedPath := pathOrText
        fixText := ImportWatcher_CompanionReadUtf8(pathOrText)
    } else
        fixText := pathOrText
    fixText := Trim(fixText)
    if (fixText = "")
        return false
    try {
        if (PackPipeline_IsActive())
            return !!PackPipeline_SendFixAndSubmit(fixText)
    } catch {
    }
    if (ImportWatcher_CompanionAiFixInFlight())
        return true
    ok := ImportWatcher_CompanionSubmitAiFixOnce(fixText)
    if (!ok)
        return false
    if (submittedPath != "")
        ImportWatcher_CompanionMarkAiFixSubmitted(submittedPath)
    else
        ImportWatcher_CompanionMarkAiFixSubmittedText(fixText)
    return true
}

; After an import run: send a fresh AI-fix once. Skip when a fix is already with the companion.
ImportWatcher_CompanionHandleAiFixAfterImport(importStartStamp) {
    global g_ImportWatcherAiFixSubmittedPath
    if (ImportWatcher_CompanionAiFixInFlight())
        return false
    path := ImportWatcher_CompanionNewestAiFixSince(importStartStamp)
    if (path = "")
        return false
    if (path = g_ImportWatcherAiFixSubmittedPath)
        return false
    return ImportWatcher_CompanionPasteAndSubmitFix(path)
}
