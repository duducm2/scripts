; =============================================================================
; Utils module: import_watcher_companion.ahk
; Route Desktop AI-fix text into the active (or resolved) AI companion prompt.
; Watcher after-import path: paste only. Hub / manual path: paste+submit when
; PackPipeline is inactive (pipeline-active fixes are submitted by PackPipeline).
; Agent docs: docs/prompt-data-output-and-finance-packs.md
; =============================================================================

ImportWatcher_CompanionAiFixPaths() {
    return [
        A_Desktop . "\FINANCE_AI_FIX.txt",
        A_Desktop . "\TASK_AI_FIX.txt",
        A_Desktop . "\PALACE_AI_FIX.txt",
        A_Desktop . "\PACK_AI_FIX.txt"
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

; Paste + submit AI-fix text (hub / non-pipeline path). When PackPipeline is active,
; prefer PackPipeline_SendFixAndSubmit so userHwnd restore + re-monitor stay consistent.
; pathOrText: Desktop fix path or raw fix body.
ImportWatcher_CompanionPasteAndSubmitFix(pathOrText) {
    fixText := ""
    if (pathOrText != "" && FileExist(pathOrText))
        fixText := ImportWatcher_CompanionReadUtf8(pathOrText)
    else
        fixText := pathOrText
    fixText := Trim(fixText)
    if (fixText = "")
        return false
    try {
        if (PackPipeline_IsActive())
            return !!PackPipeline_SendFixAndSubmit(fixText)
    } catch {
    }
    companionId := ImportWatcher_CompanionResolveTarget()
    label := ImportWatcher_CompanionLabel(companionId)
    hwnd := 0
    try {
        switch companionId {
            case "enterprise":
                hwnd := GeminiEnterprise_NavigateFocusAndPaste(fixText, true)
            case "copilot":
                hwnd := CopilotWeb_NavigateFocusAndPaste(fixText, true)
            default:
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
            TrayTip("Import", "AI fix send failed")
        }
        return false
    }
    msg := "AI fix sent to " . label
    try ShowCenteredOverlay_Utils(msg, 2800, BANNER_ACCENT_INFO)
    catch {
        TrayTip("Import", msg)
    }
    return !!hwnd
}

; After an import run: if a fresh AI-fix file appeared, paste it into the companion.
ImportWatcher_CompanionHandleAiFixAfterImport(importStartStamp) {
    path := ImportWatcher_CompanionNewestAiFixSince(importStartStamp)
    if (path = "")
        return false
    body := ImportWatcher_CompanionReadUtf8(path)
    if (body = "")
        return false
    return ImportWatcher_CompanionPasteFixText(body)
}
