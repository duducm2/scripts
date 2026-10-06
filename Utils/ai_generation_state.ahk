; =============================================================================
; Utils module: ai_generation_state.ahk
; Global AI generation state U macro
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; Global AI generation state: Cursor + Gemini stop-button detectors (Efficiency Canon)
; =============================================================================
; Cursor: Type 50026 (Group), ClassName contains "stop-button".
; Gemini: Chrome window title contains "gemini"; Type 50000, Name "Stop response", ClassName match.
; =============================================================================
Cursor_HasGeneratingStopButton() {
    global UIA
    try {
        cursorHwnds := WinGetList("ahk_exe Cursor.exe")
        if (!cursorHwnds.Length)
            return false
        cr := UIA.CreateCacheRequest(["Type", "ClassName"], , 5)
        for hwnd in cursorHwnds {
            if (!hwnd || !WinExist("ahk_id " hwnd))
                continue
            try {
                root := UIA.ElementFromHandleBuildCache(cr, hwnd)
            } catch {
                try root := UIA.ElementFromHandle(hwnd)
                catch
                    continue
            }
            if (!root)
                continue
            try {
                el := root.FindFirstBuildCache(cr, { Type: 50026, ClassName: "stop-button", matchmode: "Substring" }, 4
                )
                if (el)
                    return true
            } catch {
            }
        }
    } catch {
    }
    return false
}

Gemini_HasGeneratingStopButton() {
    global UIA
    try {
        for hwnd in WinGetList("ahk_exe chrome.exe") {
            if (!hwnd || !WinExist("ahk_id " hwnd))
                continue
            try {
                if (!IsConsumerGeminiChromeTitle(WinGetTitle("ahk_id " hwnd)))
                    continue
            } catch {
                continue
            }
            try {
                cr := UIA.CreateCacheRequest(["Type", "ClassName", "Name"], , 5)
                root := UIA.ElementFromHandleBuildCache(cr, hwnd)
            } catch {
                try root := UIA.ElementFromHandle(hwnd)
                catch
                    continue
            }
            if (!root)
                continue
            ; Stop response: Type 50000, Name "Stop response", ClassName contains "send-button" and "stop"
            try {
                el := root.FindFirstBuildCache(cr, { Type: 50000, Name: "Stop response", ClassName: "send-button",
                    matchmode: "Substring" }, 4)
                if (el)
                    return true
            } catch {
            }
        }
    } catch {
    }
    return false
}

IsAnyAiGenerating() {
    return Cursor_HasGeneratingStopButton() || Gemini_HasGeneratingStopButton()
}

; =============================================================================
; After Enter: prove the composer let go of the prompt and the stop control is up.
; Returns working | held | empty | unconfirmed | unreadable.
; snapOk false: the composer could not be read before Enter, so a blank sentText
; is not "empty" — only a visible stop control counts, otherwise unreadable.
; =============================================================================
AiCompanion_IsComposerPlaceholder(text) {
    ; A blank Value is not proof the field is empty. ProseMirror often reports no value
    ; while the prompt is still visible.
    t := Trim(text)
    if (t = "")
        return false
    for p in ["Ask anything", "Ask Gemini", "Message Copilot", "Message ChatGPT", "Pergunte qualquer coisa"] {
        if (t = p)
            return true
    }
    return false
}

AiCompanion_NormalizeComposerText(text) {
    return Trim(RegExReplace(text, "\s+", " "))
}

AiCompanion_ComposerHolds(composerText, sentText) {
    sent := AiCompanion_NormalizeComposerText(sentText)
    if (sent = "")
        return false
    cur := AiCompanion_NormalizeComposerText(composerText)
    if (cur = "" || AiCompanion_IsComposerPlaceholder(cur))
        return false
    needle := SubStr(sent, 1, 80)
    return InStr(cur, needle, false) > 0
}

AiCompanion_ReadElementText(el, &text) {
    text := ""
    if (!IsObject(el))
        return false
    got := false
    try {
        text := Trim(el.Value)
        got := true
    } catch {
    }
    if (text = "") {
        try {
            text := Trim(el.TextPattern.DocumentRange.GetText(-1))
            got := true
        } catch {
        }
    }
    if (!got)
        return false
    if (AiCompanion_IsComposerPlaceholder(text))
        text := ""
    return true
}

; One targeted FindFirst. No FindAll, no UIA_Browser (that activates Chrome and walks the document).
AiCompanion_FindComposerElement(root, companionId) {
    if (!IsObject(root))
        return 0
    companionId := StrLower(Trim(companionId))
    try {
        if (companionId = "enterprise")
            return GeminiEnterprise_FindComposer(root)
        if (companionId = "copilot")
            return CopilotWeb_FindComposer(root)
        if (companionId = "chatgpt") {
            try {
                saved := UiElements_TrySaved(root, "ChatGPT", "Prompt")
                if (saved)
                    return saved
            } catch {
            }
            try {
                el := root.FindFirst({ AutomationId: "prompt-textarea" })
                if (el)
                    return el
            } catch {
            }
            for name in ["Message ChatGPT", "Mensagem ChatGPT"] {
                try {
                    el := root.FindFirst({ Name: name, Type: 50004 })
                    if (el)
                        return el
                } catch {
                }
            }
            return 0
        }
        return FindGeminiPromptField(root)
    } catch {
    }
    return 0
}

AiCompanion_ComposerSlot(hwnd, companionId, el := 0, remember := false) {
    static slot := { hwnd: 0, companionId: "", el: 0, tick: 0 }
    companionId := StrLower(Trim(companionId))
    if (remember) {
        slot.hwnd := hwnd
        slot.companionId := companionId
        slot.el := el
        slot.tick := A_TickCount
        return el
    }
    if (!IsObject(slot.el) || slot.hwnd != hwnd || slot.companionId != companionId)
        return 0
    if ((A_TickCount - slot.tick) > 2500)
        return 0
    return slot.el
}

; Value only. TextPattern.GetText walks the element and is reserved for the one-time snapshot.
AiCompanion_ReadElementValue(el, &text) {
    text := ""
    if (!IsObject(el))
        return false
    try {
        text := Trim(el.Value)
        if (AiCompanion_IsComposerPlaceholder(text))
            text := ""
        return true
    } catch {
    }
    return false
}

AiCompanion_StopOnRoot(root, companionId) {
    if (!IsObject(root))
        return false
    companionId := StrLower(Trim(companionId))
    try {
        if (companionId = "enterprise")
            return !!GeminiEnterprise_FindStopButton(root)
        if (companionId = "copilot")
            return !!CopilotWeb_FindStopGenerating(root)
        if (companionId = "chatgpt") {
            try {
                if (UiElements_TrySaved(root, "ChatGPT", "Stop"))
                    return true
            } catch {
            }
            for n in ["Stop streaming", "Interromper transmissão"] {
                try {
                    if (root.FindFirst({ Name: n, Type: 50000 }))
                        return true
                } catch {
                }
            }
            return false
        }
        return Gemini_HasGeneratingStopButtonForUia(root)
    } catch {
    }
    return false
}

; "ok" and text (maybe blank) when the composer element was read; "missing" otherwise.
AiCompanion_ReadComposer(hwnd, companionId, &text) {
    text := ""
    companionId := StrLower(Trim(companionId))
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return "missing"
    try {
        root := UIA.ElementFromHandle(hwnd)
        if (!IsObject(root))
            return "missing"
        pf := AiCompanion_ComposerSlot(hwnd, companionId)
        if (!pf)
            pf := AiCompanion_FindComposerElement(root, companionId)
        if (pf)
            AiCompanion_ReadElementText(pf, &text)
        if (Trim(text) = "") {
            ; Value/TextPattern stay blank on several composers. Clipboard is the read that sees the text.
            text := AiCompanion_ComposerClipboardFallback(hwnd, companionId)
        }
        if (Trim(text) = "")
            return "missing"
        if (pf)
            AiCompanion_ComposerSlot(hwnd, companionId, pf, true)
        return "ok"
    } catch {
    }
    return "missing"
}

; Clipboard read when UIA Value/TextPattern cannot see the composer. Empty string if unavailable.
AiCompanion_ComposerClipboardFallback(hwnd, companionId) {
    companionId := StrLower(Trim(companionId))
    try {
        if (companionId = "enterprise")
            return GeminiEnterprise_ComposerGetText(hwnd)
        if (companionId = "copilot")
            return CopilotWeb_ComposerGetTextViaClipboard(hwnd)
        if (companionId = "chatgpt")
            return ChatGPT_ComposerGetTextViaClipboard(hwnd)
    } catch {
    }
    return ""
}

AiCompanion_SnapshotComposer(hwnd, companionId, &status) {
    text := ""
    status := AiCompanion_ReadComposer(hwnd, companionId, &text)
    return text
}

AiCompanion_IsGenerating(hwnd, companionId) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    try {
        root := UIA.ElementFromHandle(hwnd)
        return AiCompanion_StopOnRoot(root, companionId)
    } catch {
    }
    return false
}

; Runs only after a tracked Enter. One composer lookup, then at most a few stop-button
; checks on that same window. Never UIA_Browser, never FindAll, never a background timer.
; A prompt still in the composer returns in ~200ms. Hard cap is 1s even if timeoutMs is higher.
AiCompanion_ConfirmAfterEnter(hwnd, companionId, sentText, timeoutMs := 1000, snapOk := true) {
    companionId := StrLower(Trim(companionId))
    if (companionId = "" || !hwnd || !WinExist("ahk_id " hwnd))
        return "unreadable"
    if (snapOk && Trim(sentText) = "")
        return "empty"
    if (timeoutMs > 1000)
        timeoutMs := 1000
    if (timeoutMs < 200)
        timeoutMs := 200
    tStart := A_TickCount
    root := 0
    try root := UIA.ElementFromHandle(hwnd)
    catch
        root := 0
    if (!IsObject(root))
        return "unreadable"
    composer := AiCompanion_ComposerSlot(hwnd, companionId)
    if (!composer)
        composer := AiCompanion_FindComposerElement(root, companionId)
    if (!composer)
        return "unreadable"
    cur := ""
    if (snapOk) {
        if (!AiCompanion_ReadElementValue(composer, &cur) || Trim(cur) = "") {
            cur := AiCompanion_ComposerClipboardFallback(hwnd, companionId)
            if (Trim(cur) = "" && !AiCompanion_ReadElementText(composer, &cur))
                cur := ""
        }
    }
    if (snapOk && AiCompanion_ComposerHolds(cur, sentText)) {
        Sleep 200
        ; A blank clipboard read means the composer let go. Only a failed read keeps the snapshot.
        cur := sentText
        try cur := AiCompanion_ComposerClipboardFallback(hwnd, companionId)
        catch
            cur := sentText
        if (Trim(cur) = "" && companionId != "enterprise" && companionId != "copilot" && companionId != "chatgpt") {
            if (!AiCompanion_ReadElementValue(composer, &cur))
                cur := sentText
        }
        if (AiCompanion_ComposerHolds(cur, sentText))
            return "held"
    }
    deadline := tStart + timeoutMs
    loop {
        if (AiCompanion_StopOnRoot(root, companionId))
            return "working"
        if (A_TickCount >= deadline)
            break
        Sleep 200
    }
    return "unconfirmed"
}

AiCompanion_AnnounceConfirm(state) {
    global BANNER_ACCENT_SUCCESS, BANNER_ACCENT_ERROR
    switch state {
        case "working":
            msg := "Prompt received — AI is working"
            color := BANNER_ACCENT_SUCCESS
            ms := 1800
        case "held":
            msg := "Prompt still in the composer"
            color := BANNER_ACCENT_ERROR
            ms := 2200
        case "empty":
            msg := "Composer was empty"
            color := BANNER_ACCENT_ERROR
            ms := 2200
        case "unconfirmed":
            msg := "No stop control — not confirmed"
            color := BANNER_ACCENT_ERROR
            ms := 2200
        default:
            msg := "Could not read this companion"
            color := BANNER_ACCENT_ERROR
            ms := 2200
    }
    try ShowCenteredOverlay_Utils(msg, ms, color)
    catch {
    }
    return state
}

; True only when the stop control is visible. Announces every result.
AiCompanion_TrackAfterEnter(hwnd, companionId, sentText, snapOk := true) {
    if (snapOk && Trim(sentText) = "") {
        AiCompanion_AnnounceConfirm("empty")
        return false
    }
    state := AiCompanion_ConfirmAfterEnter(hwnd, companionId, sentText, 3000, snapOk)
    AiCompanion_AnnounceConfirm(state)
    return state = "working"
}

; Physical Enter/Ctrl+Enter is delivered immediately. Stop-button confirm runs on a short timer.
; Set false to restore snapshot-then-submit on those hotkeys.
global AI_COMPANION_ENTER_SEND_FIRST := true
; When true, a confirmed Enter/Send arms the shared response banner instead of a chime-only watch.
global AI_COMPANION_D2C_BANNER_FOR_ALL := true
; ChatGPT keeps its loading-bar wait when this is false.
global AI_COMPANION_CHATGPT_ENABLE_D2C := true
; Copilot hotkey Enter uses CopilotWeb_SubmitComposer when true, else the thinner Enter-only submit.
global AI_COMPANION_USE_SUBMIT_COMPOSER_FOR_COPILOT := true

; Set only when a catalog Send click returns true. Readers take it once.
global g_AiCompanionCatalogSendClicked := false
; True when ArmResponseWatch started the D2C monitor for this click.
global g_AiCompanionResponseWatchArmedBySend := false

AiCompanion_MarkCatalogSendClick() {
    global g_AiCompanionCatalogSendClicked
    g_AiCompanionCatalogSendClicked := true
}

AiCompanion_TakeCatalogSendClick() {
    global g_AiCompanionCatalogSendClicked
    clicked := !!g_AiCompanionCatalogSendClicked
    g_AiCompanionCatalogSendClicked := false
    return clicked
}

AiCompanion_UseD2CBanner(companionId) {
    global AI_COMPANION_D2C_BANNER_FOR_ALL, AI_COMPANION_CHATGPT_ENABLE_D2C
    if (!AI_COMPANION_D2C_BANNER_FOR_ALL)
        return false
    if (StrLower(Trim(companionId)) = "chatgpt" && !AI_COMPANION_CHATGPT_ENABLE_D2C)
        return false
    return true
}

; Clears the one-shot Send flag and arms the shared response banner.
; True means this submit owns the D2C path (do not also start a chime-only watch).
AiCompanion_FinishConfirmedSubmit(hwnd, companionId, originHwnd := 0) {
    AiCompanion_TakeCatalogSendClick()
    if (!AiCompanion_UseD2CBanner(companionId))
        return false
    AiCompanion_ArmResponseWatch(hwnd, companionId, originHwnd)
    return true
}

; One existing completion watch after a catalog Send click. No UIA and no Send search.
; Pack pipeline and an in-progress D2C watch keep their own timer.
AiCompanion_ArmResponseWatch(hwnd, companionId, originHwnd := 0) {
    global g_AiCompanionResponseWatchArmedBySend
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    try {
        if (PackPipeline_IsActive())
            return false
    } catch {
        return false
    }
    flow := D2C_FlowManager.GetInstance()
    phase := flow.CurrentPhase
    if (phase = "Monitoring" || phase = "PromptingAction")
        return false
    flow.GeminiHwnd := hwnd
    flow.CompanionId := StrLower(Trim(companionId))
    if (originHwnd && originHwnd != hwnd && WinExist("ahk_id " originHwnd))
        flow.OriginHwnd := originHwnd
    flow.StartGeminiMonitor()
    g_AiCompanionResponseWatchArmedBySend := true
    return true
}

; Stop the watch this click started so pack pipeline is the only Stop poll.
AiCompanion_DisarmResponseWatch() {
    global g_AiCompanionResponseWatchArmedBySend
    if (!g_AiCompanionResponseWatchArmedBySend)
        return false
    g_AiCompanionResponseWatchArmedBySend := false
    flow := D2C_FlowManager.GetInstance()
    if (flow.CurrentPhase != "Monitoring")
        return false
    if (flow.MonitorTimer != "") {
        try SetTimer(flow.MonitorTimer, 0)
    }
    flow.MonitorTimer := ""
    flow.CurrentPhase := "Submitting"
    return true
}

; One-shot abort of in-flight reply watches. Posted to the other scripts that include Utils.
WM_ABORT_AI_COMPANION_REPLY_WATCH := 0x8009

; Clear this process only. Does not post, and does not touch PromptingSubmit / PromptingAction.
AiCompanion_AbortReplyWatchesLocal() {
    global g_AiCompanionEnterWatchCb
    if (g_AiCompanionEnterWatchCb != "") {
        try SetTimer(g_AiCompanionEnterWatchCb, 0)
        g_AiCompanionEnterWatchCb := ""
    }
    try {
        flow := D2C_FlowManager.GetInstance()
        if (flow.CurrentPhase = "Monitoring")
            flow.Reset()
    } catch {
    }
    try {
        if (PackPipeline_IsActive())
            PackPipeline_Reset()
    } catch {
    }
    try GeminiEnterprise_StopGenerationWatch()
    catch {
    }
    try CopilotWeb_StopGenerationWatch()
    catch {
    }
    AiCompanion_StopGeminiDelayedSubmitMonitor()
}

; Gemini.ahk owns the delayed-submit timer. Other scripts ask it over 0x8003.
; SendMessage to our own window would wait on this same thread.
AiCompanion_StopGeminiDelayedSubmitMonitor() {
    if (A_ScriptName = "Gemini.ahk") {
        try GeminiDelayedSubmitMonitorStop()
        catch {
        }
        return
    }
    try GeminiDelayedSubmitMonitorStopFromUtils()
    catch {
    }
}

AiCompanion_FindAhkScriptHwnd(titleNeedle) {
    prevDetect := A_DetectHiddenWindows
    prevMatch := A_TitleMatchMode
    DetectHiddenWindows true
    SetTitleMatchMode 2
    hwnd := 0
    try hwnd := WinExist(titleNeedle " ahk_class AutoHotkey")
    catch
        hwnd := 0
    if (!hwnd) {
        for exe in ["AutoHotkey64.exe", "AutoHotkey32.exe", "AutoHotkey.exe"] {
            try list := WinGetList("ahk_exe " exe)
            catch
                continue
            for candidate in list {
                title := ""
                try title := WinGetTitle("ahk_id " candidate)
                catch
                    continue
                if (InStr(title, titleNeedle)) {
                    hwnd := candidate
                    break
                }
            }
            if (hwnd)
                break
        }
    }
    DetectHiddenWindows prevDetect
    SetTitleMatchMode prevMatch
    return hwnd
}

AiCompanion_BroadcastAbortReplyWatches() {
    global WM_ABORT_AI_COMPANION_REPLY_WATCH
    selfHwnd := 0
    try selfHwnd := A_ScriptHwnd
    catch
        selfHwnd := 0
    prevDetect := A_DetectHiddenWindows
    DetectHiddenWindows true
    for name in ["Shift keys.ahk", "AppLaunchers.ahk", "Gemini.ahk"] {
        hwnd := AiCompanion_FindAhkScriptHwnd(name)
        if (!hwnd || hwnd = selfHwnd)
            continue
        try PostMessage(WM_ABORT_AI_COMPANION_REPLY_WATCH, 0, 0, , "ahk_id " hwnd)
        catch {
        }
    }
    DetectHiddenWindows prevDetect
}

; Menu path: clear this process, ask the other hosts to clear theirs, then one overlay.
AiCompanion_AbortReplyWatches() {
    AiCompanion_AbortReplyWatchesLocal()
    AiCompanion_BroadcastAbortReplyWatches()
    try ShowCenteredOverlay_Utils("Reply watch stopped", 1500, BANNER_ACCENT_SUCCESS)
    catch {
    }
}

AiCompanion_OnAbortReplyWatches(*) {
    AiCompanion_AbortReplyWatchesLocal()
}

OnMessage(WM_ABORT_AI_COMPANION_REPLY_WATCH, AiCompanion_OnAbortReplyWatches)

; Snapshot, skip a blank composer, call sendFn, then track. True only when working.
AiCompanion_SendAndConfirm(hwnd, companionId, sendFn) {
    snapStatus := ""
    sentText := AiCompanion_SnapshotComposer(hwnd, companionId, &snapStatus)
    if (snapStatus = "ok" && Trim(sentText) = "") {
        AiCompanion_AnnounceConfirm("empty")
        return false
    }
    try sendFn.Call()
    catch {
    }
    return AiCompanion_TrackAfterEnter(hwnd, companionId, sentText, snapStatus = "ok")
}

; One in-flight Enter confirm. A newer Enter replaces it. Not a background poll.
global g_AiCompanionEnterWatchCb := ""

; Hotkey path: the chord goes out before any composer or Send-button lookup.
AiCompanion_SendEnterFirst(hwnd, companionId, ctrlEnter := false) {
    global g_AiCompanionEnterWatchCb
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    companionId := StrLower(Trim(companionId))
    if (g_AiCompanionEnterWatchCb != "") {
        try SetTimer(g_AiCompanionEnterWatchCb, 0)
        g_AiCompanionEnterWatchCb := ""
    }
    SendInput(ctrlEnter ? "^{Enter}" : "{Enter}")
    state := { hwnd: hwnd, companionId: companionId, deadline: A_TickCount + 1000, cb: "" }
    cb := AiCompanion_EnterWatchTick.Bind(state)
    state.cb := cb
    g_AiCompanionEnterWatchCb := cb
    SetTimer(cb, -1)
    return true
}

; One stop lookup, then return, so the script thread is not held for the whole second.
AiCompanion_EnterWatchTick(state) {
    global g_AiCompanionEnterWatchCb
    if (g_AiCompanionEnterWatchCb != state.cb)
        return
    stopWatch := false
    if (!WinExist("ahk_id " state.hwnd))
        stopWatch := true
    else if (AiCompanion_IsGenerating(state.hwnd, state.companionId)) {
        stopWatch := true
        AiCompanion_AnnounceConfirm("working")
        AiCompanion_FinishConfirmedSubmit(state.hwnd, state.companionId, state.hwnd)
    } else if (A_TickCount >= state.deadline)
        stopWatch := true
    if (stopWatch) {
        try SetTimer(state.cb, 0)
        if (g_AiCompanionEnterWatchCb = state.cb)
            g_AiCompanionEnterWatchCb := ""
        return
    }
    if (g_AiCompanionEnterWatchCb != state.cb)
        return
    SetTimer(state.cb, -200)
}

PlayAiWorkingStateSound(isWorking) {
    try {
        if (isWorking)
            ScriptSoundPlay(A_ScriptDir . "\assets\sounds\robots-are-working.wav")
        else
            ScriptSoundPlay(A_ScriptDir . "\assets\sounds\no-robot-working.wav")
    } catch {
    }
}

; =============================================================================
; U macro: Global AI generation state (Cursor + Gemini) with sound and banner
; =============================================================================
; Runs Cursor + Gemini stop-button checks, plays robots-are-working / no-robot-working,
; shows red banner when any AI is working, green when none.
; =============================================================================
Cursor_FindComposerIconAcrossInstances() {
    global BANNER_ACCENT_SUCCESS, BANNER_ACCENT_ERROR
    try {
        isWorking := IsAnyAiGenerating()
        PlayAiWorkingStateSound(isWorking)
        if (isWorking)
            ShowCenteredOverlay_Utils("AI is working (stop button found)", 2000, BANNER_ACCENT_ERROR)
        else
            ShowCenteredOverlay_Utils("No AI is working", 2000, BANNER_ACCENT_SUCCESS)
    } catch Error as e {
        PlayAiWorkingStateSound(false)
        ShowCenteredOverlay_Utils("No AI is working", 2000, BANNER_ACCENT_SUCCESS)
    }
}

; Initialize macros
InitMacros() {
    ; Add specific word to Handy
    RegisterMacro(AddWordToHandy, "➕ Add specific word to Handy")
    ; Show Handy for manual edits, or re-hide into background suppress after model switches.
    RegisterMacro(Handy_ToggleVisible, "👁 Toggle Handy visible / hidden", "h")
    ; Email note: new mail to both inboxes (work Outlook / personal Gmail)
    RegisterMacro(EmailNote_Create, "📧 Email note (both inboxes)", "o")
    RegisterMacro(UnescapeMarkdownClipboard, "📋 Unescape markdown clipboard", "e")
    ; Toggle Sound
    RegisterMacro(ToggleSoundState, "🔊 Toggle AutoHotkey mute (volume mixer)")
}

InitMacros()

; ------------
; Optional: scope Explorer-only hotstrings used for renaming
; Uncomment to restrict selected triggers to File Explorer or Save dialogs
;------------
;#HotIf WinActive("ahk_exe explorer.exe") || WinActive("ahk_class #32770")
;:o:gdash::
;    InsertText("GS_E&S_CIP Dashboard research and design")
;return
;#HotIf

; --- Hotkeys & Functions -----------------------------------------------------

; Ensure per-monitor DPI awareness so coordinates are physical pixels across mixed scaling
InitDpiAwareness() {
    static PER_MONITOR_AWARE_V2 := -4 ; DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2
    try DllCall("SetProcessDpiAwarenessContext", "ptr", PER_MONITOR_AWARE_V2, "ptr")
}

InitDpiAwareness()

; Auto-execute: chime after QuickUpdateScripts relaunches AppLaunchers.ahk with "/Updated".
; The success overlay is shown later, after the Tasks server is restarted (Utils.ahk).
if (A_Args.Length > 0 && A_Args[1] = "/Updated") {
    try {
        soundPath := A_ScriptDir "\assets\sounds\quick-update-success.wav"
        ; Play success chime to completion before scheduling volume: async SoundPlay can register a new session after
        ; the first Apply pass, leaving that session at a low default (~10% in the mixer) until something re-enumerates.
        try {
            if (FileExist(soundPath))
                ScriptSoundPlay(soundPath, true)
        } catch {
        }
        ; Volume is applied here, while the other scripts are starting. The success overlay waits until the Tasks server has been restarted.
        ScheduleApplyScriptMasterVolumeTargetAfterQuickUpdate()
    } catch {
    }
}
