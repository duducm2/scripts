; =============================================================================
; Utils module: d2c_flow_manager.ahk
; D2C_FlowManager dictation-Gemini-Cursor state machine
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; D2C_FlowManager: Unified state machine for Dictation → Gemini → Cursor flow.
; Replaces legacy fragmented functions with a central authority to prevent race conditions.
; =============================================================================
class D2C_FlowManager {
    static _instance := 0

    static GetInstance() {
        if (!D2C_FlowManager._instance)
            D2C_FlowManager._instance := D2C_FlowManager()
        return D2C_FlowManager._instance
    }

    __New() {
        this.Reset()
    }

    Reset() {
        this.CurrentPhase := "Idle"
        this.OriginHwnd := 0
        this.GeminiHwnd := 0
        this.CursorHwnd := 0
        this.CompanionId := ""
        this.MonitorTimer := ""
        this.MonitorRetryCount := 0
        this.MonitorMaxRetries := 300 ; 150s
        this.MonitorButtonEverFound := false
        this.MonitorLastCheckTick := 0
        this.HasCopiedForThisResponse := false
    }

    ; --- Entry Points ---

    StartFromDictation() {
        global g_D2C_DictationSubmitMenuCycleFinished
        if (g_D2C_DictationSubmitMenuCycleFinished)
            return
        if (this.CurrentPhase = "PromptingSubmit")
            return
        if (this.CurrentPhase != "Idle")
            return
        this.Reset()
        this.OriginHwnd := 0
        try this.OriginHwnd := WinGetID("A")
        this.PromptForGeminiSubmit()
    }

    StartFromHotstring() {
        if (this.CurrentPhase != "Idle") {
            return
        }
        this.Reset()
        this.OriginHwnd := WinActive("A")
        this.ExecuteGeminiSubmit(true, "", true)
    }

    ; --- Phase 1: Submit Prompt ---

    PromptForGeminiSubmit() {
        this.CurrentPhase := "PromptingSubmit"
        keyCallbacks := Map(
            "G", this.OnSubmitG.Bind(this),
            "A", this.OnSubmitA.Bind(this),
            "T", this.OnSubmitT.Bind(this),
            "Y", this.OnSubmitY.Bind(this),
            "S", this.OnSubmitS.Bind(this),
            "V", this.OnSubmitV.Bind(this),
            "W", this.OnSubmitW.Bind(this),
            "E", this.OnSubmitE.Bind(this),
            "F", this.OnSubmitF.Bind(this),
            "O", this.OnSubmitO.Bind(this),
            "M", this.OnSubmitM.Bind(this),
            "N", this.OnSubmitN.Bind(this)
        )
        StandardLoadingBar_ShowWithKeys(
            "❓ Send dictation? (5s)",
            keyCallbacks,
            D2C_SUBMIT_MENU_TIMEOUT_MS,
            0,
            this.OnSubmitTimeout.Bind(this),
            BANNER_ACCENT_INTERMEDIATE, 720, 17, "", true,
            "[G] Grammar  [A] AI opt  [T] Tasks  [Y] Send  [S] Paste only  [V] Paste dictated  [W] Paste to window  [E] Paste & send  [F] Favorite  [O] Clip Angel  [M] Teams  [N] Cancel",
            true,
            true,
            true
        )
    }

    ; OriginHwnd is captured when the submit menu opens, before the keys overlay activates.
    ActivateOriginForPaste() {
        if (!this.OriginHwnd || !WinExist("ahk_id " this.OriginHwnd))
            return
        WinActivate("ahk_id " this.OriginHwnd)
        if (!WinActive("ahk_id " this.OriginHwnd))
            WinWaitActive("ahk_id " this.OriginHwnd, , 0.3)
    }

    OnSubmitG(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return
        this.ExecuteGeminiSubmit(true, "grammar")
    }

    OnSubmitA(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return
        this.ExecuteGeminiSubmit(true, "aiopt")
    }

    OnSubmitT(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return
        this.ExecuteGeminiSubmit(true, "mtask")
    }

    OnSubmitY(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return
        this.ExecuteGeminiSubmit(true)
    }

    OnSubmitS(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return
        this.ExecuteGeminiSubmit(false)
    }

    OnSubmitV(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return
        this.PasteDictationToActiveWindow()
    }

    ; Shared by menu [W] and #!+L global hotkey.
    ; Pick a visible window and paste the OS clipboard (^v). Does not touch Clip Angel.
    ; If exe+title has a saved main field (paste_field_mappings.ini), focus it first.
    ; If unmapped, after paste prompt Y/N to learn/persist the focused field.
    PasteClipboardToVisibleWindow(originHwnd := 0) {
        if (!originHwnd)
            try originHwnd := WinGetID("A")
        targetHwnd := Dictation_ShowVisiblePasteSelector(originHwnd)
        if (!targetHwnd || !WinExist("ahk_id " targetHwnd))
            return false
        try WinActivate("ahk_id " targetHwnd)
        if (!WinActive("ahk_id " targetHwnd))
            WinWaitActive("ahk_id " targetHwnd, , 0.3)
        Sleep 60
        hadMapping := PasteField_TryFocusMappedField(targetHwnd)
        Send "^v"
        Sleep 80
        if (!hadMapping)
            PasteField_PromptSaveMainField(targetHwnd)
        return true
    }

    ; [W] Pick a visible window, paste OS clipboard (^v), end flow.
    OnSubmitW(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return
        this.CurrentPhase := "PickingVisiblePaste"
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        HideDictationIndicator()
        if (this.PasteClipboardToVisibleWindow(this.OriginHwnd)) {
            global g_D2C_DictationSubmitMenuCycleFinished
            g_D2C_DictationSubmitMenuCycleFinished := true
            this.Reset()
            return
        }
        ; Picker cancelled — restore submit menu; keep OriginHwnd.
        this.PromptForGeminiSubmit()
    }

    ; Paste dictated text into the active window and end the D2C flow (menu [V]).
    PasteDictationToActiveWindow() {
        targetHwnd := 0
        try targetHwnd := WinGetID("A")

        this.CurrentPhase := "PastingDictation"
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        HideDictationIndicator()

        if (targetHwnd && WinExist("ahk_id " targetHwnd)) {
            try WinActivate("ahk_id " targetHwnd)
            if (!WinActive("ahk_id " targetHwnd))
                WinWaitActive("ahk_id " targetHwnd, , 0.2)
        }
        Sleep 60
        Send("^v")

        global g_D2C_DictationSubmitMenuCycleFinished
        g_D2C_DictationSubmitMenuCycleFinished := true
        this.Reset()
    }

    ; Paste at the field that was focused when E was pressed, then Enter (no Gemini).
    OnSubmitE(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return

        ; Lock target at keypress time: last text field the user selected before pressing E.
        targetHwnd := 0
        try targetHwnd := WinGetID("A")

        this.CurrentPhase := "PastingSendHere"
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        HideDictationIndicator()

        if (targetHwnd && WinExist("ahk_id " targetHwnd)) {
            try WinActivate("ahk_id " targetHwnd)
            if (!WinActive("ahk_id " targetHwnd))
                WinWaitActive("ahk_id " targetHwnd, , 0.2)
        }
        Sleep 60
        Send("^v")
        Sleep 150
        Send("{Enter}")

        global g_D2C_DictationSubmitMenuCycleFinished
        g_D2C_DictationSubmitMenuCycleFinished := true
        this.Reset()
    }

    ; [F] Mark newest Clip Angel clip (dictation) as favorite; no Gemini.
    OnSubmitF(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        HideDictationIndicator()
        try {
            clipRaw := A_Clipboard
            clip := Trim(clipRaw)
            if (clip = "" || StrLen(clip) < 10) {
                ShowCenteredOverlay_Utils("❌ Nothing to favorite - clipboard empty or too short", 2000,
                    BANNER_ACCENT_ERROR)
                if (this.OriginHwnd && WinExist("ahk_id " this.OriginHwnd))
                    WinActivate("ahk_id " this.OriginHwnd)
                return
            }
            MarkLastClipAsFavorite("first", true)
        } finally {
            global g_D2C_DictationSubmitMenuCycleFinished
            g_D2C_DictationSubmitMenuCycleFinished := true
            this.Reset()
        }
    }

    ; [O] Open Clip Angel (Row 0), Shift+P (native: leave favorites / all-clips view), then Edit text (F4 — same as Shift+E in Shift keys.ahk for ClipAngel). O avoids C = Transfer on Copy response? banner.
    OnSubmitO(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return
        StandardLoadingBar_CloseKeysOverlay()
        HideDictationIndicator()

        if !ClipAngel_TryAcquireAutomationLock() {
            global g_D2C_DictationSubmitMenuCycleFinished
            g_D2C_DictationSubmitMenuCycleFinished := true
            this.Reset()
            return
        }

        StandardLoadingBar_Show("⏳ Clip Angel: opening...", BANNER_ACCENT_INTERMEDIATE)
        try {
            ; Use the origin window (what the user was looking at) to decide the target monitor for ClipAngel.
            originHwnd := this.OriginHwnd
            if (!originHwnd)
                try originHwnd := WinGetID("A")
            originMon := GetAhkMonitorIndexFromHwnd(originHwnd)

            ActivateClipAngelWithFocusCorrection(true, originMon, true)
            clipHwnd := ClipAngel_MainHwnd()
            if (!clipHwnd) {
                StandardLoadingBar_Update("❌ Clip Angel: window not found", BANNER_ACCENT_ERROR)
                return
            }

            if (!WinWaitActive("ahk_id " clipHwnd, , 0.6)) {
                StandardLoadingBar_Update("❌ Clip Angel: failed to activate", BANNER_ACCENT_ERROR)
                return
            }

            StandardLoadingBar_Update("⏳ Clip Angel: opening editor...", BANNER_ACCENT_INTERMEDIATE)
            ClipAngel_LeaveFavoritesFilter(clipHwnd)
            priorSendLevel := A_SendLevel
            SendLevel 0
            SendInput "{Tab}"
            Sleep 40
            SendInput "^a"
            Sleep 40
            SendInput "^c"
            try ClipWait(0.3)
            catch {
            }
            SendInput "{F10}"
            ClipAngel_UiaWaitPreviewFocused(clipHwnd, 150)
            SendInput "{Up}"
            Sleep 40
            SendInput "{F4}"
            SendLevel priorSendLevel

            StandardLoadingBar_Update("⏳ Clip Angel: maximizing...", BANNER_ACCENT_INTERMEDIATE)
            TryMaximizeWindow(clipHwnd)
            StandardLoadingBar_Update("✅ Clip Angel: ready", BANNER_ACCENT_SUCCESS)
        } finally {
            StandardLoadingBar_Hide(350)
            ClipAngel_ReleaseAutomationLock()
            global g_D2C_DictationSubmitMenuCycleFinished
            g_D2C_DictationSubmitMenuCycleFinished := true
            this.Reset()
        }
    }

    OnSubmitN(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return
        this.CancelFlow(GetGlobalAIProviderLabel() . " submission cancelled")
    }

    ; [M] Prompt for Teams contact, jump to chat, paste dictated text (no Enter).
    OnSubmitM(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return

        this.CurrentPhase := "PastingToTeams"
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        HideDictationIndicator()

        messageText := A_Clipboard
        clipSaved := ClipboardAll()

        try {
            ib := InputBox("Enter a Teams contact name:", "Jump to Chat")
            contact := Trim(ib.Value)
            if (ib.Result = "Cancel" || contact = "") {
                return
            }

            if (!TeamsJumpToChat(contact)) {
                return
            }

            A_Clipboard := ""
            A_Clipboard := messageText
            if (!ClipWait(2)) {
                ShowCenteredOverlay_Utils("❌ CLIPBOARD ERROR - COULD NOT RESTORE DICTATION", 3000, BANNER_ACCENT_ERROR)
                return
            }

            Sleep 100
            Send "^v"
        } finally {
            A_Clipboard := clipSaved
            if (ClipWait(1)) {
            }
            global g_D2C_DictationSubmitMenuCycleFinished
            g_D2C_DictationSubmitMenuCycleFinished := true
            this.Reset()
        }
    }

    OnSubmitTimeout(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return
        this.CancelFlow(GetGlobalAIProviderLabel() . " submission cancelled")
    }

    ; --- Phase 2: Submit Execute ---

    ; presetMode: "" = Clip Angel first snippet; "grammar" | "aiopt" | "mtask" = preset from assets/prompt/*.txt + clipboard dictation via InsertText.
    ; showPreMovementWarning: true only for non-banner-triggered submits (e.g., hotstring path).
    ExecuteGeminiSubmit(autoSubmit := true, presetMode := "", showPreMovementWarning := false) {
        this.CurrentPhase := "Submitting"
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        HideDictationIndicator()

        aiLabel := GetGlobalAIProviderLabel()
        ; For explicit first-banner choices (Y/G/A/T/S), skip the handoff cue: user intentionally chose the AI target.
        if (showPreMovementWarning)
            PlayPreMovementWarning(aiLabel)

        optionalSnippet := ""
        if (presetMode = "grammar" || presetMode = "aiopt" || presetMode = "mtask") {
            dictation := ""
            try dictation := A_Clipboard
            if (presetMode = "grammar")
                preset := GetGrammarPromptText()
            else if (presetMode = "aiopt")
                preset := GetAioptPromptText()
            else
                preset := GetMtaskPromptText()
            optionalSnippet := D2C_CombinePresetWithDictation(preset, dictation)
        }

        this.CompanionId := ResolveGlobalAICompanion()
        if (this.CompanionId = "enterprise") {
            this.GeminiHwnd := GeminiEnterprise_NavigateFocusAndPaste(optionalSnippet, false)
            if (!this.GeminiHwnd)
                this.GeminiHwnd := GetGeminiEnterpriseWindowHwnd()
        } else if (this.CompanionId = "copilot") {
            this.GeminiHwnd := CopilotWeb_NavigateFocusAndPaste(optionalSnippet, false)
            if (!this.GeminiHwnd)
                this.GeminiHwnd := GetCopilotWebWindowHwnd()
        } else {
            if (optionalSnippet != "")
                geminiHwnd := GeminiNavigateFocusAndPasteFirstSnippet(optionalSnippet, false)
            else
                geminiHwnd := GeminiNavigateFocusAndPasteFirstSnippet("", false)
            this.GeminiHwnd := geminiHwnd ? geminiHwnd : WinExist("A")
        }

        if (autoSubmit) {
            if (this.CompanionId = "enterprise") {
                Sleep 1000
                endTick := A_TickCount + 5000
                while (A_TickCount < endTick) {
                    if (GeminiEnterprise_ComposerGetText(this.GeminiHwnd) != "")
                        break
                    Sleep 200
                }
                try {
                    uia := UIA_Browser("ahk_id " this.GeminiHwnd)
                    GeminiEnterprise_TrySubmit(uia)
                } catch {
                    Send("{Enter}")
                }
            } else if (this.CompanionId = "copilot") {
                Sleep 1000 ; Pre-enter delay
                endTick := A_TickCount + 5000
                while (A_TickCount < endTick) {
                    if (CopilotWeb_ComposerGetText(this.GeminiHwnd) != "")
                        break
                    Sleep 200
                }
                try {
                    uia := UIA_Browser("ahk_id " this.GeminiHwnd)
                    CopilotWeb_TrySubmit(uia)
                } catch {
                    Send("{Enter}")
                }
            } else
                Gemini_WaitForPromptContentAndSubmit(this.GeminiHwnd)
            this.StartGeminiMonitor()
        }

        if (this.OriginHwnd && WinExist("ahk_id " this.OriginHwnd))
            WinActivate("ahk_id " this.OriginHwnd)

        if (!autoSubmit)
            this.Reset()
    }

    ; --- Phase 3: Monitor ---

    StartGeminiMonitor() {
        this.CurrentPhase := "Monitoring"
        this.MonitorRetryCount := 0
        this.MonitorButtonEverFound := false
        this.MonitorTimer := this.CheckGeminiCompletion.Bind(this)
        SetTimer(this.MonitorTimer, 500)
    }

    CheckGeminiCompletion() {
        delta := this.MonitorLastCheckTick ? (A_TickCount - this.MonitorLastCheckTick) : -1
        this.MonitorLastCheckTick := A_TickCount
        if (this.CurrentPhase != "Monitoring") {
            SetTimer(this.MonitorTimer, 0)
            return
        }

        this.MonitorRetryCount++
        if (this.MonitorRetryCount > this.MonitorMaxRetries) {
            SetTimer(this.MonitorTimer, 0)
            this.Reset()
            return
        }

        btn := ""
        companion := this.CompanionId != "" ? this.CompanionId : ResolveGlobalAICompanion()
        buttonNames := ["Stop streaming", "Interromper transmissão", "Stop response"]
        root := 0
        try {
            if (companion = "copilot") {
                root := CopilotWeb_ReadRootFromHwnd(this.GeminiHwnd)
                if (root)
                    btn := CopilotWeb_FindStopGenerating(root)
            } else if (companion = "enterprise") {
                root := GeminiEnterprise_ReadRootFromHwnd(this.GeminiHwnd)
                if (root)
                    btn := GeminiEnterprise_FindStopButton(root)
            } else {
                root := UIA.ElementFromHandle(this.GeminiHwnd)
                for n in buttonNames {
                    try {
                        btn := root.FindElement({ Name: n, Type: "Button" })
                    } catch {
                        btn := ""
                    }
                    if (btn)
                        break
                }
            }
        } catch {
            return
        }

        if (btn) {
            this.MonitorButtonEverFound := true
            return
        }

        if (this.MonitorButtonEverFound) {
            ; Suspend timer to prevent re-entrancy during the 800ms Sleep block
            SetTimer(this.MonitorTimer, 0)

            isTrulyGone := true
            loop 4 {
                Sleep 200
                try {
                    if (companion = "copilot") {
                        copRoot := CopilotWeb_ReadRootFromHwnd(this.GeminiHwnd)
                        if (copRoot && CopilotWeb_FindStopGenerating(copRoot))
                            isTrulyGone := false
                    } else if (companion = "enterprise") {
                        geRoot := GeminiEnterprise_ReadRootFromHwnd(this.GeminiHwnd)
                        if (geRoot && GeminiEnterprise_FindStopButton(geRoot))
                            isTrulyGone := false
                    } else {
                        for n in buttonNames {
                            if root.ElementExist({ Name: n, Type: "Button" }) {
                                isTrulyGone := false
                                break
                            }
                        }
                    }
                } catch {
                    isTrulyGone := true
                }
                if (!isTrulyGone)
                    break
            }

            if (isTrulyGone) {
                ; Timer is already stopped, proceed to next phase
                try ScriptSoundPlay(A_ScriptDir . "\assets\sounds\gemini-completion.wav")
                catch {
                    ; Ignore chime failures
                }
                this.PromptForResponseAction()
            } else {
                ; False alarm, the button is still there. Resume polling.
                SetTimer(this.MonitorTimer, 500)
            }
        }
    }

    ; --- Phase 4: Action Prompt ---

    PromptForResponseAction() {
        ; Prevent duplicate banner spawns from timer re-entrancy
        if (this.CurrentPhase = "PromptingAction") {
            return
        }
        this.CurrentPhase := "PromptingAction"
        keyCallbacks := Map(
            "Y", this.OnActionY.Bind(this),
            "C", this.OnActionC.Bind(this),
            "R", this.OnActionR.Bind(this),
            "N", this.OnActionN.Bind(this),
            "F", this.OnActionF.Bind(this)
        )
        pk := "[Y] Copy  [N] No  [R] Copy+Read  [C] Transfer  [F] Copy+Favorite"
        StandardLoadingBar_ShowWithKeys(
            "❓ Copy response? (5s)",
            keyCallbacks,
            D2C_SUBMIT_MENU_TIMEOUT_MS,
            0,
            this.OnActionTimeout.Bind(this),
            BANNER_ACCENT_INTERMEDIATE, 520, 17, "", true,
            pk,
            true,
            true
        )
    }

    OnActionY(*) {
        if (this.CurrentPhase != "PromptingAction")
            return
        this.ExecuteAction(false, false)
    }

    OnActionC(*) {
        if (this.CurrentPhase != "PromptingAction")
            return
        this.PromptForCursorTransfer()
    }

    OnActionR(*) {
        if (this.CurrentPhase != "PromptingAction")
            return
        this.ExecuteAction(true, false)
    }

    ; F: copy last Gemini reply, then mark newest Clip Angel clip as favorite (same as Gemini.ahk CopyAndFavorite).
    OnActionF(*) {
        if (this.CurrentPhase != "PromptingAction")
            return
        this.CleanupActionPrompt()
        try {
            this.DoCopyCore(false, false)
            clipRaw := A_Clipboard
            clip := Trim(clipRaw)
            if (clip = "" || StrLen(clip) < 10) {
                ShowCenteredOverlay_Utils("❌ Copy failed or empty - try again", 2000, BANNER_ACCENT_ERROR)
                if (this.OriginHwnd && WinExist("ahk_id " this.OriginHwnd))
                    WinActivate("ahk_id " this.OriginHwnd)
                return
            }
            MarkLastClipAsFavorite("first", true)
        } finally {
            this.Reset()
        }
    }

    OnActionN(*) {
        if (this.CurrentPhase != "PromptingAction")
            return
        this.CleanupActionPrompt()
        this.Reset()
    }

    OnActionTimeout(*) {
        if (this.CurrentPhase != "PromptingAction")
            return
        this.CleanupActionPrompt()
        this.Reset()
    }

    CleanupActionPrompt() {
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
    }

    ExecuteAction(readAloud := false, skipRestoreFocus := false) {
        this.CleanupActionPrompt()
        try {
            this.DoCopyCore(readAloud, skipRestoreFocus)
        } finally {
            this.Reset()
        }
    }

    DoCopyCore(readAloud := false, skipRestoreFocus := false) {
        if (this.HasCopiedForThisResponse) {
            return
        }
        this.HasCopiedForThisResponse := true

        aiLabel := GetGlobalAIProviderLabel()
        companion := this.CompanionId != "" ? this.CompanionId : ResolveGlobalAICompanion()
        PlayPreMovementWarning(aiLabel)

        if (!WinExist("ahk_id " this.GeminiHwnd)) {
            if (this.OriginHwnd && WinExist("ahk_id " this.OriginHwnd))
                WinActivate("ahk_id " this.OriginHwnd)
            return
        }

        ; By the time DoCopyCore runs, Gemini should already be active. If it is not,
        ; just activate it without a pre-movement warning (source is no longer Original).
        if (!WinActive("ahk_id " this.GeminiHwnd)) {
            try WinActivate("ahk_id " this.GeminiHwnd)
            catch {
                if (this.OriginHwnd && WinExist("ahk_id " this.OriginHwnd))
                    WinActivate("ahk_id " this.OriginHwnd)
                return
            }
            if (!WinWaitActive("ahk_exe chrome.exe", , 0.5)) {
                if (this.OriginHwnd && WinExist("ahk_id " this.OriginHwnd))
                    WinActivate("ahk_id " this.OriginHwnd)
                return
            }
        }

        ; Enterprise: copy/read-aloud not transposed yet — focus prompt only.
        if (companion = "enterprise") {
            try {
                root := GeminiEnterprise_ReadRootFromHwnd(this.GeminiHwnd)
                if (IsObject(root))
                    GeminiEnterprise_FocusComposer(root, true)
            } catch {
            }
            if (!skipRestoreFocus && this.OriginHwnd && WinExist("ahk_id " this.OriginHwnd) && !WinActive("ahk_id " this
                .OriginHwnd
            )) {
                WinActivate("ahk_id " this.OriginHwnd)
                if (!WinActive("ahk_id " this.OriginHwnd))
                    WinWaitActive("ahk_id " this.OriginHwnd, , 0.5)
            }
            return
        }

        ; Y / R / C / timeout: same synchronous copy first. R then blocks on read-aloud IPC (wParam=1 skips duplicate Copy in Gemini).
        clipBefore := A_Clipboard
        seqBefore := Clipboard_GetSequenceNumber()
        WM_COPY_LAST_GEMINI := 0x8001
        WM_COPY_LAST_COPILOT := 0x8005
        WM_TRIGGER_READ_ALOUD := 0x8004
        WM_TRIGGER_COPILOT_READ_ALOUD := 0x8006
        useCopilot := (companion = "copilot")
        wmCopy := useCopilot ? WM_COPY_LAST_COPILOT : WM_COPY_LAST_GEMINI
        wmRead := useCopilot ? WM_TRIGGER_COPILOT_READ_ALOUD : WM_TRIGGER_READ_ALOUD
        targetHwnd := GetGeminiScriptMsgTargetHwnd()
        sendOk := false
        clipOk := false

        if (targetHwnd) {
            prevDH := A_DetectHiddenWindows
            DetectHiddenWindows true
            try {
                SendMessage(wmCopy, 0, this.GeminiHwnd, , "ahk_id " targetHwnd, , , , 20000)
                sendOk := true
            } catch {
            } finally {
                DetectHiddenWindows prevDH
            }
            changed := Clipboard_WaitForSequenceChange(seqBefore, 2000, 850)
            clipOk := sendOk && changed && A_Clipboard != clipBefore && Trim(A_Clipboard) != ""
            if (!sendOk)
                ShowCenteredOverlay_Utils("❌ " . aiLabel . " copy timed out or IPC failed", 3500, BANNER_ACCENT_ERROR)
            else if (!clipOk)
                ShowCenteredOverlay_Utils("❌ Copy failed or clipboard empty - try again", 3000, BANNER_ACCENT_ERROR)
            else if (readAloud) {
                DetectHiddenWindows true
                try {
                    SendMessage(wmRead, 1, this.OriginHwnd, , "ahk_id " targetHwnd, , , , 120000)
                } catch {
                    ShowCenteredOverlay_Utils("❌ Read aloud failed or timed out", 4000, BANNER_ACCENT_ERROR)
                } finally {
                    DetectHiddenWindows prevDH
                }
            } else
                try ScriptSoundPlay(A_ScriptDir . "\assets\sounds\copy.wav")
        } else {
            ShowCenteredOverlay_Utils("❌ Gemini.ahk not running", 2000, BANNER_ACCENT_ERROR)
        }

        ; Gemini/Clipboard → Original: return transitions are immediate (no warning).
        if (!skipRestoreFocus && this.OriginHwnd && WinExist("ahk_id " this.OriginHwnd) && !WinActive("ahk_id " this.OriginHwnd
        )) {
            WinActivate("ahk_id " this.OriginHwnd)
            ; Fast-path: avoid WinWaitActive if we are already active.
            if (!WinActive("ahk_id " this.OriginHwnd))
                WinWaitActive("ahk_id " this.OriginHwnd, , 0.5)
        }
    }

    ; --- Phase 5: Cursor Transfer ---

    PromptForCursorTransfer() {
        this.CurrentPhase := "Transferring"
        this.CleanupActionPrompt()
        try {
            ; Skip restoring focus so clipboard is not overwritten
            this.DoCopyCore(false, true)

            clipRaw := A_Clipboard
            clip := Trim(clipRaw)
            if (clip = "" || StrLen(clip) < 10) {
                ShowCenteredOverlay_Utils("❌ Copy failed or empty - try again", 2000, BANNER_ACCENT_ERROR)
                if (this.OriginHwnd && WinExist("ahk_id " this.OriginHwnd))
                    WinActivate("ahk_id " this.OriginHwnd)
                return
            }

            ; Restore the pre-handoff anchored window so the user sees the selector/paste
            ; happening in the exact app they were monitoring before the Gemini handoff.
            if (this.OriginHwnd && WinExist("ahk_id " this.OriginHwnd)) {
                try {
                    WinActivate("ahk_id " this.OriginHwnd)
                    ; Fast-path: avoid WinWaitActive if we are already active.
                    if (!WinActive("ahk_id " this.OriginHwnd))
                        WinWaitActive("ahk_id " this.OriginHwnd, , 0.5)
                } catch {
                }
            }
            try A_Clipboard := clipRaw

            tSelectorStart := A_TickCount
            this.CursorHwnd := CursorTransfer_ShowWindowSelector(0)
            tSelectorMs := A_TickCount - tSelectorStart
            if (!this.CursorHwnd) {
                if (this.OriginHwnd && WinExist("ahk_id " this.OriginHwnd))
                    WinActivate("ahk_id " this.OriginHwnd)
                try A_Clipboard := clipRaw
                return
            }

            ; Gemini → Cursor: no pre-movement warning (source is not Original).
            try A_Clipboard := clipRaw
            CursorTransfer_ActivateFocusPaste(this.CursorHwnd, this.OriginHwnd)
        } finally {
            this.Reset()
        }
    }

    ; --- Helpers ---

    CancelFlow(message) {
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        ShowCenteredOverlay_Utils("⚠ " . message, 1500, BANNER_ACCENT_INTERMEDIATE)
        global g_D2C_DictationSubmitMenuCycleFinished
        g_D2C_DictationSubmitMenuCycleFinished := true
        this.Reset()
    }
}
