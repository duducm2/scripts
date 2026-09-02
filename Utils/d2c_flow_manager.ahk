; =============================================================================
; Utils module: d2c_flow_manager.ahk
; D2C_FlowManager dictation-Gemini-Cursor state machine
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; After visible-window pick (#!+L / D2C [W]): Y = paste+Enter, N = paste only, Esc = abort, timeout = paste only.
global g_PasteWindowAutoSendChoice := ""
global g_PasteWindowAutoSendActive := false

PasteWindow_FinishAutoSendWait(choice) {
    global g_PasteWindowAutoSendChoice, g_PasteWindowAutoSendActive
    if (!g_PasteWindowAutoSendActive)
        return
    g_PasteWindowAutoSendChoice := choice
    g_PasteWindowAutoSendActive := false
}

PasteWindow_OnAutoSendY(*) {
    PasteWindow_FinishAutoSendWait("send")
}

PasteWindow_OnAutoSendN(*) {
    PasteWindow_FinishAutoSendWait("paste")
}

PasteWindow_OnAutoSendEsc(*) {
    PasteWindow_FinishAutoSendWait("cancel")
}

PasteWindow_OnAutoSendTimeout(*) {
    PasteWindow_FinishAutoSendWait("paste")
}

; Show banner immediately; block until Y / N / Esc / timeout. Returns "send", "cancel", or "paste".
PasteWindow_ShowAutoSendOptionsAndWait() {
    global g_PasteWindowAutoSendChoice, g_PasteWindowAutoSendActive
    g_PasteWindowAutoSendChoice := ""
    g_PasteWindowAutoSendActive := true
    StandardLoadingBar_CloseKeysOverlay()
    StandardLoadingBar_Hide(0)
    keyCallbacks := Map(
        "Y", PasteWindow_OnAutoSendY,
        "N", PasteWindow_OnAutoSendN,
        "Escape", PasteWindow_OnAutoSendEsc
    )
    StandardLoadingBar_ShowWithKeys(
        "❓ Paste to window? (3s)",
        keyCallbacks,
        3000,
        0,
        PasteWindow_OnAutoSendTimeout,
        BANNER_ACCENT_INTERMEDIATE,
        420,
        17,
        "",
        true,
        "[Y] Send after paste  [N] Paste only  [Esc] Cancel",
        true,
        true,
        true
    )
    while (g_PasteWindowAutoSendActive)
        Sleep 30
    return g_PasteWindowAutoSendChoice
}

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
        try {
            if (this.MonitorTimer)
                SetTimer(this.MonitorTimer, 0)
        } catch {
        }
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
        ; Same-wave duplicates use g_D2C_DictationSubmitMenuCycleFinished.
        ; PromptingSubmit (e.g. post-[B] menu still open) must be closed and replaced
        ; so a new dictation wave can show Send dictation? again.
        if (this.CurrentPhase != "Idle") {
            try StandardLoadingBar_CloseKeysOverlay()
            try StandardLoadingBar_Hide(0)
        }
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
        cancelCb := this.OnSubmitN.Bind(this)
        keyCallbacks := Map(
            "G", this.OnSubmitG.Bind(this),
            "A", this.OnSubmitA.Bind(this),
            "T", this.OnSubmitT.Bind(this),
            "D", this.OnSubmitD.Bind(this),
            "Y", this.OnSubmitY.Bind(this),
            "S", this.OnSubmitS.Bind(this),
            "V", this.OnSubmitV.Bind(this),
            "W", this.OnSubmitW.Bind(this),
            "E", this.OnSubmitE.Bind(this),
            "F", this.OnSubmitF.Bind(this),
            "O", this.OnSubmitO.Bind(this),
            "M", this.OnSubmitM.Bind(this),
            "K", this.OnSubmitK.Bind(this),
            "Z", this.OnSubmitZ.Bind(this),
            "P", this.OnSubmitP.Bind(this),
            "R", this.OnSubmitR.Bind(this),
            "L", this.OnSubmitL.Bind(this),
            "C", this.OnSubmitC.Bind(this),
            "B", this.OnSubmitB.Bind(this),
            "N", cancelCb,
            "Escape", cancelCb
        )
        StandardLoadingBar_ShowWithKeys(
            "❓ Send dictation? (5s)",
            keyCallbacks,
            D2C_SUBMIT_MENU_TIMEOUT_MS,
            0,
            this.OnSubmitTimeout.Bind(this),
            BANNER_ACCENT_INTERMEDIATE, 1000, 17, "", true,
            "[G] Grammar  [A] AI opt  [T] Tasks pack  [D] Finance daily  [Y] Send  [S] Paste only  [V] Paste dictated  [W] Paste to window  [E] Paste & send  [F] Favorite  [O] Clip Angel  [M] Teams to  [K] Teams paste  [Z] WhatsApp  [P] Spotify  [R] Replay  [B] Model+retranscribe  [L] Email note  [C] Chrome  [N]/Esc Cancel",
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
        this.ExecuteGeminiSubmit(true, "task_pack")
    }

    OnSubmitD(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return
        this.ExecuteGeminiSubmit(true, "finance_daily")
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
    ; After pick: banner Y = paste+Enter, N = paste only, Esc = abort, timeout = paste only.
    ; If exe+title/url has a saved main field (paste_field_mappings.ini), focus it first.
    ; If unmapped, after paste prompt Y/N to learn/persist the focused field.
    PasteClipboardToVisibleWindow(originHwnd := 0, onDone := "") {
        if (!originHwnd)
            try originHwnd := WinGetID("A")
        targetHwnd := Dictation_ShowVisiblePasteSelector(originHwnd)
        if (!targetHwnd || !WinExist("ahk_id " targetHwnd))
            return false

        choice := PasteWindow_ShowAutoSendOptionsAndWait()
        if (choice = "cancel" || choice = "")
            return false

        autoSend := (choice = "send")
        SetTimer(this._FinishDeferredPaste.Bind(this, targetHwnd, onDone, autoSend), -50)
        return true
    }

    _FinishDeferredPaste(targetHwnd, onDone, autoSend := false) {
        if (!WinExist("ahk_id " targetHwnd)) {
            if (onDone)
                onDone.Call()
            return
        }

        try WinActivate("ahk_id " targetHwnd)
        if (!WinActive("ahk_id " targetHwnd))
            WinWaitActive("ahk_id " targetHwnd, , 0.3)
        Sleep 60

        mappingResult := PasteField_FocusMappedField(targetHwnd)
        if (mappingResult.hasMapping && !mappingResult.focused) {
            loop 2 {
                try WinActivate("ahk_id " targetHwnd)
                Sleep 40
                mappingResult := PasteField_FocusMappedField(targetHwnd)
                if (mappingResult.focused)
                    break
            }
        }

        try WinActivate("ahk_id " targetHwnd)
        Sleep 40
        Send "^v"
        Sleep 80
        if (autoSend) {
            Send "{Enter}"
            Sleep 80
        }

        if (!mappingResult.hasMapping)
            PasteField_PromptSaveMainField(targetHwnd)

        if (onDone)
            onDone.Call()
    }

    ; [W] Pick a visible window, paste OS clipboard (^v), end flow.
    OnSubmitW(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return
        this.CurrentPhase := "PickingVisiblePaste"
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        HideDictationIndicator()

        onDone := this._OnSubmitWDone.Bind(this)
        if (this.PasteClipboardToVisibleWindow(this.OriginHwnd, onDone)) {
            return
        }
        ; Picker cancelled or post-pick Esc — restore submit menu; keep OriginHwnd.
        this.PromptForGeminiSubmit()
    }

    _OnSubmitWDone() {
        global g_D2C_DictationSubmitMenuCycleFinished
        g_D2C_DictationSubmitMenuCycleFinished := true
        this.Reset()
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

    ; [O] Open Clip Angel (Row 0), then Edit text (F4). Shared with Copy-response [O].
    OnSubmitO(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return
        StandardLoadingBar_CloseKeysOverlay()
        HideDictationIndicator()
        try {
            this._OpenClipAngelEditForOrigin(this.OriginHwnd)
        } finally {
            global g_D2C_DictationSubmitMenuCycleFinished
            g_D2C_DictationSubmitMenuCycleFinished := true
            this.Reset()
        }
    }

    ; Open Clip Angel on newest clip and Edit text (F4) — same path as #!+p [O] / HotkeyCopy_OnClipAngelEdit.
    _OpenClipAngelEditForOrigin(originHwnd := 0) {
        if !ClipAngel_TryAcquireAutomationLock()
            return

        StandardLoadingBar_Show("⏳ Clip Angel: opening...", BANNER_ACCENT_INTERMEDIATE)
        try {
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
        }
    }

    OnSubmitN(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return
        this.DismissSubmitPromptInstantly()
    }

    ; N / Esc: close the Send dictation? menu immediately (no follow-up banner).
    DismissSubmitPromptInstantly() {
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        global g_D2C_DictationSubmitMenuCycleFinished
        g_D2C_DictationSubmitMenuCycleFinished := true
        this.Reset()
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

            hwnd := TeamsJump_ResolveChatHwnd()
            if (hwnd <= 0 || !TeamsJump_ActivateWindowWithRetry(hwnd, 3, 300)) {
                ShowCenteredOverlay_Utils("❌ Could not activate Teams chat.", 2500, BANNER_ACCENT_ERROR)
                return
            }
            ; Focus composer + paste with verify/retry (replaces blind ^v).
            TeamsJump_PasteAndVerify(hwnd, messageText)
        } finally {
            A_Clipboard := clipSaved
            if (ClipWait(1)) {
            }
            global g_D2C_DictationSubmitMenuCycleFinished
            g_D2C_DictationSubmitMenuCycleFinished := true
            this.Reset()
        }
    }

    ; [K] Activate Teams chat, UIA-focus composer, paste dictated text (no contact picker, no Enter).
    OnSubmitK(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return

        this.CurrentPhase := "PastingToTeamsComposer"
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        HideDictationIndicator()

        messageText := A_Clipboard
        try {
            TeamsJump_PasteToComposer(messageText)
        } finally {
            global g_D2C_DictationSubmitMenuCycleFinished
            g_D2C_DictationSubmitMenuCycleFinished := true
            this.Reset()
        }
    }

    ; [Z] Prompt for WhatsApp contact, jump to chat, paste dictated text (no Enter).
    OnSubmitZ(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return

        this.CurrentPhase := "PastingToWhatsApp"
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        HideDictationIndicator()

        messageText := A_Clipboard
        clipSaved := ClipboardAll()

        try {
            ib := InputBox("Enter a WhatsApp contact name:", "Jump to Chat")
            contact := Trim(ib.Value)
            if (ib.Result = "Cancel" || contact = "") {
                return
            }

            ; keepBarVisible so Loading Indication continues through paste.
            if (!WhatsAppJumpToChat(contact, true)) {
                return
            }

            WhatsAppJump_UpdateLoading("⏳ Pasting message...")
            A_Clipboard := ""
            A_Clipboard := messageText
            if (!ClipWait(2)) {
                WhatsAppJump_HideLoading()
                ShowCenteredOverlay_Utils("❌ CLIPBOARD ERROR - COULD NOT RESTORE DICTATION", 3000, BANNER_ACCENT_ERROR)
                return
            }

            Sleep 200
            Send "^v"
            ; Let Ctrl+V finish reading the message before restoring the prior clipboard.
            Sleep 350
            StandardLoadingBar_Update("✅ Message ready in WhatsApp")
            Sleep 400
        } finally {
            WhatsAppJump_HideLoading()
            A_Clipboard := clipSaved
            if (ClipWait(1)) {
            }
            global g_D2C_DictationSubmitMenuCycleFinished
            g_D2C_DictationSubmitMenuCycleFinished := true
            this.Reset()
        }
    }

    ; [R] Open Handy History and play the last recording.
    OnSubmitR(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return

        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        HideDictationIndicator()

        global g_D2C_DictationSubmitMenuCycleFinished
        g_D2C_DictationSubmitMenuCycleFinished := true
        this.Reset()

        HandyReplay_PlayLastRecording()
    }

    ; [B] Toggle Parakeet <-> Cohere, re-transcribe newest History entry, copy, re-open menu.
    OnSubmitB(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return

        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        HideDictationIndicator()

        originHwnd := this.OriginHwnd
        this.CurrentPhase := "Retranscribing"

        try {
            HandyRetranscribe_ToggleModelAndCopy()
        } catch {
            ShowCenteredOverlay_Utils("❌ Model toggle / re-transcribe failed.", 2200, BANNER_ACCENT_ERROR)
        }

        ; Keep origin for paste/send; do not mark cycle finished — menu continues.
        this.OriginHwnd := originHwnd
        this.PromptForGeminiSubmit()
    }

    ; [P] Open Spotify, Ctrl+K search, paste dictation, Enter, then immerse (play + fullscreen).
    OnSubmitP(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return

        this.CurrentPhase := "PlayingOnSpotify"
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        HideDictationIndicator()

        messageText := A_Clipboard

        try {
            SpotifyDictation_PlayFromClipboard(messageText)
        } finally {
            global g_D2C_DictationSubmitMenuCycleFinished
            g_D2C_DictationSubmitMenuCycleFinished := true
            this.Reset()
        }
    }

    ; [C] New Chrome window, paste dictation into address bar, Enter.
    OnSubmitC(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return

        this.CurrentPhase := "OpeningChromeNavigate"
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        HideDictationIndicator()

        messageText := A_Clipboard

        try {
            ChromeDictation_NavigateFromClipboard(messageText)
        } finally {
            global g_D2C_DictationSubmitMenuCycleFinished
            g_D2C_DictationSubmitMenuCycleFinished := true
            this.Reset()
        }
    }

    ; [L] Email note: new mail to both inboxes, dictated text as Subject.
    OnSubmitL(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return

        this.CurrentPhase := "CreatingEmailNote"
        StandardLoadingBar_CloseKeysOverlay()
        messageText := A_Clipboard
        StandardLoadingBar_Hide(0)
        HideDictationIndicator()

        try {
            EmailNote_Create(messageText)
        } finally {
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

    ; presetMode: "" = Clip Angel first snippet; "grammar" | "aiopt" | "task_pack" | "mtask" | "finance_daily"
    ; load Utility Shortcuts prompt by char (Prompt Manager metadata applied via PreparedBodyForSend).
    ; Finance daily / registry presets: attach context files then paste prompt + dictation.
    ; Finance daily sends only (no wait / download / import); user finishes manually.
    ; showPreMovementWarning: true only for non-banner-triggered submits (e.g., hotstring path).
    ExecuteGeminiSubmit(autoSubmit := true, presetMode := "", showPreMovementWarning := false) {
        this.CurrentPhase := "Submitting"
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        HideDictationIndicator()

        aiLabel := GetGlobalAIProviderLabel()
        ; For explicit first-banner choices (Y/G/A/T/D/S), skip the handoff cue: user intentionally chose the AI target.
        if (showPreMovementWarning)
            PlayPreMovementWarning(aiLabel)

        optionalSnippet := ""
        registryPrompt := false
        useRegistryPastePath := false
        dictation := ""
        try dictation := A_Clipboard

        if (presetMode = "finance_daily" || presetMode = "grammar" || presetMode = "aiopt"
            || presetMode = "mtask" || presetMode = "task_pack") {
            PromptData_Load(true)
            charKey := ""
            if (presetMode = "finance_daily")
                charKey := "d"
            else if (presetMode = "grammar")
                charKey := "1"
            else if (presetMode = "aiopt")
                charKey := "3"
            else if (presetMode = "task_pack")
                charKey := "k"
            else if (presetMode = "mtask")
                charKey := "2"
            registryPrompt := PromptData_FindByChar(charKey)
            if (presetMode = "finance_daily" && !IsObject(registryPrompt)) {
                ShowCenteredOverlay_Utils("⚠ Finance daily prompt not found (char d)", 2200, BANNER_ACCENT_ERROR)
                this.Reset()
                return
            }
            if (presetMode = "task_pack" && !IsObject(registryPrompt)) {
                ShowCenteredOverlay_Utils("⚠ Convert to Task pack prompt not found (char k)", 2200, BANNER_ACCENT_ERROR
                )
                this.Reset()
                return
            }
            presetBody := ""
            if (IsObject(registryPrompt)) {
                presetBody := PromptData_PreparedBodyForSend(registryPrompt)
                if (presetBody = "") {
                    this.Reset()
                    return
                }
                useRegistryPastePath := true
            } else if (presetMode = "grammar") {
                presetBody := GetGrammarPromptText()
            } else if (presetMode = "aiopt") {
                presetBody := GetAioptPromptText()
            } else if (presetMode = "mtask") {
                presetBody := GetMtaskPromptText()
            }
            optionalSnippet := D2C_CombinePresetWithDictation(presetBody, dictation)
        }

        this.CompanionId := ResolveGlobalAICompanion()
        if (useRegistryPastePath) {
            resolved := UtilitySelector_ResolveContextEntries(registryPrompt)
            if (resolved = false) {
                this.Reset()
                return
            }
            contextEntries := resolved.entries
            ; Focus only (no ClipAngel paste). Attach Prompt Manager context, then paste once.
            if (this.CompanionId = "enterprise") {
                GeminiEnterprise_OpenOrFocus()
                this.GeminiHwnd := GetGeminiEnterpriseWindowHwnd()
            } else if (this.CompanionId = "copilot") {
                CopilotWeb_OpenOrFocus()
                this.GeminiHwnd := GetCopilotWebWindowHwnd()
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
                    geminiHwnd := WinExist("A")
                }
                try {
                    uia := geminiHwnd ? UIA_Browser("ahk_id " geminiHwnd) : UIA_Browser()
                    Gemini_FocusPromptWithChime(uia, { playChime: false, useAnchorFallback: true })
                } catch {
                }
                this.GeminiHwnd := geminiHwnd ? geminiHwnd : WinExist("A")
            }
            if (!this.GeminiHwnd)
                this.GeminiHwnd := WinExist("A")
            if (presetMode = "finance_daily") {
                try StandardLoadingBar_Show("⏳ Attaching finance context…", BANNER_ACCENT_INTERMEDIATE, {
                    passive: false,
                    centerOnHwnd: this.GeminiHwnd
                })
                catch {
                }
            }
            UtilitySelector_AttachPromptContextFiles(registryPrompt, contextEntries)
            if (optionalSnippet != "")
                InsertText(optionalSnippet)
        } else if (this.CompanionId = "enterprise") {
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
            if (presetMode = "finance_daily") {
                try StandardLoadingBar_Update("⏳ Waiting for context files…", BANNER_ACCENT_INTERMEDIATE)
                catch {
                    try StandardLoadingBar_Show("⏳ Waiting for context files…", BANNER_ACCENT_INTERMEDIATE, {
                        passive: false,
                        centerOnHwnd: this.GeminiHwnd
                    })
                    catch {
                    }
                }
                ready := false
                try ready := PromptContext_WaitForSendReady(this.GeminiHwnd, this.CompanionId, 45000)
                catch {
                }
                try StandardLoadingBar_Hide(0)
                catch {
                }
                if (!ready) {
                    try ShowCenteredOverlay_Utils("⚠ Send not ready — submitting anyway", 2200, BANNER_ACCENT_ERROR)
                    catch {
                    }
                }
            }
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
            if (presetMode = "finance_daily") {
                ; Send-only: leave companion generating; user downloads/imports manually.
                if (this.OriginHwnd && WinExist("ahk_id " this.OriginHwnd))
                    WinActivate("ahk_id " this.OriginHwnd)
                try ShowCenteredOverlay_Utils("✅ Finance daily sent — finish manually", 2200, BANNER_ACCENT_SUCCESS)
                catch {
                }
                global g_D2C_DictationSubmitMenuCycleFinished
                g_D2C_DictationSubmitMenuCycleFinished := true
                this.Reset()
                return
            }
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
        companion := this.CompanionId != "" ? this.CompanionId : ResolveGlobalAICompanion()
        ; Same key strip as #!+p HotkeyCopy_ShowPostCopyBanner, plus [P] Copy-only (#!+p 1× without a destination).
        keyCallbacks := Map(
            "P", this.OnActionP.Bind(this),
            "Y", this.OnActionY.Bind(this),
            "F", this.OnActionF.Bind(this),
            "C", this.OnActionC.Bind(this),
            "W", this.OnActionW.Bind(this),
            "O", this.OnActionO.Bind(this),
            "N", this.OnActionN.Bind(this),
            "Escape", this.OnActionN.Bind(this)
        )
        if (companion != "enterprise")
            keyCallbacks["R"] := this.OnActionR.Bind(this)
        if (companion = "enterprise")
            pk := "[P] Copy  [Y] Desktop  [F] Favorite  [C] Transfer  [W] Paste window  [O] Clip Angel  [N] No"
        else
            pk :=
                "[P] Copy  [Y] Desktop  [F] Favorite  [C] Transfer  [R] Read  [W] Paste window  [O] Clip Angel  [N] No"
        StandardLoadingBar_ShowWithKeys(
            "❓ Response ready — what next? (5s)",
            keyCallbacks,
            D2C_SUBMIT_MENU_TIMEOUT_MS,
            0,
            this.OnActionTimeout.Bind(this),
            BANNER_ACCENT_INTERMEDIATE, 900, 17, "", true,
            pk,
            true,
            true
        )
    }

    ; [Y] Copy reply then export last Clip Angel clip to Desktop (same as #!+p [Y]).
    OnActionY(*) {
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
            Sleep CLIPANGEL_PRE_FAVORITE_INGEST_DELAY_MS
            ClipAngel_ExportLastClipToDesktop()
        } finally {
            this.Reset()
        }
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

    ; [P] Copy reply to clipboard only (same as #!+p 1× without a destination action).
    OnActionP(*) {
        if (this.CurrentPhase != "PromptingAction")
            return
        this.ExecuteAction(false, false)
    }

    ; [F] Copy reply, then mark newest Clip Angel clip as favorite (same as #!+p [F]).
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

    ; [W] Copy reply, then paste to a picked visible window (same as #!+p [W]).
    OnActionW(*) {
        if (this.CurrentPhase != "PromptingAction")
            return
        this.CleanupActionPrompt()
        this.CurrentPhase := "PickingVisiblePaste"
        try {
            this.DoCopyCore(false, true)
            clipRaw := A_Clipboard
            clip := Trim(clipRaw)
            if (clip = "" || StrLen(clip) < 10) {
                ShowCenteredOverlay_Utils("❌ Copy failed or empty - try again", 2000, BANNER_ACCENT_ERROR)
                if (this.OriginHwnd && WinExist("ahk_id " this.OriginHwnd))
                    WinActivate("ahk_id " this.OriginHwnd)
                this.Reset()
                return
            }
            try A_Clipboard := clipRaw
            onDone := this._OnActionWDone.Bind(this)
            if (this.PasteClipboardToVisibleWindow(this.OriginHwnd, onDone))
                return
            if (this.OriginHwnd && WinExist("ahk_id " this.OriginHwnd))
                WinActivate("ahk_id " this.OriginHwnd)
            this.Reset()
        } catch {
            this.Reset()
        }
    }

    _OnActionWDone() {
        this.Reset()
    }

    ; [O] Copy reply, then open Clip Angel Edit (same as #!+p [O]).
    OnActionO(*) {
        if (this.CurrentPhase != "PromptingAction")
            return
        this.CleanupActionPrompt()
        try {
            this.DoCopyCore(false, true)
            clipRaw := A_Clipboard
            clip := Trim(clipRaw)
            if (clip = "" || StrLen(clip) < 10) {
                ShowCenteredOverlay_Utils("❌ Copy failed or empty - try again", 2000, BANNER_ACCENT_ERROR)
                if (this.OriginHwnd && WinExist("ahk_id " this.OriginHwnd))
                    WinActivate("ahk_id " this.OriginHwnd)
                return
            }
            Sleep CLIPANGEL_PRE_FAVORITE_INGEST_DELAY_MS
            this._OpenClipAngelEditForOrigin(this.OriginHwnd)
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

        ; Y/F/C/R/W/O: same synchronous copy first. R then blocks on read-aloud IPC (wParam=1 skips duplicate Copy in Gemini).
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
        if (message != "")
            ShowCenteredOverlay_Utils("⚠ " . message, 1500, BANNER_ACCENT_INTERMEDIATE)
        global g_D2C_DictationSubmitMenuCycleFinished
        g_D2C_DictationSubmitMenuCycleFinished := true
        this.Reset()
    }
}
