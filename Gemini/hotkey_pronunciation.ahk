; =============================================================================
; Gemini module: hotkey_pronunciation.ahk
; Language picker and #!+8 pronunciation hotkey
; Extracted verbatim from Gemini.ahk; loaded via #include into the
; Gemini.ahk process, which remains the entry point / source of truth.
; =============================================================================

; =============================================================================
; Get Pronunciation — language picker + lookup (invoked off the #!+8 hotkey thread via timer)
; =============================================================================
GeminiHotkey_ShowPronunciationLanguagePicker(selectedText) {
    onSelect(lang) {
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        if (lang != "")
            (GeminiAsyncLookup(lang, selectedText)).Start()
    }

    onTimeout() {
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        StandardLoadingBar_Show("⏳ Detecting language…", BANNER_ACCENT_INTERMEDIATE, { textWidth: 450, fontSize: 17 })
        ; Do not call GeminiIPC_DetectLang here: EnsureReady/HealthCheck use synchronous pipe I/O (WriteFile/Connect)
        ; that can block the main thread indefinitely; UI then freezes on this banner. Heuristic is same-thread-safe.
        lang := DetectLang_AhkFallback(selectedText)
        if !(lang = "pt" || lang = "en" || lang = "de")
            lang := "en"
        (GeminiAsyncLookup(lang, selectedText)).Start()
    }

    keyCallbacks := Map(
        "1", (*) => onSelect("pt"),
        "2", (*) => onSelect("en"),
        "3", (*) => onSelect("de"),
        "*Escape", (*) => onSelect("")
    )

    StandardLoadingBar_ShowWithKeys("❓ Auto-detect in 2s — press to override", keyCallbacks, 2000, 0, onTimeout,
        BANNER_ACCENT_INTERMEDIATE,
        450, 17, "", false, "[1] Portuguese  [2] English  [3] German  [Esc] Cancel", false, true)
}

; =============================================================================
; Get Pronunciation
; Hotkey: Win+Alt+Shift+8 — async: submit to Gemini, restore focus, show result in banner when ready
; =============================================================================
#!+8:: {
    global g_StandardLoadingBarIsKeysOverlay
    if (g_StandardLoadingBarIsKeysOverlay) {
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        return
    }

    A_Clipboard := ""
    if (!TryCopySelectionToClipboard_QuickLookAware()) {
        ShowCenteredOverlay_Utils("❌ Copy failed (no selection).", 1800, BANNER_ACCENT_ERROR)
        return
    }
    selectedText := A_Clipboard

    ; Run picker after hotkey returns — StandardLoadingBar_* registers/unregisters global hotkeys; doing that from
    ; inside the same #!+8 thread can deadlock so the banner never appears (no after_show_with_keys_returned in logs).
    ; Do NOT schedule GeminiIPC_EnsureReady here before the picker: timers run FIFO on the main thread; a blocking
    ; HealthCheck/SendRequest can delay or starve the picker timer so the banner never appears.
    ; Daemon warm-up happens inside GeminiIPC_DetectLang → EnsureReady on timeout path.
    companion := ResolveGlobalAICompanion()
    if (companion = "enterprise") {
        GeminiEnterprise_FocusPromptOnly()
        return
    }
    if (companion = "copilot")
        SetTimer(CopilotHotkey_ShowPronunciationLanguagePicker.Bind(selectedText), -1)
    else
        SetTimer(GeminiHotkey_ShowPronunciationLanguagePicker.Bind(selectedText), -1)
}
