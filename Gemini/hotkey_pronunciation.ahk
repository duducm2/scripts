; =============================================================================
; Gemini module: hotkey_pronunciation.ahk
; Tiered #!+8 pronunciation: 1× English, 2× German, hold = language ListView
; Loaded via #include into the Gemini.ahk process (entry point / source of truth).
; =============================================================================

PRONUNCIATION_HOLD_MS := 700
PRONUNCIATION_DOUBLE_TAP_MS := 400

global g_Pronunciation_DoubleTapArmed := false
global g_Pronunciation_LastPressTick := 0
global g_Pronunciation_DoubleTapTimer := 0
global g_PronunciationLangGui := false
global g_PronunciationLangLv := false
global g_PronunciationLangActive := false
global g_PronunciationLangText := ""
global g_PronunciationLangHotkeyHandlers := []
global g_PronunciationLangHotkeysBound := false

; =============================================================================
; Companion-aware lookup start (no picker)
; =============================================================================
PronunciationHotkey_StartLookup(lang, selectedText) {
    if (lang = "" || selectedText = "")
        return
    companion := ResolveGlobalAICompanion()
    if (companion = "enterprise")
        (GeminiEnterpriseAsyncLookup(lang, selectedText)).Start()
    else if (companion = "copilot")
        (CopilotAsyncLookup(lang, selectedText)).Start()
    else
        (GeminiAsyncLookup(lang, selectedText)).Start()
}

PronunciationHotkey_CopySelection() {
    A_Clipboard := ""
    if (!TryCopySelectionToClipboard_QuickLookAware()) {
        ShowCenteredOverlay_Utils("❌ Copy failed (no selection).", 1800, BANNER_ACCENT_ERROR)
        return ""
    }
    return A_Clipboard
}

PronunciationHotkey_CopyAndLookup(lang) {
    selectedText := PronunciationHotkey_CopySelection()
    if (selectedText = "")
        return
    PronunciationHotkey_StartLookup(lang, selectedText)
}

PronunciationHotkey_CopyAndShowPicker() {
    selectedText := PronunciationHotkey_CopySelection()
    if (selectedText = "")
        return
    PronunciationHotkey_ShowLanguagePicker(selectedText)
}

; =============================================================================
; Hold modal — compact Utility Shortcuts / project-selector ListView
; =============================================================================
PronunciationHotkey_ShowLanguagePicker(selectedText) {
    global g_PronunciationLangActive, g_PronunciationLangGui, g_PronunciationLangText
    if (g_PronunciationLangActive && IsObject(g_PronunciationLangGui)) {
        PronunciationHotkey_CleanupLanguagePicker()
        return
    }
    g_PronunciationLangText := selectedText
    PronunciationHotkey_BuildLanguagePickerGui()
}

PronunciationHotkey_CleanupLanguagePicker() {
    global g_PronunciationLangActive, g_PronunciationLangGui, g_PronunciationLangLv
    global g_PronunciationLangText, g_PronunciationLangHotkeysBound, g_OnEscapePressed

    PronunciationHotkey_UnbindLanguagePickerHotkeys()
    if (g_OnEscapePressed = PronunciationHotkey_OnLanguagePickerEscape)
        g_OnEscapePressed := ""
    g_PronunciationLangActive := false
    g_PronunciationLangText := ""
    g_PronunciationLangLv := false
    if (IsObject(g_PronunciationLangGui)) {
        try g_PronunciationLangGui.Destroy()
        catch {
        }
    }
    g_PronunciationLangGui := false
    g_PronunciationLangHotkeysBound := false
}

PronunciationHotkey_OnLanguagePickerEscape(*) {
    PronunciationHotkey_CleanupLanguagePicker()
    return true
}

PronunciationHotkey_SelectLang(lang) {
    global g_PronunciationLangText
    selectedText := g_PronunciationLangText
    PronunciationHotkey_CleanupLanguagePicker()
    if (lang != "" && selectedText != "")
        PronunciationHotkey_StartLookup(lang, selectedText)
}

PronunciationHotkey_OnLanguagePickerActivate(*) {
    global g_PronunciationLangLv
    if (!IsObject(g_PronunciationLangLv))
        return
    row := 0
    try row := g_PronunciationLangLv.GetNext(0)
    catch {
        row := 0
    }
    if (row < 1)
        return
    lang := ""
    try lang := g_PronunciationLangLv.GetText(row, 3)
    catch {
        lang := ""
    }
    if (lang != "")
        PronunciationHotkey_SelectLang(lang)
}

PronunciationHotkey_BindOne(key, handler) {
    global g_PronunciationLangHotkeyHandlers
    try {
        Hotkey(key, handler, "On")
        g_PronunciationLangHotkeyHandlers.Push({ key: key, handler: handler })
    } catch {
    }
}

PronunciationHotkey_UnbindLanguagePickerHotkeys() {
    global g_PronunciationLangGui, g_PronunciationLangHotkeyHandlers, g_PronunciationLangHotkeysBound
    hwnd := 0
    try {
        if (IsObject(g_PronunciationLangGui))
            hwnd := g_PronunciationLangGui.Hwnd
    } catch {
        hwnd := 0
    }
    if (hwnd) {
        try HotIfWinActive("ahk_id " hwnd)
        catch {
        }
    }
    for handler in g_PronunciationLangHotkeyHandlers {
        try {
            if (handler.HasProp("key") && handler.key != "")
                Hotkey(handler.key, "Off")
        } catch {
        }
    }
    if (hwnd) {
        try HotIf()
        catch {
        }
    }
    g_PronunciationLangHotkeyHandlers := []
    g_PronunciationLangHotkeysBound := false
}

PronunciationHotkey_BindLanguagePickerHotkeys() {
    global g_PronunciationLangGui, g_PronunciationLangHotkeysBound
    PronunciationHotkey_UnbindLanguagePickerHotkeys()
    hwnd := 0
    try {
        if (IsObject(g_PronunciationLangGui))
            hwnd := g_PronunciationLangGui.Hwnd
    } catch {
        hwnd := 0
    }
    if (!hwnd)
        return
    try HotIfWinActive("ahk_id " hwnd)
    catch {
        return
    }

    PronunciationHotkey_BindOne("1", (*) => PronunciationHotkey_SelectLang("en"))
    PronunciationHotkey_BindOne("Numpad1", (*) => PronunciationHotkey_SelectLang("en"))
    PronunciationHotkey_BindOne("e", (*) => PronunciationHotkey_SelectLang("en"))
    PronunciationHotkey_BindOne("E", (*) => PronunciationHotkey_SelectLang("en"))
    PronunciationHotkey_BindOne("2", (*) => PronunciationHotkey_SelectLang("de"))
    PronunciationHotkey_BindOne("Numpad2", (*) => PronunciationHotkey_SelectLang("de"))
    PronunciationHotkey_BindOne("g", (*) => PronunciationHotkey_SelectLang("de"))
    PronunciationHotkey_BindOne("G", (*) => PronunciationHotkey_SelectLang("de"))
    PronunciationHotkey_BindOne("3", (*) => PronunciationHotkey_SelectLang("pt"))
    PronunciationHotkey_BindOne("Numpad3", (*) => PronunciationHotkey_SelectLang("pt"))
    PronunciationHotkey_BindOne("p", (*) => PronunciationHotkey_SelectLang("pt"))
    PronunciationHotkey_BindOne("P", (*) => PronunciationHotkey_SelectLang("pt"))
    PronunciationHotkey_BindOne("Enter", PronunciationHotkey_OnLanguagePickerActivate)
    PronunciationHotkey_BindOne("NumpadEnter", PronunciationHotkey_OnLanguagePickerActivate)
    PronunciationHotkey_BindOne("Escape", PronunciationHotkey_OnLanguagePickerEscape)

    try HotIf()
    catch {
    }
    g_PronunciationLangHotkeysBound := true
}

PronunciationHotkey_BuildLanguagePickerGui() {
    global g_PronunciationLangGui, g_PronunciationLangLv, g_PronunciationLangActive, g_OnEscapePressed

    g_PronunciationLangGui := Gui("+AlwaysOnTop +ToolWindow", "Pronunciation language")
    g_PronunciationLangGui.SetFont("s10", "Segoe UI")
    g_PronunciationLangGui.Add("Text", "xm w420",
        "Char / letter = select   Enter / double-click = select   Esc = cancel")
    g_PronunciationLangLv := g_PronunciationLangGui.Add("ListView", "xm w420 h90 -Multi", ["Char", "Name", "Lang"])
    g_PronunciationLangLv.OnEvent("DoubleClick", PronunciationHotkey_OnLanguagePickerActivate)
    g_PronunciationLangGui.Add("Button", "w100 Section", "Close").OnEvent("Click",
        PronunciationHotkey_OnLanguagePickerEscape)
    g_PronunciationLangGui.OnEvent("Close", PronunciationHotkey_OnLanguagePickerEscape)
    g_PronunciationLangGui.OnEvent("Escape", PronunciationHotkey_OnLanguagePickerEscape)

    ; English first (primary), then German, then Portuguese. Col 3 hidden lang code.
    g_PronunciationLangLv.Add("", "1", "English", "en")
    g_PronunciationLangLv.Add("", "2", "German", "de")
    g_PronunciationLangLv.Add("", "3", "Portuguese", "pt")
    try g_PronunciationLangLv.ModifyCol(1, 50)
    try g_PronunciationLangLv.ModifyCol(2, 350)
    try g_PronunciationLangLv.ModifyCol(3, 0)
    try g_PronunciationLangLv.Modify(1, "Select Focus Vis")

    mon := UtilitySelector_ActiveMonitorWorkArea()
    guiW := 450
    guiH := 180
    guiX := mon.left + (mon.width - guiW) // 2
    guiY := mon.top + (mon.height - guiH) // 2
    if (guiX < mon.left + 20)
        guiX := mon.left + 20
    if (guiY < mon.top + 20)
        guiY := mon.top + 20

    g_PronunciationLangGui.Show("x" . guiX . " y" . guiY . " w" . guiW . " h" . guiH)
    try g_PronunciationLangLv.Focus()
    catch {
    }

    g_PronunciationLangActive := true
    g_OnEscapePressed := PronunciationHotkey_OnLanguagePickerEscape
    PronunciationHotkey_BindLanguagePickerHotkeys()
}

; =============================================================================
; Double-tap timer object
; =============================================================================
class Pronunciation_DoubleTapTimerObj {
    static OnSingleTapTimeout() {
        global g_Pronunciation_DoubleTapArmed, g_Pronunciation_DoubleTapTimer
        if (!g_Pronunciation_DoubleTapArmed)
            return
        g_Pronunciation_DoubleTapArmed := false
        g_Pronunciation_DoubleTapTimer := 0
        ; Off hotkey thread — copy + English lookup
        SetTimer((*) => PronunciationHotkey_CopyAndLookup("en"), -1)
    }
}

PronunciationHotkey_DisarmDoubleTap() {
    global g_Pronunciation_DoubleTapArmed, g_Pronunciation_DoubleTapTimer, g_Pronunciation_LastPressTick
    g_Pronunciation_DoubleTapArmed := false
    g_Pronunciation_LastPressTick := 0
    if (g_Pronunciation_DoubleTapTimer) {
        SetTimer(g_Pronunciation_DoubleTapTimer, 0)
        g_Pronunciation_DoubleTapTimer := 0
    }
}

; =============================================================================
; Hotkey: Win+Alt+Shift+8 — 1× en, 2× de, hold = language ListView
; =============================================================================
#!+8:: {
    global g_StandardLoadingBarIsKeysOverlay
    global g_PronunciationLangActive, g_PronunciationLangGui
    global g_Pronunciation_DoubleTapArmed, g_Pronunciation_LastPressTick, g_Pronunciation_DoubleTapTimer

    if (g_StandardLoadingBarIsKeysOverlay) {
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        return
    }
    if (g_PronunciationLangActive && IsObject(g_PronunciationLangGui)) {
        PronunciationHotkey_CleanupLanguagePicker()
        return
    }

    ; Detect double-tap on key-down (before KeyWait) so a slow release cannot miss the window.
    pressTime := A_TickCount
    elapsed := (g_Pronunciation_LastPressTick > 0) ? (pressTime - g_Pronunciation_LastPressTick) : 9999
    isSecondTap := g_Pronunciation_DoubleTapArmed && elapsed >= 0 && elapsed < PRONUNCIATION_DOUBLE_TAP_MS

    KeyWait "8", "T" . (PRONUNCIATION_HOLD_MS / 1000)
    isHold := (A_TickCount - pressTime) >= PRONUNCIATION_HOLD_MS

    if (isHold) {
        PronunciationHotkey_DisarmDoubleTap()
        ; Run after hotkey returns — GUI/hotkey bind from inside #!+8 can deadlock.
        SetTimer((*) => PronunciationHotkey_CopyAndShowPicker(), -1)
        return
    }

    if (isSecondTap) {
        PronunciationHotkey_DisarmDoubleTap()
        SetTimer((*) => PronunciationHotkey_CopyAndLookup("de"), -1)
        return
    }

    ; Arm after quick release: window starts now (matches #!+9 / AI_QD).
    if (g_Pronunciation_DoubleTapTimer) {
        SetTimer(g_Pronunciation_DoubleTapTimer, 0)
        g_Pronunciation_DoubleTapTimer := 0
    }
    g_Pronunciation_LastPressTick := A_TickCount
    g_Pronunciation_DoubleTapArmed := true
    g_Pronunciation_DoubleTapTimer := ObjBindMethod(Pronunciation_DoubleTapTimerObj, "OnSingleTapTimeout")
    SetTimer(g_Pronunciation_DoubleTapTimer, -PRONUNCIATION_DOUBLE_TAP_MS)
}
