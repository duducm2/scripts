; =============================================================================
; Utils module: hotstring_selector_cleanup.ahk
; CleanupHotstringSelector
; =============================================================================

CleanupHotstringSelector() {
    global g_HotstringSelectorActive, g_HotstringSelectorGui, g_HotstringSelectorLv, g_HotstringSelectorHint
    global g_HotstringHotkeyHandlers
    global g_HotstringPromptCharMap, g_HotstringGeminiArmed
    global g_HS_SelectorOpenFile, g_HS_SelectorCloseRequestFile, g_HS_SelectorCloseCheckTimer
    global g_UtilitySelectorMode, g_UtilitySelectorCategory, g_UtilitySelectorHotkeysBound
    global g_UtilitySelectorRows, g_OnEscapePressed
    global g_UtilitySelectorFilterQuery, g_UtilitySelectorFilterTyping, g_UtilitySelectorSuppressFilterKillFocus
    global g_HotstringSelectorFilterCtrl

    g_HotstringSelectorActive := false
    g_UtilitySelectorFilterTyping := false
    g_UtilitySelectorSuppressFilterKillFocus := false
    g_UtilitySelectorFilterQuery := ""
    if (IsObject(g_HotstringSelectorFilterCtrl)) {
        try g_HotstringSelectorFilterCtrl.Value := ""
        catch {
        }
    }
    try SetTimer(UtilitySelector_StartIpc, 0)
    catch {
    }

    UtilitySelector_UnbindModalHotkeys()

    if (g_OnEscapePressed = HandleHotstringEscape)
        g_OnEscapePressed := ""
    Utils_EnsureGlobalEscapeHotkey()

    try SetTimer(Utils_CheckHotstringSelectorCloseRequest, 0)
    g_HS_SelectorCloseCheckTimer := ""
    try FileDelete(g_HS_SelectorOpenFile)
    catch {
    }
    try FileDelete(g_HS_SelectorCloseRequestFile)
    catch {
    }

    g_HotstringHotkeyHandlers := []
    g_HotstringPromptCharMap := Map()
    g_HotstringGeminiArmed := false
    g_UtilitySelectorMode := "top"
    g_UtilitySelectorCategory := ""
    g_UtilitySelectorRows := []
    g_UtilitySelectorHotkeysBound := false

    ; Hide and reuse; Destroy is expensive and forces a full Gui rebuild on the next #!+U.
    if (IsObject(g_HotstringSelectorGui)) {
        try g_HotstringSelectorGui.Hide()
        catch {
        }
    }
}
