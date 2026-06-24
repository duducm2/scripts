; =============================================================================
; Utils module: hotstring_selector_cleanup.ahk
; CleanupHotstringSelector and auto-close idle
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

global g_HotstringSelectorAutoCloseMs := 10000 ; Auto-close Utility Shortcuts if no choice made in time

HotstringSelector_AutoCloseIfIdle() {
    global g_HotstringSelectorActive
    if (g_HotstringSelectorActive)
        CleanupHotstringSelector()
}

; =============================================================================
; CleanupHotstringSelector()
; =============================================================================
; PURPOSE: Closes hotstring selector GUI and disables all associated hotkeys.
;          Called when selector is closed via Escape key, character selection, or toggle.
;
; PROCESS:
;   1. Sets g_HotstringSelectorActive = false to prevent further character processing
;   2. Disables all character hotkeys (including uppercase variants for lowercase letters)
;   3. Handles special VK codes for comma (vkBC) and period (vkBE)
;   4. Disables Escape key handler
;   5. Destroys GUI object if it exists
;   6. Clears hotkey handlers array
;   7. Clears character mapping maps
;
; RETURNS: None (void function)
; SIDE EFFECTS: Resets all global state variables to initial values
; =============================================================================
CleanupHotstringSelector() {
    global g_HotstringSelectorActive, g_HotstringSelectorGui, g_HotstringHotkeyHandlers
    global g_HotstringPromptCharMap, g_HotstringGeminiArmed
    global g_HotstringCharMap
    global g_HS_SelectorOpenFile, g_HS_SelectorCloseRequestFile, g_HS_SelectorCloseCheckTimer
    global g_UtilitySelectorMode, g_UtilitySelectorCategory

    SetTimer(HotstringSelector_AutoCloseIfIdle, 0)
    ; Disable active flag
    g_HotstringSelectorActive := false

    ; Disable all character hotkeys
    for handler in g_HotstringHotkeyHandlers {
        try {
            key := handler.HasProp("key") ? handler.key : handler.char
            char := handler.HasProp("char") ? handler.char : key
            ; Handle special VK codes
            if (key = "vkBC" || char = ",") {
                Hotkey("vkBC", "Off")
            } else if (key = "vkBE" || char = ".") {
                Hotkey("vkBE", "Off")
            } else {
                Hotkey(key, "Off")
            }
        } catch {
            ; Silently ignore errors
        }
    }

    ; Disable Escape hotkey
    try {
        Hotkey("Escape", "Off")
    } catch {
        ; Ignore
    }
    Utils_EnsureGlobalEscapeHotkey()

    ; Disable Backspace hotkey (menu back)
    try {
        Hotkey("Backspace", "Off")
    } catch {
        ; Ignore
    }

    ; Stop IPC timer and clear sentinel files
    try SetTimer(Utils_CheckHotstringSelectorCloseRequest, 0)
    g_HS_SelectorCloseCheckTimer := ""
    try FileDelete(g_HS_SelectorOpenFile)
    catch {
    }
    try FileDelete(g_HS_SelectorCloseRequestFile)
    catch {
    }

    ; Clear handlers array
    g_HotstringHotkeyHandlers := []
    g_HotstringPromptCharMap := Map()
    g_HotstringGeminiArmed := false
    g_UtilitySelectorMode := "top"
    g_UtilitySelectorCategory := ""

    ; Close and destroy GUI
    if (IsObject(g_HotstringSelectorGui)) {
        try {
            g_HotstringSelectorGui.Destroy()
        } catch {
            ; Ignore
        }
        g_HotstringSelectorGui := false
    }

    ; Clear char map
    g_HotstringCharMap := Map()
}
