; =============================================================================
; WindowManagement module: cursor_composer.ahk
; Cursor AI composer focus (WM_EnsureComposerHasFocus / FocusCursorAITextField)
; Extracted verbatim from WindowManagement.ahk; loaded via #include into the
; WindowManagement.ahk process, which remains the entry point / source of truth.
; =============================================================================

; =============================================================================
; Focus Cursor AI text field. Handles both UI states via UIA when available:
; - AI side panel hidden: send Ctrl+I to open, then wait for and focus composer input via UIA.
; - AI side panel open: focus composer input via UIA only (do not send Ctrl+I or panel closes).
; - If UIA cannot locate the composer input, do NOT fall back to generic Tab navigation (which can hit toolbar buttons like Go Back).
; targetHwnd: if provided, explicitly activate this window first (ensures paste goes to correct project).
; =============================================================================
WM_EnsureComposerHasFocus(editEl) {
    if (!editEl)
        return false
    try {
        editEl.SetFocus()
    } catch {
    }
    ; Bounded retry loop: check HasKeyboardFocus a few times with small delays.
    loop 3 {
        try {
            if (editEl.HasKeyboardFocus)
                return true
        } catch {
        }
        Sleep 40
    }
    ; If simple SetFocus did not succeed, try scroll + click once, then re-check.
    try {
        editEl.ScrollIntoView()
    } catch {
    }
    WMAutomation_SuppressCursorCentering("cursor_composer_click", 1200)
    try {
        editEl.Click()
    } catch {
    }
    Sleep 60
    try {
        return editEl.HasKeyboardFocus
    } catch {
        return false
    }
}

FocusCursorAITextField(targetHwnd := 0) {
    try {
        if (targetHwnd) {
            WMAutomation_SuppressCursorCentering("cursor_focus_textfield", 1800)
            WinActivate("ahk_id " targetHwnd)
            WinWaitActive("ahk_id " targetHwnd, , 2)
        } else {
            targetHwnd := WinExist("ahk_exe Cursor.exe")
            if (!targetHwnd)
                return false
            WMAutomation_SuppressCursorCentering("cursor_focus_textfield", 1800)
            WinWaitActive("ahk_id " targetHwnd, , 2)
        }
        Sleep 200

        ; Show loading indicator while navigating to the AI text field.
        ; Center on the Cursor window and use the standard intermediate accent color.
        stateText := "⏳ Focando campo de texto da IA..."
        StandardLoadingBar_Show(stateText, BANNER_ACCENT_INTERMEDIATE, { centerOnHwnd: 0, passive: false })
        barShown := true

        ; Track whether AI pane was detected as open so we only ever send Ctrl+I to open it (never to close it).
        paneWasOpen := false
        focusDone := false
        if (IsSet(UIA)) {
            try {
                root := UIA.ElementFromHandle(targetHwnd)
                if (root) {
                    ; Detect AI pane state: "Toggle AI Pane (Ctrl+Alt+B)" CheckBox has "checked" in ClassName when open
                    toggleEl := root.FindFirst({ Type: UIA.Type.CheckBox, Name: "Toggle AI Pane", matchmode: 2 })
                    paneOpen := toggleEl && InStr(toggleEl.ClassName, "checked")
                    paneWasOpen := paneOpen

                    if (paneOpen) {
                        ; Panel already open: focus composer input directly (do not send Ctrl+I)
                        editEl := _WM_FindCursorComposerInput(root)
                        if (editEl) {
                            if (WM_EnsureComposerHasFocus(editEl))
                                focusDone := true
                        }
                    } else {
                        ; Additional safety: if we can already find the composer, treat as open and DO NOT send Ctrl+I.
                        editEl := _WM_FindCursorComposerInput(root)
                        if (editEl) {
                            if (WM_EnsureComposerHasFocus(editEl))
                                focusDone := true
                        } else {
                            ; Panel hidden: open with Ctrl+I, then wait for composer input and focus via UIA
                            Send "^i"
                            loop 15 {
                                Sleep 200
                                root := UIA.ElementFromHandle(targetHwnd)
                                editEl := _WM_FindCursorComposerInput(root)
                                if (editEl) {
                                    if (WM_EnsureComposerHasFocus(editEl))
                                        focusDone := true
                                    break
                                }
                            }
                        }
                    }
                }
            } catch {
                ; UIA failed; fall through to keyboard path
            }
        }

        if (!focusDone) {
            ; Before any keyboard fallback, try one more UIA-based composer search without relying on the toggle.
            if (IsSet(UIA)) {
                try {
                    root := UIA.ElementFromHandle(targetHwnd)
                    if (root) {
                        editEl := _WM_FindCursorComposerInput(root)
                        if (editEl) {
                            if (WM_EnsureComposerHasFocus(editEl))
                                focusDone := true
                        }
                    }
                } catch {
                }
            }

            if (!focusDone) {
                ; Keyboard fallback: only send Ctrl+I if the pane was not previously detected as open.
                if (!paneWasOpen) {
                    Send "^i"
                    Sleep 1200
                }
                ; Avoid blind Tab navigation that can land on title-bar navigation buttons (Go Back / Forward).
                ; Without a reliable target, leave focus as-is and report failure.
                return false
            }
        }
        return true
    } catch {
        return false
    } finally {
        ; Always hide the loading indicator once navigation is complete or has failed.
        try StandardLoadingBar_Hide(0)
        catch {
        }
    }
}

; Returns the Cursor composer input Edit (aislash-editor-input, not readonly) or "" if not found.
_WM_FindCursorComposerInput(root) {
    try {
        allEdits := root.FindAll({ Type: UIA.Type.Edit })
        for editEl in allEdits {
            cn := editEl.ClassName
            if (InStr(cn, "aislash-editor-input") && !InStr(cn, "readonly"))
                return editEl
        }
    } catch {
    }
    return ""
}
