; =============================================================================
; Utils module: cursor_composer_focus.ahk
; Cursor AI text field focus and VS Code chat helpers
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; Cursor AI text field focus (shared logic for Gemini transfer and WindowManagement)
; =============================================================================

; VS Code exposes both the main editor and chat composer as native-edit-context Edit controls.
; To avoid false positives, pick the lowest compact edit near the Send button region.
VSCode_FindChatInputField(root) {
    bestEdit := ""
    bestTop := -2147483647
    sendBr := ""
    if (!root)
        return ""
    try {
        sendBtn := VSCode_FindChatSendButton(root)
        if (sendBtn)
            sendBr := sendBtn.BoundingRectangle
    } catch {
        sendBr := ""
    }
    try {
        allEdits := root.FindAll({ Type: UIA.Type.Edit })
        for editEl in allEdits {
            try {
                className := editEl.ClassName
                if (!InStr(className, "native-edit-context"))
                    continue
                if (editEl.GetPropertyValue(UIA.Property.IsOffscreen))
                    continue
                br := editEl.BoundingRectangle
                h := br.b - br.t
                w := br.r - br.l
                if (h <= 0 || w <= 0)
                    continue

                ; Skip large editor surfaces; chat composer is typically compact.
                if (h > 260)
                    continue

                ; When send button exists, require geometric proximity to chat footer area.
                if (sendBr != "") {
                    cx := (br.l + br.r) / 2
                    cy := (br.t + br.b) / 2
                    sendCx := (sendBr.l + sendBr.r) / 2
                    sendCy := (sendBr.t + sendBr.b) / 2
                    if (Abs(cx - sendCx) > 900)
                        continue
                    if (Abs(cy - sendCy) > 420)
                        continue
                }

                ; Prefer the visually lowest eligible edit field.
                if (br.t >= bestTop) {
                    bestTop := br.t
                    bestEdit := editEl
                }
            } catch {
                continue
            }
        }
    } catch {
    }
    return bestEdit
}

VSCode_EnsureChatInputHasFocus(editEl) {
    if (!editEl)
        return false
    try {
        editEl.SetFocus()
    } catch {
    }
    loop 3 {
        try {
            if (editEl.HasKeyboardFocus)
                return true
        } catch {
        }
        Sleep 40
    }
    try {
        editEl.ScrollIntoView()
    } catch {
    }
    try {
        editEl.Click()
    } catch {
    }
    Sleep 80
    try {
        return editEl.HasKeyboardFocus
    } catch {
        return false
    }
}

VSCode_FindChatSendButton(root) {
    if (!root)
        return ""
    lastMatchingButton := ""
    try {
        buttons := root.FindAll({ Type: UIA.Type.Button })
        for buttonEl in buttons {
            try {
                buttonName := buttonEl.Name
                buttonClass := buttonEl.ClassName
                if (!(InStr(buttonName, "Send ") || InStr(buttonName, "Send to New Chat")))
                    continue
                if (InStr(buttonClass, "arrow-up"))
                    return buttonEl
                lastMatchingButton := buttonEl
            } catch {
                continue
            }
        }
    } catch {
    }
    return lastMatchingButton
}

VSCode_IsChatSendReady(targetHwnd) {
    if (!IsSet(UIA))
        return true
    try {
        root := UIA.ElementFromHandle(targetHwnd)
        sendButton := VSCode_FindChatSendButton(root)
        if (!sendButton)
            return false
        try {
            if (sendButton.GetPropertyValue(UIA.Property.IsEnabled))
                return true
        } catch {
        }
        try {
            return !InStr(sendButton.ClassName, "disabled")
        } catch {
        }
    } catch {
    }
    return false
}

VSCode_IsChatInputFocused(targetHwnd) {
    if (!IsSet(UIA))
        return false
    try {
        root := UIA.ElementFromHandle(targetHwnd)
        ; Fast path: check which element currently has focus.
        ; The main code editor uses native-edit-context but is far from the Send button.
        try {
            focusedEl := UIA.GetFocusedElement()
            if (focusedEl) {
                focusedClass := ""
                try focusedClass := focusedEl.ClassName
                if (InStr(focusedClass, "native-edit-context")) {
                    sendBtn := VSCode_FindChatSendButton(root)
                    if (!sendBtn)
                        return false
                    try {
                        focusedBr := focusedEl.BoundingRectangle
                        sendBr := sendBtn.BoundingRectangle
                        if (Abs(((focusedBr.l + focusedBr.r) / 2) - ((sendBr.l + sendBr.r) / 2)) > 900)
                            return false
                        if (Abs(((focusedBr.t + focusedBr.b) / 2) - ((sendBr.t + sendBr.b) / 2)) > 420)
                            return false
                        return true
                    } catch {
                        return false
                    }
                }
            }
        } catch {
        }
        ; Fallback: scan for chat input and verify keyboard focus.
        editEl := VSCode_FindChatInputField(root)
        if (!editEl)
            return false
        try {
            return editEl.HasKeyboardFocus
        } catch {
        }
    } catch {
    }
    return false
}

VSCode_IsSafeChatPasteTarget(targetHwnd) {
    ; Conservative gate: if uncertain, do not paste.
    if (!IsSet(UIA))
        return false
    if (!targetHwnd || !WinExist("ahk_id " targetHwnd))
        return false
    try {
        root := UIA.ElementFromHandle(targetHwnd)
        if (!root)
            return false
        editEl := VSCode_FindChatInputField(root)
        sendBtn := VSCode_FindChatSendButton(root)
        if (!editEl || !sendBtn)
            return false

        ; Must be keyboard-focused to avoid pasting into editor/other controls.
        try {
            if (!editEl.HasKeyboardFocus)
                return false
        } catch {
            return false
        }

        editBr := editEl.BoundingRectangle
        sendBr := sendBtn.BoundingRectangle
        ew := editBr.r - editBr.l
        eh := editBr.b - editBr.t
        if (ew < 260 || eh <= 0 || eh > 260)
            return false

        ; Composer and send button must be in the same footer region.
        editCx := (editBr.l + editBr.r) / 2
        editCy := (editBr.t + editBr.b) / 2
        sendCx := (sendBr.l + sendBr.r) / 2
        sendCy := (sendBr.t + sendBr.b) / 2
        if (Abs(editCx - sendCx) > 900)
            return false
        if (Abs(editCy - sendCy) > 260)
            return false

        ; Require right-side chat pane placement (avoids inline "Get comment" editors).
        rect := Buffer(16, 0)
        if (!DllCall("GetWindowRect", "ptr", targetHwnd, "ptr", rect))
            return false
        winLeft := NumGet(rect, 0, "int")
        winRight := NumGet(rect, 8, "int")
        winWidth := winRight - winLeft
        if (winWidth <= 0)
            return false
        thresholdX := winLeft + (winWidth * 0.52)
        if (editCx < thresholdX || sendCx < thresholdX)
            return false

        return true
    } catch {
    }
    return false
}

VSCode_SubmitChat(targetHwnd) {
    if (IsSet(UIA)) {
        try {
            root := UIA.ElementFromHandle(targetHwnd)
            sendButton := VSCode_FindChatSendButton(root)
            if (sendButton && VSCode_IsChatSendReady(targetHwnd)) {
                try {
                    sendButton.Click()
                    return true
                } catch {
                }
                try {
                    sendButton.SetFocus()
                    Sleep 40
                    SendInput "{Enter}"
                    return true
                } catch {
                }
            }
        } catch {
        }
    }
    SendInput "{Enter}"
    return true
}

; Activate VS Code window and focus chat input. Returns true on success.
VSCode_FocusChatInput(targetHwnd := 0) {
    try {
        if (targetHwnd) {
            WinActivate("ahk_id " targetHwnd)
            if (!WinWaitActive("ahk_id " targetHwnd, , 2))
                return false
        } else {
            targetHwnd := WinExist("ahk_exe Code.exe")
            if (!targetHwnd)
                return false
            WinActivate("ahk_id " targetHwnd)
            if (!WinWaitActive("ahk_id " targetHwnd, , 2))
                return false
        }
        Sleep 180

        ; Ensure the chat surface exists first.
        SendInput "^!i"
        Sleep 350

        if (!IsSet(UIA)) {
            SendInput "{Tab}"
            Sleep 120
            return true
        }

        loop 2 {
            try {
                root := UIA.ElementFromHandle(targetHwnd)
                editEl := VSCode_FindChatInputField(root)
                if (editEl && VSCode_EnsureChatInputHasFocus(editEl)) {
                    return true
                }
            } catch {
            }

            ; Fallback nudge inside the chat view without clicking toolbar buttons.
            SendInput "{Tab}"
            Sleep 140
        }

        return false
    } catch {
        return false
    }
}

Cursor_EnsureComposerHasFocus(editEl) {
    if (!editEl)
        return false
    try {
        editEl.SetFocus()
    } catch {
    }
    loop 3 {
        try {
            if (editEl.HasKeyboardFocus)
                return true
        } catch {
        }
        Sleep 40
    }
    try {
        editEl.ScrollIntoView()
    } catch {
    }
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

Cursor_FindComposerInput(root) {
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

; Activate Cursor window and focus AI composer input. Returns true on success.
Cursor_FocusAITextField(targetHwnd := 0) {
    try {
        if (targetHwnd) {
            WinActivate("ahk_id " targetHwnd)
            WinWaitActive("ahk_id " targetHwnd, , 2)
        } else {
            targetHwnd := WinExist("ahk_exe Cursor.exe")
            if (!targetHwnd)
                return false
            WinWaitActive("ahk_id " targetHwnd, , 2)
        }
        Sleep 200
        paneWasOpen := false
        focusDone := false
        if (IsSet(UIA)) {
            try {
                root := UIA.ElementFromHandle(targetHwnd)
                if (root) {
                    toggleEl := root.FindFirst({ Type: UIA.Type.CheckBox, Name: "Toggle AI Pane", matchmode: 2 })
                    paneOpen := toggleEl && InStr(toggleEl.ClassName, "checked")
                    paneWasOpen := paneOpen
                    if (paneOpen) {
                        editEl := Cursor_FindComposerInput(root)
                        if (editEl) {
                            if (Cursor_EnsureComposerHasFocus(editEl))
                                focusDone := true
                        }
                    } else {
                        editEl := Cursor_FindComposerInput(root)
                        if (editEl) {
                            if (Cursor_EnsureComposerHasFocus(editEl))
                                focusDone := true
                        } else {
                            Send "^i"
                            loop 15 {
                                Sleep 200
                                root := UIA.ElementFromHandle(targetHwnd)
                                editEl := Cursor_FindComposerInput(root)
                                if (editEl) {
                                    if (Cursor_EnsureComposerHasFocus(editEl))
                                        focusDone := true
                                    break
                                }
                            }
                        }
                    }
                }
            } catch {
            }
        }
        if (!focusDone) {
            if (IsSet(UIA)) {
                try {
                    root := UIA.ElementFromHandle(targetHwnd)
                    if (root) {
                        editEl := Cursor_FindComposerInput(root)
                        if (editEl) {
                            if (Cursor_EnsureComposerHasFocus(editEl))
                                focusDone := true
                        }
                    }
                } catch {
                }
            }
            if (!focusDone) {
                if (!paneWasOpen) {
                    Send "^i"
                    Sleep 1200
                }
                return false
            }
        }
        return true
    } catch {
        return false
    }
}

; Minimum clipboard length for transfer (match Gemini/bridge validation)
CURSOR_TRANSFER_MIN_CLIPBOARD_LENGTH := 10
; Electron/Cursor needs time after paste before Enter; after Enter before foreground changes or submit can drop.
CURSOR_TRANSFER_POST_PASTE_BEFORE_ENTER_MS := 80
CURSOR_TRANSFER_POST_ENTER_BEFORE_RESTORE_MS := 400
VSCODE_TRANSFER_POST_PASTE_SETTLE_MS := 180
VSCODE_TRANSFER_PASTE_RETRY_COUNT := 2

; Activate Cursor/VS Code window, focus AI field, paste clipboard, send Enter. Non-blocking feedback on failure.
; restoreFocusHwnd: if set, WinActivate this window after Enter is processed (before success overlay) so focus does not stay on the target window.
CursorTransfer_ActivateFocusPaste(targetHwnd, restoreFocusHwnd := 0) {
    appDisplayName := CursorTransfer_GetTargetAppName()
    targetIsVSCode := (CursorTransfer_GetTargetAppExecutable() = "Code.exe")
    if (!targetHwnd || !WinExist("ahk_id " targetHwnd)) {
        ShowCenteredOverlay_Utils("❌ " appDisplayName " window not found", 2000, BANNER_ACCENT_ERROR)
        return
    }
    clip := Trim(A_Clipboard)
    if (clip = "" || StrLen(clip) < CURSOR_TRANSFER_MIN_CLIPBOARD_LENGTH) {
        ShowCenteredOverlay_Utils("❌ Clipboard empty or too short", 2000, BANNER_ACCENT_ERROR)
        return
    }
    try {
        WinActivate("ahk_id " targetHwnd)
        if (!WinWaitActive("ahk_id " targetHwnd, , 2)) {
            ShowCenteredOverlay_Utils("❌ Could not activate " appDisplayName, 2000, BANNER_ACCENT_ERROR)
            return
        }
        Sleep 100
        focusSucceeded := false
        if (targetIsVSCode) {
            focusSucceeded := VSCode_FocusChatInput(targetHwnd)
        } else {
            focusSucceeded := Cursor_FocusAITextField(targetHwnd)
        }
        if (!focusSucceeded) {
            ShowCenteredOverlay_Utils("❌ Could not focus AI field", 2000, BANNER_ACCENT_ERROR)
            return
        }
        Sleep(targetIsVSCode ? 220 : 100)
        try {
            if (WinGetID("A") != targetHwnd) {
                WinActivate("ahk_id " targetHwnd)
                WinWaitActive("ahk_id " targetHwnd, , 2)
            }
        } catch {
        }
        if (Trim(A_Clipboard) = "" || StrLen(Trim(A_Clipboard)) < CURSOR_TRANSFER_MIN_CLIPBOARD_LENGTH) {
            ShowCenteredOverlay_Utils("❌ Clipboard lost before paste", 2000, BANNER_ACCENT_ERROR)
            return
        }
        pasteDetected := true
        pasteAttempts := targetIsVSCode ? VSCODE_TRANSFER_PASTE_RETRY_COUNT : 1
        loop pasteAttempts {
            if (targetIsVSCode) {
                ; Hard gate: never paste until strict chat-pane targeting is verified.
                if (!VSCode_IsSafeChatPasteTarget(targetHwnd)) {
                    if (!VSCode_FocusChatInput(targetHwnd))
                        break
                    Sleep 160
                    ; Require stability across two checks to avoid transient focus races.
                    if (!VSCode_IsSafeChatPasteTarget(targetHwnd)) {
                        Sleep 90
                    }
                    if (!VSCode_IsSafeChatPasteTarget(targetHwnd)) {
                        if (A_Index < pasteAttempts)
                            continue
                        break
                    }
                }
            }
            SendInput "^v"
            Sleep CURSOR_TRANSFER_POST_PASTE_BEFORE_ENTER_MS + (targetIsVSCode ? VSCODE_TRANSFER_POST_PASTE_SETTLE_MS :
                0)
            if (!targetIsVSCode) {
                pasteDetected := true
                break
            }
            pasteDetected := VSCode_IsChatSendReady(targetHwnd)
            if (pasteDetected)
                break
            if (A_Index < pasteAttempts) {
                if (!VSCode_FocusChatInput(targetHwnd))
                    break
                Sleep 150
            }
        }
        if (!pasteDetected) {
            ShowCenteredOverlay_Utils("❌ Paste blocked: AI text field not confidently focused", 2600,
                BANNER_ACCENT_ERROR)
            return
        }
        if (targetIsVSCode) {
            if (!VSCode_SubmitChat(targetHwnd)) {
                ShowCenteredOverlay_Utils("❌ Could not submit VS Code chat", 2200, BANNER_ACCENT_ERROR)
                return
            }
        } else {
            SendInput "{Enter}"
        }
        if (restoreFocusHwnd && WinExist("ahk_id " restoreFocusHwnd)) {
            ; Wait so target editor keeps foreground until paste + Enter are processed; restoring sooner drops Enter.
            Sleep CURSOR_TRANSFER_POST_ENTER_BEFORE_RESTORE_MS
            try {
                WinActivate("ahk_id " restoreFocusHwnd)
                if (!WinActive("ahk_id " restoreFocusHwnd))
                    WinWaitActive("ahk_id " restoreFocusHwnd, , 0.5)
            } catch {
            }
        }
        appDisplayName := CursorTransfer_GetTargetAppName()
        ShowCenteredOverlay_Utils("✅ Sent to " . appDisplayName, 1500, BANNER_ACCENT_SUCCESS)
    } catch as err {
        ShowCenteredOverlay_Utils("❌ Transfer failed", 2000, BANNER_ACCENT_ERROR)
    }
}
