; =============================================================================
; Utils module: clip_angel_constant_paste.ahk
; Constant Pasting: toggle loop pastes from current Clip Angel selection
; (All or Favorites; no Row-0 jump), Enter only for text, 1.5s interruptible gap.
; Loaded via #include into Utils.ahk after clip_angel_favorite / activate.
; =============================================================================

CONSTANT_PASTE_GAP_MS := 1500
CONSTANT_PASTE_POLL_MS := 50
CONSTANT_PASTE_CLIPBOARD_WAIT_MS := 400

global g_ClipAngelConstantPasteActive := false
global g_ClipAngelConstantPasteStopRequested := false

ClipAngel_ConstantPaste_ClipboardHasImage() {
    try {
        return !!(DllCall("IsClipboardFormatAvailable", "UInt", 8, "Int")
            || DllCall("IsClipboardFormatAvailable", "UInt", 17, "Int")
            || DllCall("IsClipboardFormatAvailable", "UInt", 2, "Int"))
    } catch {
        return false
    }
}

; True when clipboard looks like text-only (not image / file drop). Ambiguous => false.
ClipAngel_ConstantPaste_IsTextOnlyClip() {
    if ClipAngel_ConstantPaste_ClipboardHasImage()
        return false
    try {
        if Clipboard_HasFileDrop()
            return false
    } catch {
    }
    try {
        if DllCall("IsClipboardFormatAvailable", "UInt", 13, "Int")  ; CF_UNICODETEXT
            return true
        if DllCall("IsClipboardFormatAvailable", "UInt", 1, "Int")   ; CF_TEXT
            return true
    } catch {
    }
    return false
}

ClipAngel_ConstantPaste_WaitClipboardSettle(timeoutMs := 0) {
    if (!timeoutMs)
        timeoutMs := CONSTANT_PASTE_CLIPBOARD_WAIT_MS
    try ClipWait(timeoutMs / 1000.0, 1)
    catch {
    }
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        if (ClipAngel_ConstantPaste_ClipboardHasImage())
            return
        try {
            if Clipboard_HasFileDrop()
                return
        } catch {
        }
        try {
            if (DllCall("IsClipboardFormatAvailable", "UInt", 13, "Int")
                || DllCall("IsClipboardFormatAvailable", "UInt", 1, "Int"))
                return
        } catch {
        }
        Sleep CONSTANT_PASTE_POLL_MS
    }
}

; Selected DataGrid row name ("Row N") or "" if none.
ClipAngel_ConstantPaste_GetSelectedRowName(hwnd, root := 0) {
    if !hwnd
        return ""
    try {
        if !root {
            root := UIA.ElementFromHandle(hwnd)
            if !root
                return ""
        }
        dataGrid := ClipAngel_UiaGetDataGrid(hwnd, root)
        if !dataGrid
            return ""
        try {
            if dataGrid.GetPropertyValue(UIA.Property.IsSelectionPatternAvailable) {
                sel := dataGrid.GetSelectionPattern().GetSelection()
                if (sel && sel.Length >= 1) {
                    try {
                        n := sel[1].Name
                        if RegExMatch(n, "i)^(?:Row|Linha)\s*\d+")
                            return n
                    } catch {
                    }
                }
            }
        } catch {
        }
        rows := 0
        try rows := dataGrid.FindAll({ Type: 50025 })
        catch
            rows := 0
        if !rows
            return ""
        for row in rows {
            try {
                if row.GetPropertyValue(UIA.Property.SelectionItemIsSelected)
                    return row.Name
            } catch {
            }
            if ClipAngel_UiaRowLegacySelected(row) {
                try return row.Name
                catch
                    return ""
            }
        }
    } catch {
    }
    return ""
}

ClipAngel_ConstantPaste_ControlSend(hwnd, keys) {
    if !hwnd
        return false
    try {
        ControlSend(keys, , "ahk_id " hwnd)
        return true
    } catch {
        return false
    }
}

; Interruptible gap: returns false if stop requested.
ClipAngel_ConstantPaste_WaitGap(timeoutMs := 0) {
    global g_ClipAngelConstantPasteActive, g_ClipAngelConstantPasteStopRequested
    if (!timeoutMs)
        timeoutMs := CONSTANT_PASTE_GAP_MS
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        if (!g_ClipAngelConstantPasteActive || g_ClipAngelConstantPasteStopRequested)
            return false
        Sleep CONSTANT_PASTE_POLL_MS
    }
    return g_ClipAngelConstantPasteActive && !g_ClipAngelConstantPasteStopRequested
}

ClipAngel_ConstantPaste_PrepareGrid(hwnd, &root, &dataGrid) {
    root := 0
    dataGrid := 0
    if !hwnd
        return false
    try root := UIA.ElementFromHandle(hwnd)
    catch
        root := 0
    dataGrid := ClipAngel_UiaGetDataGrid(hwnd, root)
    if !dataGrid
        return false
    ClipAngel_UiaEnsureGridListFocus(dataGrid, hwnd, root)
    return true
}

ClipAngel_ConstantPaste_Toggle(*) {
    global g_ClipAngelConstantPasteActive, g_ClipAngelConstantPasteStopRequested
    if (g_ClipAngelConstantPasteActive) {
        g_ClipAngelConstantPasteStopRequested := true
        g_ClipAngelConstantPasteActive := false
        return
    }
    ClipAngel_ConstantPaste_Run()
}

ClipAngel_ConstantPaste_Run() {
    global g_ClipAngelConstantPasteActive, g_ClipAngelConstantPasteStopRequested
    if (g_ClipAngelConstantPasteActive)
        return false

    priorHwnd := ClipAngel_ResolvePriorHwnd(0)
    if (!priorHwnd) {
        ShowCenteredOverlay_Utils(
            "❌ Constant Pasting: focus the paste target first (not Clip Angel).",
            2500, BANNER_ACCENT_ERROR)
        return false
    }
    if !ClipAngel_TryAcquireAutomationLock()
        return false

    g_ClipAngelConstantPasteActive := true
    g_ClipAngelConstantPasteStopRequested := false
    pastedCount := 0
    stopReason := "stopped"

    StandardLoadingBar_Show("⏳ Constant Pasting…  [#!+;] stop", BANNER_ACCENT_INFO, {
        passive: true,
        fontSize: 17,
        textWidth: 480
    })

    try {
        ClipAngel_WaitChordModifiersReleased()
        ClipAngel_ReleaseChordModifiersForSend()

        ; Show CA without forcing Row 0 or MarkFilter.
        if !ActivateClipAngelWithFocusCorrection(true, 0, true, false) {
            stopReason := "Clip Angel not available"
            return false
        }

        hwnd := ClipAngel_MainHwnd()
        if !hwnd {
            stopReason := "Clip Angel window missing"
            return false
        }

        root := 0
        dataGrid := 0
        if !ClipAngel_ConstantPaste_PrepareGrid(hwnd, &root, &dataGrid) {
            stopReason := "clip list not ready"
            return false
        }

        while (g_ClipAngelConstantPasteActive && !g_ClipAngelConstantPasteStopRequested) {
            hwnd := ClipAngel_MainHwnd()
            if !hwnd {
                stopReason := "Clip Angel closed"
                break
            }
            if !ClipAngel_ConstantPaste_PrepareGrid(hwnd, &root, &dataGrid) {
                stopReason := "clip list not ready"
                break
            }

            rowBefore := ClipAngel_ConstantPaste_GetSelectedRowName(hwnd, root)
            if (rowBefore = "") {
                stopReason := "no clip selected"
                break
            }

            ClipAngel_WaitChordModifiersReleased()
            ClipAngel_ReleaseChordModifiersForSend()
            ; ControlSend avoids Shift keys $Enter → SelectClipPasteThenMinimize.
            if !ClipAngel_ConstantPaste_ControlSend(hwnd, "{Enter}") {
                stopReason := "paste send failed"
                break
            }

            ClipAngel_ConstantPaste_WaitClipboardSettle()
            if ClipAngel_ConstantPaste_IsTextOnlyClip() {
                ClipAngel_RestorePriorFocus(priorHwnd)
                ClipAngel_ReleaseChordModifiersForSend()
                Send "{Enter}"
            }

            pastedCount += 1
            StandardLoadingBar_Update("⏳ Constant Pasting… " pastedCount "  [#!+;] stop", BANNER_ACCENT_INFO)

            if !ClipAngel_ConstantPaste_WaitGap(CONSTANT_PASTE_GAP_MS) {
                stopReason := "stopped"
                break
            }

            hwnd := ClipAngel_MainHwnd()
            if !hwnd {
                stopReason := "Clip Angel closed"
                break
            }
            ; Keep CA visible; re-focus grid then advance one row.
            try WinShow("ahk_id " hwnd)
            catch {
            }
            if !ClipAngel_EnsureWindowActive(hwnd, 400) {
                stopReason := "could not focus Clip Angel"
                break
            }
            if !ClipAngel_ConstantPaste_PrepareGrid(hwnd, &root, &dataGrid) {
                stopReason := "clip list not ready"
                break
            }

            ClipAngel_ReleaseChordModifiersForSend()
            ClipAngel_ConstantPaste_ControlSend(hwnd, "{Down}")
            Sleep 40

            rowAfter := ClipAngel_ConstantPaste_GetSelectedRowName(hwnd, root)
            if (rowAfter = "" || rowAfter = rowBefore) {
                stopReason := "end of list"
                break
            }
        }
    } catch as e {
        stopReason := e.Message
    } finally {
        g_ClipAngelConstantPasteActive := false
        g_ClipAngelConstantPasteStopRequested := false
        try StandardLoadingBar_Hide(0)
        catch {
        }
        ClipAngel_CloseAndRestoreFocus(priorHwnd)
        ClipAngel_ReleaseAutomationLock()
    }

    if (pastedCount > 0 && (stopReason = "stopped" || stopReason = "end of list")) {
        msg := (stopReason = "end of list")
            ? "✅ Constant Pasting done (" pastedCount ")"
            : "✅ Constant Pasting stopped (" pastedCount ")"
        ShowCenteredOverlay_Utils(msg, 1500, BANNER_ACCENT_SUCCESS)
        return true
    }
    if (pastedCount > 0) {
        ShowCenteredOverlay_Utils("⚠ Constant Pasting ended: " stopReason " (" pastedCount ")", 2500,
            BANNER_ACCENT_INTERMEDIATE)
        return true
    }
    ShowCenteredOverlay_Utils("❌ Constant Pasting failed: " stopReason, 2500, BANNER_ACCENT_ERROR)
    return false
}
