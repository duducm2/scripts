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

; #region agent log
ClipAngel_ConstantPaste_DebugLog(hypothesisId, location, message, data := "") {
    try {
        dataJson := "{}"
        if (IsObject(data)) {
            parts := []
            for k, v in data {
                try {
                    vv := "" v
                    vv := StrReplace(vv, "\", "\\")
                    vv := StrReplace(vv, '"', '\"')
                    vv := StrReplace(vv, "`r", "")
                    vv := StrReplace(vv, "`n", " ")
                    parts.Push('"' k '":"' vv '"')
                } catch {
                }
            }
            joined := ""
            for i, p in parts
                joined .= (i = 1 ? "" : ",") p
            dataJson := "{" joined "}"
        }
        line := '{"sessionId":"688ad7","hypothesisId":"' hypothesisId '","location":"' location '","message":"' message '","data":' dataJson ',"timestamp":' A_TickCount ',"runId":"post-fix"}`n'
        FileAppend(line, A_ScriptDir "\debug-688ad7.log", "UTF-8")
    } catch {
    }
}
; #endregion

ClipAngel_ConstantPaste_IsActive() {
    global g_ClipAngelConstantPasteActive
    return !!g_ClipAngelConstantPasteActive
}

; When started from Clip Angel (always for Shift+P), pick the top z-order non-CA window.
ClipAngel_ConstantPaste_ResolveTargetHwnd() {
    prior := ClipAngel_ResolvePriorHwnd(0)
    if (prior)
        return prior
    try {
        for hwnd in WinGetList() {
            if !hwnd || !WinExist("ahk_id " hwnd)
                continue
            try {
                if (StrLower(WinGetProcessName("ahk_id " hwnd)) = "clipangel.exe")
                    continue
                if !DllCall("IsWindowVisible", "ptr", hwnd)
                    continue
                title := WinGetTitle("ahk_id " hwnd)
                if (title = "")
                    continue
                cls := WinGetClass("ahk_id " hwnd)
                if (cls = "tooltips_class32" || cls = "Shell_TrayWnd" || cls = "DV2ControlHost")
                    continue
                return hwnd
            } catch {
            }
        }
    } catch {
    }
    return 0
}

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
                sel := dataGrid.SelectionPattern.GetSelection()
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

; Advance one row after paste. ControlSend Down to the main hwnd does not move
; WinForms DataGridView selection (runtime: rowAfter stayed Row 0). Prefer Send
; while CA is focused; UIA-select next Row N+1 if selection unchanged.
ClipAngel_ConstantPaste_AdvanceSelection(hwnd, root, rowBefore) {
    ClipAngel_ReleaseChordModifiersForSend()
    Send "{Down}"
    Sleep 50
    rowAfter := ClipAngel_ConstantPaste_GetSelectedRowName(hwnd, root)
    if (rowAfter != "" && rowAfter != rowBefore)
        return rowAfter

    if !RegExMatch(rowBefore, "i)(?:Row|Linha)\s*(\d+)", &m)
        return rowAfter
    nextIdx := Integer(m[1]) + 1
    nextName := "Row " nextIdx
    dataGrid := ClipAngel_UiaGetDataGrid(hwnd, root)
    if !dataGrid
        return ClipAngel_ConstantPaste_GetSelectedRowName(hwnd, root)
    row := ClipAngel_UiaFindFirst(dataGrid, { Type: 50025, Name: nextName })
    if !row {
        try row := ClipAngel_UiaFindFirst(dataGrid, { Type: 50025, Name: "Linha " nextIdx })
        catch
            row := 0
    }
    if !row
        return ClipAngel_ConstantPaste_GetSelectedRowName(hwnd, root)
    try {
        if row.GetPropertyValue(UIA.Property.IsSelectionItemPatternAvailable)
            row.SelectionItemPattern.Select()
        else
            ClipAngel_UiaTryLegacySelectRow(row)
    } catch {
        ClipAngel_UiaTryLegacySelectRow(row)
    }
    try {
        if row.GetPropertyValue(UIA.Property.IsScrollItemPatternAvailable)
            row.ScrollItemPattern.ScrollIntoView()
    } catch {
    }
    Sleep 50
    return ClipAngel_ConstantPaste_GetSelectedRowName(hwnd, root)
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
    ; #region agent log
    ClipAngel_ConstantPaste_DebugLog("H2", "Toggle:entry", "toggle invoked", Map(
        "active", g_ClipAngelConstantPasteActive ? "1" : "0",
        "stopReq", g_ClipAngelConstantPasteStopRequested ? "1" : "0",
        "fg", WinGetProcessName("A")
    ))
    ; #endregion
    if (g_ClipAngelConstantPasteActive) {
        g_ClipAngelConstantPasteStopRequested := true
        g_ClipAngelConstantPasteActive := false
        ; #region agent log
        ClipAngel_ConstantPaste_DebugLog("H2", "Toggle:stop", "stop requested", Map())
        ; #endregion
        return
    }
    ClipAngel_ConstantPaste_Run()
}

ClipAngel_ConstantPaste_Run() {
    global g_ClipAngelConstantPasteActive, g_ClipAngelConstantPasteStopRequested
    if (g_ClipAngelConstantPasteActive)
        return false

    priorHwnd := ClipAngel_ConstantPaste_ResolveTargetHwnd()
    if (!priorHwnd) {
        ShowCenteredOverlay_Utils(
            "❌ Constant Pasting: no paste target window found.",
            2500, BANNER_ACCENT_ERROR)
        return false
    }
    if !ClipAngel_TryAcquireAutomationLock()
        return false

    g_ClipAngelConstantPasteActive := true
    g_ClipAngelConstantPasteStopRequested := false
    pastedCount := 0
    stopReason := "stopped"

    StandardLoadingBar_Show("⏳ Constant Pasting…  [Shift+P] stop", BANNER_ACCENT_INFO, {
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
                ; #region agent log
                ClipAngel_ConstantPaste_DebugLog("H3", "loop:break", "ca closed", Map("pasted", pastedCount))
                ; #endregion
                break
            }
            if !ClipAngel_ConstantPaste_PrepareGrid(hwnd, &root, &dataGrid) {
                stopReason := "clip list not ready"
                ; #region agent log
                ClipAngel_ConstantPaste_DebugLog("H3", "loop:break", "prepare failed", Map("pasted", pastedCount,
                    "phase", "pre-paste"))
                ; #endregion
                break
            }

            rowBefore := ClipAngel_ConstantPaste_GetSelectedRowName(hwnd, root)
            ; #region agent log
            ClipAngel_ConstantPaste_DebugLog("H1", "loop:iter", "iteration start", Map(
                "pasted", pastedCount, "rowBefore", rowBefore
            ))
            ; #endregion
            if (rowBefore = "") {
                stopReason := "no clip selected"
                ; #region agent log
                ClipAngel_ConstantPaste_DebugLog("H1", "loop:break", "no selection", Map("pasted", pastedCount))
                ; #endregion
                break
            }

            ClipAngel_WaitChordModifiersReleased()
            ClipAngel_ReleaseChordModifiersForSend()
            ; ControlSend avoids Shift keys $Enter → SelectClipPasteThenMinimize.
            if !ClipAngel_ConstantPaste_ControlSend(hwnd, "{Enter}") {
                stopReason := "paste send failed"
                ; #region agent log
                ClipAngel_ConstantPaste_DebugLog("H3", "loop:break", "paste send failed", Map("pasted", pastedCount))
                ; #endregion
                break
            }

            ClipAngel_ConstantPaste_WaitClipboardSettle()
            isText := ClipAngel_ConstantPaste_IsTextOnlyClip()
            ; #region agent log
            ClipAngel_ConstantPaste_DebugLog("H5", "loop:afterPaste", "paste done", Map(
                "isText", isText ? "1" : "0", "pastedNext", pastedCount + 1
            ))
            ; #endregion
            if isText {
                ClipAngel_RestorePriorFocus(priorHwnd)
                ClipAngel_ReleaseChordModifiersForSend()
                Send "{Enter}"
            }

            pastedCount += 1
            StandardLoadingBar_Update("⏳ Constant Pasting… " pastedCount "  [Shift+P] stop", BANNER_ACCENT_INFO)

            gapOk := ClipAngel_ConstantPaste_WaitGap(CONSTANT_PASTE_GAP_MS)
            ; #region agent log
            ClipAngel_ConstantPaste_DebugLog("H2", "loop:gap", "gap finished", Map(
                "gapOk", gapOk ? "1" : "0",
                "active", g_ClipAngelConstantPasteActive ? "1" : "0",
                "stopReq", g_ClipAngelConstantPasteStopRequested ? "1" : "0",
                "pasted", pastedCount
            ))
            ; #endregion
            if !gapOk {
                stopReason := "stopped"
                break
            }

            hwnd := ClipAngel_MainHwnd()
            if !hwnd {
                stopReason := "Clip Angel closed"
                ; #region agent log
                ClipAngel_ConstantPaste_DebugLog("H3", "loop:break", "ca closed after gap", Map("pasted", pastedCount))
                ; #endregion
                break
            }
            ; Keep CA visible; re-focus grid then advance one row.
            try WinShow("ahk_id " hwnd)
            catch {
            }
            if !ClipAngel_EnsureWindowActive(hwnd, 400) {
                stopReason := "could not focus Clip Angel"
                ; #region agent log
                ClipAngel_ConstantPaste_DebugLog("H4", "loop:break", "ensure active failed", Map("pasted", pastedCount))
                ; #endregion
                break
            }
            if !ClipAngel_ConstantPaste_PrepareGrid(hwnd, &root, &dataGrid) {
                stopReason := "clip list not ready"
                ; #region agent log
                ClipAngel_ConstantPaste_DebugLog("H3", "loop:break", "prepare failed", Map("pasted", pastedCount,
                    "phase", "pre-down"))
                ; #endregion
                break
            }

            ClipAngel_ReleaseChordModifiersForSend()
            rowAfter := ClipAngel_ConstantPaste_AdvanceSelection(hwnd, root, rowBefore)
            ; #region agent log
            ClipAngel_ConstantPaste_DebugLog("H1", "loop:afterDown", "selection after advance", Map(
                "rowBefore", rowBefore, "rowAfter", rowAfter, "pasted", pastedCount
            ))
            ; #endregion
            if (rowAfter = "" || rowAfter = rowBefore) {
                stopReason := "end of list"
                ; #region agent log
                ClipAngel_ConstantPaste_DebugLog("H1", "loop:break", "end of list gate", Map(
                    "rowBefore", rowBefore, "rowAfter", rowAfter, "pasted", pastedCount
                ))
                ; #endregion
                break
            }
        }
    } catch as e {
        stopReason := e.Message
        ; #region agent log
        ClipAngel_ConstantPaste_DebugLog("H5", "loop:catch", "exception", Map("err", e.Message, "pasted", pastedCount))
        ; #endregion
    } finally {
        ; #region agent log
        ClipAngel_ConstantPaste_DebugLog("H1", "Run:finally", "loop ended", Map(
            "stopReason", stopReason, "pastedCount", pastedCount
        ))
        ; #endregion
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
