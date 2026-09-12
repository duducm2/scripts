; =============================================================================
; Utils module: clip_angel_constant_paste.ahk
; Constant Pasting: toggle loop pastes from current Clip Angel selection
; (All or Favorites; no Row-0 jump), Enter only for text, 1.5s interruptible gap.
; Directions: "down" (Shift+P) / "up" (Shift+B). After paste, Clip Angel moves the
; used clip to Row 0 — bottom-up next target is former N-1 at new index N.
; Loaded via #include into Utils.ahk after clip_angel_favorite / activate.
; =============================================================================

CONSTANT_PASTE_GAP_MS := 1500
CONSTANT_PASTE_POLL_MS := 50
CONSTANT_PASTE_CLIPBOARD_WAIT_MS := 400

global g_ClipAngelConstantPasteActive := false
global g_ClipAngelConstantPasteStopRequested := false
global g_ClipAngelConstantPasteDirection := "down"
global g_ClipAngelConstantPasteStopHint := "Shift+P"

ClipAngel_ConstantPaste_IsActive() {
    global g_ClipAngelConstantPasteActive
    return !!g_ClipAngelConstantPasteActive
}

; When started from Clip Angel, pick the top z-order non-CA window.
; Skip other AHK hosts (AppLaunchers/Shift keys/Utils GUIs) — they often sit
; above the real paste target in z-order and steal Clip Angel's "previous window".
ClipAngel_ConstantPaste_IsExcludedPasteTarget(hwnd) {
    if !hwnd
        return true
    try {
        exe := StrLower(WinGetProcessName("ahk_id " hwnd))
        if (exe = "clipangel.exe")
            return true
        if (exe = "autohotkey64.exe" || exe = "autohotkey32.exe" || exe = "autohotkey.exe"
            || exe = "autohotkey64_u32.exe")
            return true
        cls := WinGetClass("ahk_id " hwnd)
        if (cls = "tooltips_class32" || cls = "Shell_TrayWnd" || cls = "DV2ControlHost"
            || cls = "Progman" || cls = "WorkerW")
            return true
        title := WinGetTitle("ahk_id " hwnd)
        if (title = "")
            return true
    } catch {
        return true
    }
    return false
}

ClipAngel_ConstantPaste_ResolveTargetHwnd() {
    prior := ClipAngel_ResolvePriorHwnd(0)
    if (prior && !ClipAngel_ConstantPaste_IsExcludedPasteTarget(prior))
        return prior
    try {
        for hwnd in WinGetList() {
            if !hwnd || !WinExist("ahk_id " hwnd)
                continue
            try {
                if ClipAngel_ConstantPaste_IsExcludedPasteTarget(hwnd)
                    continue
                if !DllCall("IsWindowVisible", "ptr", hwnd)
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

ClipAngel_ConstantPaste_ParseRowIndex(rowName) {
    if (rowName = "" || !RegExMatch(rowName, "i)(?:Row|Linha)\s*(\d+)", &m))
        return -1
    return Integer(m[1])
}

; Select DataGrid row by index after list reorder. Returns selected row name or "".
ClipAngel_ConstantPaste_SelectRowByIndex(hwnd, root, idx) {
    if !(idx is Integer) || idx < 0 || !hwnd
        return ""
    dataGrid := ClipAngel_UiaGetDataGrid(hwnd, root)
    if !dataGrid
        return ""
    row := ClipAngel_UiaFindFirst(dataGrid, { Type: 50025, Name: "Row " idx })
    if !row {
        try row := ClipAngel_UiaFindFirst(dataGrid, { Type: 50025, Name: "Linha " idx })
        catch
            row := 0
    }
    if !row
        return ""
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
    ClipAngel_UiaEnsureGridListFocus(dataGrid, hwnd, root)
    return ClipAngel_ConstantPaste_GetSelectedRowName(hwnd, root)
}

; Next row after paste. Clip Angel "move to top after use" remaps indices:
;   pasted at old N → new 0; old 0..N-1 → new 1..N; old N+1.. stay at N+1..
; Bottom-up (toward Row 0): want former N-1, now at index N → target = N.
; Top-down: want former N+1, still at N+1 → target = N+1.
; Returns new selected row name, or "" when traversal should stop.
ClipAngel_ConstantPaste_AdvanceSelection(hwnd, root, rowBefore, direction := "down") {
    idx := ClipAngel_ConstantPaste_ParseRowIndex(rowBefore)
    if (idx < 0)
        return ""

    goUp := (direction = "up")
    if goUp {
        ; Already at top before paste → nothing above after move-to-top.
        if (idx <= 0)
            return ""
        targetIdx := idx  ; former (idx-1) shifted down into this slot
    } else {
        targetIdx := idx + 1
    }

    ClipAngel_ReleaseChordModifiersForSend()
    rowAfter := ClipAngel_ConstantPaste_SelectRowByIndex(hwnd, root, targetIdx)
    if (ClipAngel_ConstantPaste_ParseRowIndex(rowAfter) = targetIdx)
        return rowAfter

    ; Fallback: arrow from current selection (may be Row 0 after paste).
    Send(goUp ? "{Up}" : "{Down}")
    Sleep 50
    rowAfter := ClipAngel_ConstantPaste_GetSelectedRowName(hwnd, root)
    if (ClipAngel_ConstantPaste_ParseRowIndex(rowAfter) = targetIdx)
        return rowAfter
    return ""
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

ClipAngel_ConstantPaste_ToggleDown(*) {
    ClipAngel_ConstantPaste_Toggle("down")
}

ClipAngel_ConstantPaste_ToggleUp(*) {
    ClipAngel_ConstantPaste_Toggle("up")
}

; Either Shift+P or Shift+B stops an active run; direction applies only when starting.
ClipAngel_ConstantPaste_Toggle(direction := "down") {
    global g_ClipAngelConstantPasteActive, g_ClipAngelConstantPasteStopRequested
    if (g_ClipAngelConstantPasteActive) {
        g_ClipAngelConstantPasteStopRequested := true
        g_ClipAngelConstantPasteActive := false
        return
    }
    ClipAngel_ConstantPaste_Run(direction)
}

ClipAngel_ConstantPaste_Run(direction := "down") {
    global g_ClipAngelConstantPasteActive, g_ClipAngelConstantPasteStopRequested
    global g_ClipAngelConstantPasteDirection, g_ClipAngelConstantPasteStopHint
    if (g_ClipAngelConstantPasteActive)
        return false

    direction := (direction = "up") ? "up" : "down"
    stopHint := (direction = "up") ? "Shift+B" : "Shift+P"
    dirLabel := (direction = "up") ? "↑" : "↓"

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
    g_ClipAngelConstantPasteDirection := direction
    g_ClipAngelConstantPasteStopHint := stopHint
    pastedCount := 0
    stopReason := "stopped"

    StandardLoadingBar_Show("⏳ Constant Pasting " dirLabel "…  [" stopHint "] stop", BANNER_ACCENT_INFO, {
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
            ; Prime Clip Angel's "previous window", then Send Enter (same as Alt+1).
            ; ControlSend Enter rearranges the list but does not paste into the target.
            ClipAngel_RestorePriorFocus(priorHwnd)
            if !ClipAngel_EnsureWindowActive(hwnd, 400) {
                stopReason := "could not focus Clip Angel for paste"
                break
            }
            if !ClipAngel_ConstantPaste_PrepareGrid(hwnd, &root, &dataGrid) {
                stopReason := "clip list not ready"
                break
            }
            ClipAngel_ReleaseChordModifiersForSend()
            Send "{Enter}"

            ClipAngel_ConstantPaste_WaitClipboardSettle()
            if ClipAngel_ConstantPaste_IsTextOnlyClip() {
                ClipAngel_RestorePriorFocus(priorHwnd)
                ClipAngel_ReleaseChordModifiersForSend()
                Send "{Enter}"
            }

            pastedCount += 1
            StandardLoadingBar_Update("⏳ Constant Pasting " dirLabel "… " pastedCount "  [" stopHint "] stop",
                BANNER_ACCENT_INFO)

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

            rowAfter := ClipAngel_ConstantPaste_AdvanceSelection(hwnd, root, rowBefore, direction)
            ; "" = end / failed advance. Do not compare names: bottom-up target can still be "Row N".
            if (rowAfter = "") {
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
