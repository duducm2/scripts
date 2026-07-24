; =============================================================================
; Utils module: clip_angel_favorite.ahk
; Clip Angel mark favorite and related flows
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; Clip Angel: Mark Last Clip as Favorite
; =============================================================================
; Wait after clipboard change before favoriting newest clip (copy / dictation ingest).
CLIPANGEL_PRE_FAVORITE_INGEST_DELAY_MS := 400
; Settle after row focus, before Alt+Q (all favorite paths).
CLIPANGEL_FAVORITE_UI_SETTLE_MS := 50
; Bounded poll after native Alt+P open before favoriting (cold start can exceed fixed sleeps).
CLIPANGEL_FAVORITE_OPEN_READY_MS := 1200
CLIPANGEL_INCREMENTAL_PASTE_SETTLE_MS := 250
CLIPANGEL_SEQUENTIAL_PASTE_GAP_MS := 300
CLIPANGEL_UIA_POLL_MS := 30
CLIPANGEL_GRID_WAIT_MS := 400
CLIPANGEL_ROW0_WAIT_MS := 300
CLIPANGEL_ROW0_SELECT_WAIT_MS := 250
CLIPANGEL_MIN_LAYOUT_WIDTH := 300
CLIPANGEL_MIN_LAYOUT_HEIGHT := 200
CLIPANGEL_ALT_P_SETTLE_MS := 200
CLIPANGEL_ALT_P_HWND_WAIT_MS := 800
global g_ClipAngelAutomationBusy := false

; Shortcut flow (matches app): open Clip Angel, ensure list focus (not Window tab),
; select first or last grid row, Send Alt+Q. Optional: target "last" for bottom row.
; UIA-v2 FindFirst throws TargetError when nothing matches - never chain with if !c without try.
ClipAngel_UiaFindFirst(root, conditions) {
    if !root
        return 0
    try return root.FindFirst(conditions)
    catch
        return 0
}

ClipAngel_FindFavoriteCell(row) {
    if !row
        return 0
    rn := ""
    try rn := row.Name
    catch {
        rn := ""
    }
    suffix := "0"
    if RegExMatch(rn, "i)(?:Row|Linha)\s*(\d+)", &m)
        suffix := m[1]
    else if RegExMatch(rn, "(\d+)\s*$", &m)
        suffix := m[1]
    ; EN + PT-BR column headers seen in Clip Angel / localized WinForms.
    for cand in [
        "Favorite Row " . suffix, "Favourite Row " . suffix, "Favorito Row " . suffix,
        "Favorite Linha " . suffix, "Favorito Linha " . suffix
    ] {
        c := ClipAngel_UiaFindFirst(row, { Type: UIA.Type.CheckBox, Name: cand })
        if c
            return c
        c := ClipAngel_UiaFindFirst(row, { Type: 50002, Name: cand })
        if c
            return c
    }
    try {
        for c in row.FindAll({ Type: 50002 }) {
            try n := c.Name
            catch
                continue
            if RegExMatch(n, "i)favorite|favourite|favorito")
                return c
        }
        boxes := row.FindAll({ Type: 50002 })
        if boxes.Length >= 2
            return boxes[boxes.Length]
    } catch {
    }
    return 0
}

ClipAngel_FavoriteCellIsOn(cell) {
    if !cell
        return false
    try {
        if cell.GetPropertyValue(UIA.Property.IsTogglePatternAvailable)
            return cell.TogglePattern.ToggleState = UIA.ToggleState.On
        ts := cell.GetPropertyValue(UIA.Property.ToggleToggleState)
        if ts != ""
            return ts = UIA.ToggleState.On
    } catch {
    }
    ; Value only for read-only grid cells - Legacy CHECKED (0x10) often false-positives on DataGrid cells.
    try {
        v := cell.Value
        if (v = "true" || v = "True" || v = "1")
            return true
    } catch {
    }
    return false
}

ClipAngel_MainHwnd() {
    h := WinExist("ClipAngel")
    if h
        return h
    return WinExist("ahk_exe ClipAngel.exe")
}

ClipAngel_IsWindowShown(hwnd) {
    if !hwnd || !WinExist("ahk_id " hwnd)
        return false
    try {
        if WinGetMinMax("ahk_id " hwnd) = -1
            return false
    } catch {
        return false
    }
    if DllCall("IsIconic", "ptr", hwnd)
        return false
    return DllCall("IsWindowVisible", "ptr", hwnd)
}

; True when minimized, invisible, or shrunk to a tiny bar (Alt+P toggle can leave a 1px-sized window).
ClipAngel_NeedsLayoutCorrection(hwnd) {
    if !hwnd || !WinExist("ahk_id " hwnd)
        return false
    if !ClipAngel_IsWindowShown(hwnd)
        return true
    try {
        WinGetPos(, , &w, &h, "ahk_id " hwnd)
        if (w > 0 && h > 0 && (w < CLIPANGEL_MIN_LAYOUT_WIDTH || h < CLIPANGEL_MIN_LAYOUT_HEIGHT))
            return true
    } catch {
        return true
    }
    return false
}

; AHK fallback after native Alt+P: restore, show, move/maximize when needed, optionally activate.
ClipAngel_EnsureVisibleAndLayout(hwnd, targetMon := 0, activate := true) {
    if !hwnd
        return false
    try {
        mm := WinGetMinMax("ahk_id " hwnd)
        if (mm = -1 || mm = 1)
            WinRestore("ahk_id " hwnd)
    } catch {
    }
    try WinShow("ahk_id " hwnd)
    catch {
    }
    if ClipAngel_NeedsLayoutCorrection(hwnd)
        ClipAngel_ApplyLayoutOnMonitor(hwnd, targetMon)
    return activate ? ClipAngel_EnsureWindowActive(hwnd) : true
}

ClipAngel_ShowWindow(hwnd) {
    if !hwnd
        return false
    if WinActive("ahk_id " hwnd) && ClipAngel_IsWindowShown(hwnd) && !ClipAngel_NeedsLayoutCorrection(hwnd)
        return true
    return ClipAngel_EnsureVisibleAndLayout(hwnd, 0, true)
}

; Minimize Clip Angel (process stays running; no native Alt+P / WinClose).
ClipAngel_HideWindow(hwnd) {
    if !hwnd || !WinExist("ahk_id " hwnd)
        return false
    try {
        WinMinimize("ahk_id " hwnd)
    } catch {
        return false
    }
    Sleep 50
    if !ClipAngel_IsWindowShown(hwnd)
        return true
    try {
        WinMinimize("ahk_id " hwnd)
    } catch {
        return false
    }
    Sleep 50
    return !ClipAngel_IsWindowShown(hwnd)
}

ClipAngel_WaitForMainHwnd(timeoutMs := CLIPANGEL_ALT_P_HWND_WAIT_MS) {
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        if hwnd := ClipAngel_MainHwnd()
            return hwnd
        Sleep CLIPANGEL_UIA_POLL_MS
    }
    return ClipAngel_MainHwnd()
}

; dataGridView (Type 50036, AutomationId dataGridView) — clip-angel.txt.
; Pass root when caller already has UIA.ElementFromHandle(hwnd) to avoid duplicate COM round-trips.
ClipAngel_UiaGetDataGrid(hwnd, root := 0) {
    if !hwnd
        return 0
    if !root {
        root := UIA.ElementFromHandle(hwnd)
        if !root
            return 0
    }
    return ClipAngel_UiaFindFirst(root, { Type: 50036, AutomationId: "dataGridView" })
}

ClipAngel_WaitForDataGrid(hwnd, timeoutMs := CLIPANGEL_GRID_WAIT_MS, root := 0) {
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        dataGrid := ClipAngel_UiaGetDataGrid(hwnd, root)
        if dataGrid
            return dataGrid
        if !root {
            root := UIA.ElementFromHandle(hwnd)
            if !root
                return 0
        }
        Sleep CLIPANGEL_UIA_POLL_MS
    }
    return 0
}

ClipAngel_WaitForRow0(dataGrid, timeoutMs := CLIPANGEL_ROW0_WAIT_MS) {
    if !dataGrid
        return 0
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        row0 := ClipAngel_UiaFindFirst(dataGrid, { Type: 50025, Name: "Row 0" })
        if row0
            return row0
        Sleep CLIPANGEL_UIA_POLL_MS
    }
    return 0
}

; Invoke List menu when Window tab has focus and dataGridView is missing.
ClipAngel_EnsureListView(hwnd, root := 0) {
    if !hwnd
        return false
    if !root {
        root := UIA.ElementFromHandle(hwnd)
        if !root
            return false
    }
    listItem := ClipAngel_UiaFindFirst(root, { Type: 50011, Name: "List" })
    if !listItem
        return false
    try {
        if listItem.GetPropertyValue(UIA.Property.IsInvokePatternAvailable)
            listItem.InvokePattern.Invoke()
        else
            listItem.SetFocus()
        return true
    } catch {
        return false
    }
}

ClipAngel_UiaInvokeElement(el) {
    if !el
        return false
    try {
        if el.GetPropertyValue(UIA.Property.IsInvokePatternAvailable) {
            el.InvokePattern.Invoke()
            return true
        }
    } catch {
    }
    try {
        el.Click()
        return true
    } catch {
    }
    try {
        el.SetFocus()
        Sleep 40
        Send "{Enter}"
        return true
    } catch {
    }
    return false
}

ClipAngel_UiaOpenSubmenu(menuItem) {
    if !menuItem
        return false
    try {
        if menuItem.GetPropertyValue(UIA.Property.IsExpandCollapsePatternAvailable) {
            pat := menuItem.ExpandCollapsePattern
            if pat.ExpandCollapseState = UIA.ExpandCollapseState.Collapsed {
                pat.Expand()
                return true
            }
        }
    } catch {
    }
    try {
        menuItem.SetFocus()
        Sleep 40
        Send "{Right}"
        return true
    } catch {
    }
    return ClipAngel_UiaInvokeElement(menuItem)
}

ClipAngel_UiaFindMenuItem(roots, conditions) {
    for r in roots {
        if !r
            continue
        el := ClipAngel_UiaFindFirst(r, conditions)
        if el
            return el
    }
    return 0
}

ClipAngel_InvokePasteEnterViaKeyboard() {
    Send "{Alt}"
    Sleep 80
    Send "c"
    Sleep 80
    Send "{Right}"
    Sleep 80
    Send "{Down 3}"
    Sleep 80
    Send "{Enter}"
    return true
}

; Clip > Paste > Paste file — paste selected clip as a file into the prior app (e.g. Explorer).
ClipAngel_InvokePasteEnter(hwnd := 0) {
    if !hwnd
        hwnd := ClipAngel_MainHwnd()
    if !hwnd
        return false
    try {
        root := UIA.ElementFromHandle(hwnd)
        if !root
            return ClipAngel_InvokePasteEnterViaKeyboard()
        clipMenu := ClipAngel_UiaFindFirst(root, { Type: 50011, Name: "Clip" })
        if !clipMenu
            return ClipAngel_InvokePasteEnterViaKeyboard()
        if !ClipAngel_UiaInvokeElement(clipMenu)
            return ClipAngel_InvokePasteEnterViaKeyboard()
        Sleep 120
        desktop := 0
        try desktop := UIA.GetRootElement()
        searchRoots := [root]
        if desktop
            searchRoots.Push(desktop)
        pasteItem := ClipAngel_UiaFindMenuItem(searchRoots, { Type: 50011, Name: "Paste" })
        if !pasteItem
            return ClipAngel_InvokePasteEnterViaKeyboard()
        if !ClipAngel_UiaOpenSubmenu(pasteItem)
            return ClipAngel_InvokePasteEnterViaKeyboard()
        Sleep 120
        pasteFile := ClipAngel_UiaFindFirst(pasteItem, { Type: 50011, Name: "Paste file" })
        if !pasteFile
            pasteFile := ClipAngel_UiaFindMenuItem(searchRoots, { Type: 50011, Name: "Paste file" })
        if !pasteFile
            return ClipAngel_InvokePasteEnterViaKeyboard()
        return ClipAngel_UiaInvokeElement(pasteFile)
    } catch {
        return ClipAngel_InvokePasteEnterViaKeyboard()
    }
}

ClipAngel_UiaResolveRow0(dataGrid) {
    if !dataGrid
        return 0
    row0 := ClipAngel_UiaFindFirst(dataGrid, { Type: 50025, Name: "Row 0" })
    if row0
        return row0
    try {
        rows := dataGrid.FindAll({ Type: 50025 })
        if rows && rows.Length >= 1
            return rows[1]
    } catch {
    }
    return ClipAngel_WaitForRow0(dataGrid)
}

ClipAngel_UiaGridHasSelectionPattern(dataGrid) {
    if !dataGrid
        return false
    try return dataGrid.GetPropertyValue(UIA.Property.IsSelectionPatternAvailable)
    catch
        return false
}

ClipAngel_UiaWaitRow0Selected(row0, dataGrid, timeoutMs := CLIPANGEL_ROW0_SELECT_WAIT_MS) {
    if !row0
        return false
    gridHasSel := ClipAngel_UiaGridHasSelectionPattern(dataGrid)
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        if ClipAngel_UiaRow0IsSelected(row0, dataGrid, gridHasSel)
            return true
        Sleep CLIPANGEL_UIA_POLL_MS
    }
    return false
}

ClipAngel_UiaGridSelectionIncludesRow0(dataGrid) {
    if !dataGrid
        return false
    try {
        if dataGrid.GetPropertyValue(UIA.Property.IsSelectionPatternAvailable) {
            for item in dataGrid.SelectionPattern.GetSelection() {
                try n := item.Name
                catch
                    continue
                if RegExMatch(n, "i)^Row\s*0")
                    return true
            }
        }
    } catch {
    }
    return false
}

ClipAngel_UiaRowLegacySelected(row) {
    if !row
        return false
    try {
        state := row.GetPropertyValue(UIA.Property.LegacyIAccessibleState)
        if (state & 0x2)
            return true
    } catch {
    }
    return false
}

; WinForms DataGridView: cheap property checks first; grid SelectionPattern only when available.
ClipAngel_UiaRow0IsSelected(row0, dataGrid := 0, gridHasSelectionPattern := false) {
    if row0 {
        try {
            if row0.GetPropertyValue(UIA.Property.SelectionItemIsSelected)
                return true
        } catch {
        }
        if ClipAngel_UiaRowLegacySelected(row0)
            return true
    }
    if gridHasSelectionPattern && dataGrid && ClipAngel_UiaGridSelectionIncludesRow0(dataGrid)
        return true
    return false
}

ClipAngel_UiaTryLegacySelectRow(row) {
    if !row
        return false
    try {
        if !row.GetPropertyValue(UIA.Property.IsLegacyIAccessiblePatternAvailable)
            return false
        row.LegacyIAccessiblePattern.Select(3)
        return true
    } catch {
        return false
    }
}

; F10 toggles list vs preview — only send when preview pane has focus, not when grid already focused.
ClipAngel_UiaEnsureGridListFocus(dataGrid, hwnd, root := 0) {
    if !dataGrid
        return false
    try dataGrid.SetFocus()
    catch {
    }
    deadline := A_TickCount + 200
    while (A_TickCount < deadline) {
        try {
            if dataGrid.HasKeyboardFocus
                return true
        } catch {
        }
        Sleep CLIPANGEL_UIA_POLL_MS
    }
    if !hwnd
        return true
    if !root {
        root := UIA.ElementFromHandle(hwnd)
        if !root
            return true
    }
    preview := ClipAngel_UiaFindFirst(root, { AutomationId: "richTextBox" })
    previewFocused := false
    try previewFocused := preview && preview.HasKeyboardFocus
    catch {
    }
    if previewFocused {
        ClipAngel_ReleaseChordModifiersForSend()
        Send "{F10}"
        deadline := A_TickCount + 150
        while (A_TickCount < deadline) {
            try {
                if dataGrid.HasKeyboardFocus
                    return true
            } catch {
            }
            Sleep CLIPANGEL_UIA_POLL_MS
        }
    } else {
        try dataGrid.Click()
        catch {
        }
        try dataGrid.SetFocus()
        catch {
        }
    }
    return true
}

; First list row (Row 0 / rows[1]). force=false skips work when Row 0 is already selected.
ClipAngel_UiaEnsureRow0Selected(hwnd, force := false) {
    if !hwnd
        return false
    listInvoked := false
    try {
        root := UIA.ElementFromHandle(hwnd)
        if !root
            return false
        dataGrid := ClipAngel_UiaGetDataGrid(hwnd, root)
        if !dataGrid {
            listInvoked := ClipAngel_EnsureListView(hwnd, root)
            dataGrid := ClipAngel_WaitForDataGrid(hwnd, CLIPANGEL_GRID_WAIT_MS, root)
            if !dataGrid
                return false
        }
        row0 := ClipAngel_UiaResolveRow0(dataGrid)
        if !row0 && !listInvoked {
            ClipAngel_EnsureListView(hwnd, root)
            dataGrid := ClipAngel_WaitForDataGrid(hwnd, CLIPANGEL_GRID_WAIT_MS, root)
            if dataGrid
                row0 := ClipAngel_UiaResolveRow0(dataGrid)
        }
        if !row0
            return false
        gridHasSel := ClipAngel_UiaGridHasSelectionPattern(dataGrid)
        if !force && ClipAngel_UiaRow0IsSelected(row0, dataGrid, gridHasSel)
            return true
        try {
            if row0.GetPropertyValue(UIA.Property.IsScrollItemPatternAvailable)
                row0.ScrollItemPattern.ScrollIntoView()
        } catch {
        }
        if ClipAngel_UiaTryLegacySelectRow(row0) && ClipAngel_UiaWaitRow0Selected(row0, dataGrid,
            CLIPANGEL_ROW0_SELECT_WAIT_MS)
            return true
        try {
            if row0.GetPropertyValue(UIA.Property.IsSelectionItemPatternAvailable)
                row0.SelectionItemPattern.Select()
        } catch {
            try row0.SetFocus()
        }
        if ClipAngel_UiaWaitRow0Selected(row0, dataGrid, CLIPANGEL_ROW0_SELECT_WAIT_MS)
            return true
        ClipAngel_UiaEnsureGridListFocus(dataGrid, hwnd, root)
        ClipAngel_ReleaseChordModifiersForSend()
        Send "^{Home}"
        if ClipAngel_UiaWaitRow0Selected(row0, dataGrid, CLIPANGEL_ROW0_SELECT_WAIT_MS)
            return true
        ClipAngel_ReleaseChordModifiersForSend()
        Send "{Home}"
        if ClipAngel_UiaWaitRow0Selected(row0, dataGrid, CLIPANGEL_ROW0_SELECT_WAIT_MS)
            return true
        rn := ""
        try rn := row0.Name
        catch {
            rn := ""
        }
        titleCell := ClipAngel_UiaFindFirst(row0, { Type: 50006, Name: "Title " rn })
        if !titleCell && rn != "Row 0"
            titleCell := ClipAngel_UiaFindFirst(row0, { Type: 50006, Name: "Title Row 0" })
        if titleCell {
            try {
                if titleCell.GetPropertyValue(UIA.Property.IsLegacyIAccessiblePatternAvailable)
                    titleCell.LegacyIAccessiblePattern.Select(3)
                else
                    titleCell.Click()
                if ClipAngel_UiaWaitRow0Selected(row0, dataGrid, CLIPANGEL_ROW0_SELECT_WAIT_MS)
                    return true
            } catch {
            }
        }
        try row0.SetFocus()
        catch {
        }
        return ClipAngel_UiaWaitRow0Selected(row0, dataGrid, CLIPANGEL_ROW0_SELECT_WAIT_MS)
    } catch {
        return false
    }
}

; Macro hotkeys use Ctrl+Alt+Win - if those keys are still down, Send "!q" is not plain Alt+Q (Win+Alt+... hijacks it).
ClipAngel_ReleaseChordModifiersForSend() {
    SendInput "{LWin up}{RWin up}{LControl up}{RControl up}{LAlt up}{RAlt up}{LShift up}{RShift up}"
}

; Wait for physical release (KeyWait) then synthetic up - chord hotkeys often leave keys logically down.
ClipAngel_WaitChordModifiersReleased() {
    tw := "T0.45"
    KeyWait "Ctrl", tw
    KeyWait "Alt", tw
    KeyWait "Shift", tw
    KeyWait "LWin", tw
    KeyWait "RWin", tw
}

; Legacy native Alt+V send — All Clips view in MergeNonFavoriteClips only.
; Open/paste flows use ClipAngel_ShowWindow or Alt+P (+ Enter) instead.
ClipAngel_SendToggleHotkey() {
    ClipAngel_WaitChordModifiersReleased()
    ClipAngel_ReleaseChordModifiersForSend()
    Send "!v"
}

ClipAngel_EnsureWindowActive(hwnd, timeoutMs := 800) {
    if !hwnd
        return false
    if WinActive("ahk_id " hwnd)
        return true
    endTick := A_TickCount + timeoutMs
    loop 3 {
        try WinActivate("ahk_id " hwnd)
        catch
            return false
        remaining := endTick - A_TickCount
        if (remaining <= 0)
            break
        if WinWaitActive("ahk_id " hwnd, , Max(0.05, remaining / 1000.0))
            return true
        Sleep 50
    }
    return WinActive("ahk_id " hwnd)
}

; True when outer size is a tiny bar (Alt+P toggle) — ignores minimized (caller decides).
ClipAngel_IsTinyLayout(hwnd) {
    if !hwnd || !WinExist("ahk_id " hwnd)
        return false
    try {
        WinGetPos(, , &w, &h, "ahk_id " hwnd)
        if (w > 0 && h > 0 && (w < CLIPANGEL_MIN_LAYOUT_WIDTH || h < CLIPANGEL_MIN_LAYOUT_HEIGHT))
            return true
    } catch {
        return true
    }
    return false
}

; Move to monitor work area and maximize; re-activate if layout steals focus.
; Same-monitor warm path: maximize/activate only (skip MoveWindowToMonitor + Sleep 80).
ClipAngel_ApplyLayoutOnMonitor(hwnd, targetMon := 0) {
    if !hwnd
        return false
    if (!targetMon || targetMon < 1 || targetMon > MonitorGetCount()) {
        try targetMon := GetAhkMonitorIndexFromHwnd(WinGetID("A"))
        catch
            targetMon := 0
    }
    if (!targetMon) {
        try targetMon := MonitorGetPrimary()
        catch
            targetMon := 1
    }
    curMon := 0
    try curMon := GetAhkMonitorIndexFromHwnd(hwnd)
    catch
        curMon := 0
    sameMon := (curMon >= 1 && curMon = targetMon)
    mm := 0
    try mm := WinGetMinMax("ahk_id " hwnd)
    catch
        mm := 0
    iconic := (mm = -1) || DllCall("IsIconic", "ptr", hwnd)

    if (sameMon) {
        ; Already maximized on target — activate only.
        if (mm = 1 && !ClipAngel_IsTinyLayout(hwnd))
            return ClipAngel_EnsureWindowActive(hwnd, 150)
        try {
            if (iconic || mm = 1)
                WinRestore("ahk_id " hwnd)
        } catch {
        }
        try WinShow("ahk_id " hwnd)
        catch {
        }
        ; Tiny after restore → full move path below.
        if (!ClipAngel_IsTinyLayout(hwnd)) {
            TryMaximizeWindow(hwnd)
            return ClipAngel_EnsureWindowActive(hwnd, 150)
        }
    }

    MoveWindowToMonitor(hwnd, targetMon)
    TryMaximizeWindow(hwnd)
    return ClipAngel_EnsureWindowActive(hwnd)
}

ClipAngel_IsListReady(&outHwnd := 0) {
    outHwnd := ClipAngel_MainHwnd()
    if !outHwnd
        return false
    dataGrid := ClipAngel_UiaGetDataGrid(outHwnd)
    if !dataGrid
        return false
    return ClipAngel_UiaFindFirst(dataGrid, { Type: 50025, Name: "Row 0" }) ? true : false
}

; Poll until dataGridView + Row 0 exist (e.g. after native Alt+P open). Retries row-0 selection at end.
; activateOnRetry: false when paste target must keep foreground (#!+1 / ^!b).
ClipAngel_WaitForListReady(timeoutMs := CLIPANGEL_FAVORITE_OPEN_READY_MS, activateOnRetry := true) {
    deadline := A_TickCount + timeoutMs
    hwnd := 0
    while (A_TickCount < deadline) {
        if ClipAngel_IsListReady(&hwnd) {
            ClipAngel_UiaEnsureRow0Selected(hwnd, false)
            return true
        }
        if hwnd := ClipAngel_MainHwnd()
            ClipAngel_EnsureVisibleAndLayout(hwnd, 0, activateOnRetry)
        Sleep CLIPANGEL_UIA_POLL_MS
    }
    if hwnd := ClipAngel_MainHwnd() {
        ClipAngel_EnsureVisibleAndLayout(hwnd, 0, activateOnRetry)
        ClipAngel_UiaEnsureRow0Selected(hwnd, true)
        return ClipAngel_IsListReady()
    }
    return false
}

ClipAngel_ResolvePriorHwnd(priorHwnd := 0) {
    if (priorHwnd && WinExist("ahk_id " priorHwnd))
        return priorHwnd
    try {
        activeHwnd := WinGetID("A")
        if (activeHwnd && !WinActive("ahk_exe ClipAngel.exe"))
            return activeHwnd
    } catch {
    }
    return 0
}

ClipAngel_RestorePriorFocus(priorHwnd) {
    if (!priorHwnd || !WinExist("ahk_id " priorHwnd))
        return
    try {
        WinActivate("ahk_id " priorHwnd)
        WinWaitActive("ahk_id " priorHwnd, , 2)
    } catch {
    }
}

ClipAngel_TryAcquireAutomationLock() {
    global g_ClipAngelAutomationBusy
    if (g_ClipAngelAutomationBusy) {
        ShowCenteredOverlay_Utils("⏳ Clip Angel busy.", 1200, BANNER_ACCENT_INTERMEDIATE)
        return false
    }
    g_ClipAngelAutomationBusy := true
    return true
}

ClipAngel_ReleaseAutomationLock() {
    global g_ClipAngelAutomationBusy
    g_ClipAngelAutomationBusy := false
}

ClipAngel_EnsureOpenAndReady(silent := true) {
    if !ActivateClipAngelWithFocusCorrection(silent)
        return false
    return ClipAngel_IsListReady()
}

ClipAngel_SendIncrementalPaste() {
    ClipAngel_WaitChordModifiersReleased()
    ClipAngel_ReleaseChordModifiersForSend()
    SendInput "^!b"
    Sleep(CLIPANGEL_INCREMENTAL_PASTE_SETTLE_MS)
}

ClipAngel_CloseAndRestoreFocus(priorHwnd := 0) {
    if hwnd := ClipAngel_MainHwnd()
        ClipAngel_HideWindow(hwnd)
    ClipAngel_RestorePriorFocus(priorHwnd)
}

; True when Enter on the clip list pastes into the prior app (not filter combo / preview edit).
ClipAngel_IsListPasteEnterContext(hwnd := 0) {
    if !hwnd
        hwnd := ClipAngel_MainHwnd()
    if !hwnd
        return false
    try {
        focused := UIA.GetFocusedElement()
        if !focused
            return true
        type := 0
        aid := ""
        name := ""
        try type := focused.Type
        catch {
        }
        try aid := focused.AutomationId
        catch {
        }
        try name := focused.Name
        catch {
        }
        ; DataGrid row Edit cells (e.g. VisualWeight Row N) — Enter still pastes the clip.
        if (type = 50004 && RegExMatch(name, "i)(?:Row|Linha)\s*\d+"))
            return true
        ; ComboBox / Edit / Document — filters, search, preview, edit-clip UI.
        if (type = 50003 || type = 50004 || type = 50030)
            return false
        if (aid = "MarkFilter" || aid = "TypeFilter" || aid = "comboBoxSearchString"
            || aid = "richTextBox" || aid = "urlTextBox")
            return false
        if (type = 50036 || type = 50025)  ; Table / DataGrid row
            return true
        if RegExMatch(name, "i)^(?:Row|Linha)\s*\d+")
            return true
        if (aid = "dataGridView")
            return true
        ; Title/image cells inside a row still paste via Enter.
        if (type = 50006 && RegExMatch(name, "i)(?:Row|Linha)\s*\d+"))
            return true
    } catch {
    }
    return true
}

; Down N (optional), native Enter to paste selected clip, then minimize Clip Angel.
ClipAngel_SelectClipPasteThenMinimize(downCount := 0) {
    ; Alt+1–5 leave Alt logically down; without release, Enter becomes Alt+Enter (paste file).
    ClipAngel_WaitChordModifiersReleased()
    ClipAngel_ReleaseChordModifiersForSend()
    hwnd := ClipAngel_MainHwnd()
    if (hwnd) {
        dataGrid := ClipAngel_UiaGetDataGrid(hwnd)
        if (dataGrid)
            ClipAngel_UiaEnsureGridListFocus(dataGrid, hwnd)
    }
    if (downCount > 0) {
        loop downCount
            Send "{Down}"
    }
    Send "{Enter}"
    Sleep 100
    ClipAngel_CloseAndRestoreFocus(0)
}

; Down N (optional), F10 → select-all → copy preview text, then minimize Clip Angel.
ClipAngel_SelectClipCopyThenMinimize(downCount := 0) {
    ClipAngel_WaitChordModifiersReleased()
    ClipAngel_ReleaseChordModifiersForSend()
    hwnd := ClipAngel_MainHwnd()
    if (hwnd) {
        dataGrid := ClipAngel_UiaGetDataGrid(hwnd)
        if (dataGrid)
            ClipAngel_UiaEnsureGridListFocus(dataGrid, hwnd)
    }
    if (downCount > 0) {
        loop downCount
            Send "{Down}"
        Sleep 150
    }
    Send "{F10}"
    Sleep 150
    Send "^a"
    Sleep 150
    Send "^c"
    Sleep 100
    ClipAngel_CloseAndRestoreFocus(0)
}

; Native open + row 0: release chord modifiers, Alt+P, then AHK ShowWindow/layout fallback + ^Home.
; Alt+P alone is unreliable; EnsureVisibleAndLayout restores a usable window when toggle leaves it tiny.
ClipAngel_ActivateNativeFirstClip(priorHwnd := 0) {
    ClipAngel_WaitChordModifiersReleased()
    ClipAngel_ReleaseChordModifiersForSend()
    if (priorHwnd)
        ClipAngel_EnsureWindowActive(priorHwnd)
    SendInput "{Alt up}{Shift up}{Win up}{Ctrl up}"
    Sleep CLIPANGEL_ALT_P_SETTLE_MS
    SendInput "!p"
    Sleep CLIPANGEL_ALT_P_SETTLE_MS
    targetMon := 0
    if (priorHwnd) {
        try targetMon := GetAhkMonitorIndexFromHwnd(priorHwnd)
        catch
            targetMon := 0
    }
    if (hwnd := ClipAngel_WaitForMainHwnd())
        ClipAngel_EnsureVisibleAndLayout(hwnd, targetMon, true)
    SendInput "^{Home}"
    Sleep CLIPANGEL_ALT_P_SETTLE_MS
    ClipAngel_ReleaseChordModifiersForSend()
    ; Paste flows: return focus to target before incremental paste (^!b).
    if (priorHwnd)
        ClipAngel_EnsureWindowActive(priorHwnd)
}

; Native top-item paste: open via ActivateNativeFirstClip, wait for grid, then incremental paste (^!b).
ClipAngel_SendNativeTopItemKeys(priorHwnd := 0) {
    ClipAngel_ActivateNativeFirstClip(priorHwnd)
    ClipAngel_WaitForListReady(CLIPANGEL_FAVORITE_OPEN_READY_MS, false)
    ClipAngel_ReleaseChordModifiersForSend()
    if (priorHwnd)
        ClipAngel_EnsureWindowActive(priorHwnd)
    SendInput "^!b"
}

; Send top list item via Clip Angel native keys. Minimizes Clip Angel and restores prior focus after paste.
ClipAngel_SendTopListItem(priorHwnd := 0) {
    if !ClipAngel_TryAcquireAutomationLock()
        return false
    priorHwnd := ClipAngel_ResolvePriorHwnd(priorHwnd)
    ok := false
    try {
        ClipAngel_SendNativeTopItemKeys(priorHwnd)
        ok := true
    } catch Error as e {
        ShowCenteredOverlay_Utils("❌ Clip Angel paste failed: " . e.Message, 2500, BANNER_ACCENT_ERROR)
        ok := false
    } finally {
        ClipAngel_CloseAndRestoreFocus(priorHwnd)
        ClipAngel_ReleaseAutomationLock()
    }
    return ok
}

; Paste the currently selected Clip Angel row (no ^{Home}). Used by #!+L after editing a clip —
; forcing Row 0 would re-target raw dictation still on the OS clipboard / newest ingest.
; alreadyLocked=true when caller holds g_ClipAngelAutomationBusy (e.g. across the window picker).
ClipAngel_SendSelectedListItem(priorHwnd := 0, alreadyLocked := false) {
    if (!alreadyLocked && !ClipAngel_TryAcquireAutomationLock())
        return false
    priorHwnd := ClipAngel_ResolvePriorHwnd(priorHwnd)
    ok := false
    savedClip := ClipboardAll()
    clipSynced := false
    try {
        ; #region agent log
        ClipAngel_DebugLogSelected("A", "entry", priorHwnd)
        ; #endregion
        ClipAngel_WaitChordModifiersReleased()
        ClipAngel_ReleaseChordModifiersForSend()
        hwnd := ClipAngel_MainHwnd()
        if !hwnd {
            ShowCenteredOverlay_Utils("❌ Clip Angel não está em execução.", 2000, BANNER_ACCENT_ERROR)
            return false
        }
        ; Sync OS clipboard to CA preview BEFORE show — otherwise Clip Angel's monitor
        ; ingests leftover dictation as a new Row 0 and steals the paste target.
        preview := ClipAngel_UiaGetPreviewText(hwnd)
        if (preview != "") {
            try {
                A_Clipboard := preview
                ClipWait 0.4
                clipSynced := true
            } catch {
                clipSynced := false
            }
        }
        ; #region agent log
        ClipAngel_DebugLogSelected("A", "after_clip_sync", priorHwnd)
        ; #endregion
        targetMon := 0
        if (priorHwnd) {
            try targetMon := GetAhkMonitorIndexFromHwnd(priorHwnd)
            catch
                targetMon := 0
        }
        ; #region agent log
        ClipAngel_DebugLogSelected("A", "before_layout", priorHwnd)
        ; #endregion
        ; activate:=false — avoid focus churn; ^!b is handled by Clip Angel into the foreground target.
        ClipAngel_EnsureVisibleAndLayout(hwnd, targetMon, false)
        ; #region agent log
        ClipAngel_DebugLogSelected("A", "after_layout", priorHwnd)
        ; #endregion
        dataGrid := ClipAngel_UiaGetDataGrid(hwnd)
        if (dataGrid)
            ClipAngel_UiaEnsureGridListFocus(dataGrid, hwnd)
        ; #region agent log
        ClipAngel_DebugLogSelected("B", "after_grid_focus", priorHwnd)
        ; #endregion
        ClipAngel_ReleaseChordModifiersForSend()
        if (priorHwnd)
            ClipAngel_EnsureWindowActive(priorHwnd)
        ; #region agent log
        ClipAngel_DebugLogSelected("C", "before_paste_keys", priorHwnd)
        ; #endregion
        SendInput "^!b"
        Sleep CLIPANGEL_INCREMENTAL_PASTE_SETTLE_MS
        ; #region agent log
        ClipAngel_DebugLogSelected("D", "after_paste_keys", priorHwnd)
        ; #endregion
        ok := true
    } catch Error as e {
        ShowCenteredOverlay_Utils("❌ Clip Angel paste failed: " . e.Message, 2500, BANNER_ACCENT_ERROR)
        ok := false
    } finally {
        if (clipSynced) {
            try A_Clipboard := savedClip
            catch {
            }
        }
        ClipAngel_CloseAndRestoreFocus(priorHwnd)
        ; #region agent log
        ClipAngel_DebugLogSelected("E", "after_close", priorHwnd)
        ; #endregion
        if (!alreadyLocked)
            ClipAngel_ReleaseAutomationLock()
    }
    return ok
}

; Preview pane text for the clip currently shown in Clip Angel (works while minimized).
ClipAngel_UiaGetPreviewText(hwnd) {
    if !hwnd
        return ""
    try {
        root := UIA.ElementFromHandle(hwnd)
        if !root
            return ""
        el := ClipAngel_UiaFindFirst(root, { AutomationId: "richTextBox" })
        if !el
            el := ClipAngel_UiaFindFirst(root, { AutomationId: "urlTextBox" })
        if !el
            return ""
        try {
            v := el.Value
            if (v != "")
                return String(v)
        } catch {
        }
        try return String(el.Name)
        catch {
        }
    } catch {
    }
    return ""
}

; #region agent log
; Debug: fingerprint OS clipboard + Clip Angel Row 0 / focus (no plaintext content).
ClipAngel_DebugFp(s) {
    if (s = "")
        return "0:0"
    n := StrLen(s)
    h := 2166136261
    i := 1
    while (i <= n && i <= 64) {
        h := (h * 16777619) ^ Ord(SubStr(s, i, 1))
        i += 1
    }
    return n ":" Format("{:08x}", h & 0xFFFFFFFF)
}
ClipAngel_DebugLogSelected(hypothesisId, step, priorHwnd := 0) {
    try {
        path := A_ScriptDir "\debug-be5edc.log"
        clipFp := ""
        try clipFp := ClipAngel_DebugFp(A_Clipboard)
        catch
            clipFp := "err"
        hwnd := ClipAngel_MainHwnd()
        shown := hwnd ? !!ClipAngel_IsWindowShown(hwnd) : false
        ; Write step marker first so UIA failures still leave a trail.
        FileAppend('{"sessionId":"be5edc","runId":"post-fix","hypothesisId":"' hypothesisId '","location":"clip_angel_favorite.ahk:SendSelected","message":"' step '_begin","timestamp":' A_TickCount ',"data":{"clipFp":"' clipFp '","shown":' (
            shown ? "true" : "false") ',"prior":' (priorHwnd + 0) ',"hwnd":' (hwnd + 0) '}}`n', path, "UTF-8")
        row0Fp := ""
        focusFp := ""
        focusType := 0
        try {
            if (hwnd) {
                root := UIA.ElementFromHandle(hwnd)
                dataGrid := ClipAngel_UiaGetDataGrid(hwnd, root)
                if (dataGrid) {
                    row0 := ClipAngel_UiaResolveRow0(dataGrid)
                    rn := ""
                    try rn := row0.Name
                    catch {
                    }
                    row0Fp := ClipAngel_DebugFp(rn)
                    title := ""
                    if (row0) {
                        tc := ClipAngel_UiaFindFirst(row0, { Type: 50006, Name: "Title " rn })
                        if (!tc && rn = "Row 0")
                            tc := ClipAngel_UiaFindFirst(row0, { Type: 50006, Name: "Title Row 0" })
                        try title := tc ? tc.Value : ""
                        catch {
                        }
                        if (title != "")
                            row0Fp := ClipAngel_DebugFp(String(title))
                    }
                }
            }
            fe := UIA.GetFocusedElement()
            if (fe) {
                try focusType := fe.Type
                catch {
                }
                fn := ""
                try fn := fe.Name
                catch {
                }
                focusFp := ClipAngel_DebugFp(fn)
            }
        } catch {
        }
        line := '{"sessionId":"be5edc","runId":"post-fix","hypothesisId":"' hypothesisId '","location":"clip_angel_favorite.ahk:SendSelected","message":"' step '","timestamp":' A_TickCount ',"data":{"clipFp":"' clipFp '","row0Fp":"' row0Fp '","focusFp":"' focusFp '","focusType":' focusType ',"shown":' (
            shown ? "true" : "false") ',"prior":' (priorHwnd + 0) ',"hwnd":' (hwnd + 0) '}}`n'
        FileAppend(line, path, "UTF-8")
    } catch {
    }
}
; #endregion

; Paste N top-list items in order. Opens once, incremental paste between items, minimizes + restores prior focus at end.
ClipAngel_SendTopListItemSequential(count, priorHwnd := 0) {
    if (!IsInteger(count))
        return false
    n := Integer(count)
    if (n < 1)
        return false
    if !ClipAngel_TryAcquireAutomationLock()
        return false
    priorHwnd := ClipAngel_ResolvePriorHwnd(priorHwnd)
    ok := false
    try {
        ClipAngel_WaitChordModifiersReleased()
        ClipAngel_ReleaseChordModifiersForSend()
        loop n {
            if (A_Index = 1) {
                ClipAngel_SendNativeTopItemKeys(priorHwnd)
            } else {
                Sleep(CLIPANGEL_SEQUENTIAL_PASTE_GAP_MS)
                ClipAngel_ReleaseChordModifiersForSend()
                SendInput "^!b"
            }
            Sleep(CLIPANGEL_INCREMENTAL_PASTE_SETTLE_MS)
        }
        ok := true
    } catch Error as e {
        ShowCenteredOverlay_Utils("❌ Clip Angel sequential paste failed: " . e.Message, 2500, BANNER_ACCENT_ERROR)
        ok := false
    } finally {
        ClipAngel_CloseAndRestoreFocus(priorHwnd)
        ClipAngel_ReleaseAutomationLock()
    }
    return ok
}

; target: "first" = top grid row (Row 0 / newest), "last" = last row returned by UIA FindAll
; (virtualized lists may only expose visible rows - use "first" for reliable top-clip behavior).
ClipAngel_UiaGetMarkFilterValue(hwnd) {
    try {
        root := UIA.ElementFromHandle(hwnd)
        mf := ClipAngel_UiaFindFirst(root, { AutomationId: "MarkFilter", Type: 50003 })
        return mf ? mf.Value : ""
    } catch {
        return ""
    }
}

; MarkFilter combo: leave "favorite" (or any non-all) for "all marks" via dropdown (fast fallback).
ClipAngel_UiaSetMarkFilterAllMarks(hwnd) {
    if !hwnd
        return false
    try {
        root := UIA.ElementFromHandle(hwnd)
        mf := ClipAngel_UiaFindFirst(root, { AutomationId: "MarkFilter", Type: 50003 })
        if !mf
            return false
        if (StrLower(Trim(mf.Value)) = "all marks")
            return true
        mf.SetFocus()
        mf.Click()
        Sleep 30
        Send "{Home}"
        Sleep 30
        Send "{Tab}"
        Sleep 30
        return (StrLower(Trim(ClipAngel_UiaGetMarkFilterValue(hwnd))) = "all marks")
    } catch {
        return false
    }
}

; One native Shift+P at SendLevel 0 (OnSubmitO runs under #InputLevel 10; default Send is ignored by ClipAngel).
ClipAngel_NativeSendShiftP(hwnd, timeoutMs := 220) {
    if (StrLower(Trim(ClipAngel_UiaGetMarkFilterValue(hwnd))) = "all marks")
        return true
    if !WinActive("ahk_id " hwnd)
        ClipAngel_EnsureWindowActive(hwnd, 200)
    ClipAngel_ReleaseChordModifiersForSend()
    try {
        root := UIA.ElementFromHandle(hwnd)
        dataGrid := ClipAngel_UiaGetDataGrid(hwnd, root)
        if dataGrid
            try dataGrid.SetFocus()
    } catch {
    }
    priorSendLevel := A_SendLevel
    SendLevel 0
    SendInput "+p"
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        if (StrLower(Trim(ClipAngel_UiaGetMarkFilterValue(hwnd))) = "all marks") {
            SendLevel priorSendLevel
            return true
        }
        Sleep 15
    }
    SendLevel priorSendLevel
    return (StrLower(Trim(ClipAngel_UiaGetMarkFilterValue(hwnd))) = "all marks")
}

; ^Home first; full UIA row-0 only if still not selected.
ClipAngel_FastEnsureRow0(hwnd) {
    if ClipAngel_UiaIsRow0Selected(hwnd)
        return true
    try {
        root := UIA.ElementFromHandle(hwnd)
        dataGrid := ClipAngel_UiaGetDataGrid(hwnd, root)
        if dataGrid
            try dataGrid.SetFocus()
    } catch {
    }
    priorSendLevel := A_SendLevel
    SendLevel 0
    ClipAngel_ReleaseChordModifiersForSend()
    SendInput "^{Home}"
    deadline := A_TickCount + 180
    while (A_TickCount < deadline) {
        if ClipAngel_UiaIsRow0Selected(hwnd) {
            SendLevel priorSendLevel
            return true
        }
        Sleep 15
    }
    SendLevel priorSendLevel
    if ClipAngel_UiaIsRow0Selected(hwnd)
        return true
    return ClipAngel_UiaEnsureRow0Selected(hwnd, true)
}

; Native Shift+P to leave favorites filter; fast UIA MarkFilter fallback when synthetic Send is ignored.
ClipAngel_LeaveFavoritesFilter(hwnd) {
    result := Map("method", "none", "ok", false, "before", "", "after", "", "row0Selected", false)
    if !hwnd
        return result
    result["before"] := ClipAngel_UiaGetMarkFilterValue(hwnd)
    if (StrLower(Trim(result["before"])) = "all marks") {
        result["ok"] := true
        result["method"] := "already_all_marks"
        result["after"] := result["before"]
        result["row0Selected"] := ClipAngel_FastEnsureRow0(hwnd)
        return result
    }
    if ClipAngel_NativeSendShiftP(hwnd) {
        result["ok"] := true
        result["method"] := "native_shift_p"
        result["after"] := ClipAngel_UiaGetMarkFilterValue(hwnd)
        result["row0Selected"] := ClipAngel_UiaIsRow0Selected(hwnd) || ClipAngel_FastEnsureRow0(hwnd)
        return result
    }
    if ClipAngel_UiaSetMarkFilterAllMarks(hwnd) {
        result["ok"] := true
        result["method"] := "uia_combo"
        result["after"] := ClipAngel_UiaGetMarkFilterValue(hwnd)
        result["row0Selected"] := ClipAngel_FastEnsureRow0(hwnd)
        return result
    }
    result["after"] := ClipAngel_UiaGetMarkFilterValue(hwnd)
    return result
}

ClipAngel_UiaIsRow0Selected(hwnd) {
    try {
        root := UIA.ElementFromHandle(hwnd)
        dataGrid := ClipAngel_UiaGetDataGrid(hwnd, root)
        if !dataGrid
            return false
        row0 := ClipAngel_UiaResolveRow0(dataGrid)
        if !row0
            return false
        return ClipAngel_UiaRow0IsSelected(row0, dataGrid, ClipAngel_UiaGridHasSelectionPattern(dataGrid))
    } catch {
        return false
    }
}

MarkLastClipAsFavorite(target := "first", waitForIngest := false) {
    if waitForIngest
        Sleep(CLIPANGEL_PRE_FAVORITE_INGEST_DELAY_MS)
    if !ClipAngel_TryAcquireAutomationLock()
        return
    priorHwnd := ClipAngel_ResolvePriorHwnd(0)
    try {
        if (target = "last") {
            MarkLastClipAsFavorite_UiaLastRow()
            return
        }
        ClipAngel_ActivateNativeFirstClip()
        if !ClipAngel_WaitForListReady(CLIPANGEL_FAVORITE_OPEN_READY_MS, true) {
            ShowCenteredOverlay_Utils("❌ Clip Angel did not open.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        Sleep(CLIPANGEL_FAVORITE_UI_SETTLE_MS)
        ClipAngel_WaitChordModifiersReleased()
        ClipAngel_ReleaseChordModifiersForSend()
        SendInput "!q"
        ScriptSoundPlay(A_ScriptDir "\assets\sounds\favorite-set.wav")
        ShowCenteredOverlay_Utils("✅ Sent Alt+Q - marked focused clip as favorite.", 1500, BANNER_ACCENT_SUCCESS)
    } catch Error as e {
        ShowCenteredOverlay_Utils("❌ Mark favorite failed: " . e.Message, 2500, BANNER_ACCENT_ERROR)
    } finally {
        ClipAngel_CloseAndRestoreFocus(priorHwnd)
        ClipAngel_ReleaseAutomationLock()
    }
}

; Legacy UIA path for target="last" only (no callers today; API preserved).
MarkLastClipAsFavorite_UiaLastRow() {
    ActivateClipAngelWithFocusCorrection()
    hwnd := ClipAngel_MainHwnd()
    if !hwnd {
        ShowCenteredOverlay_Utils("❌ Clip Angel did not open.", 2000, BANNER_ACCENT_ERROR)
        return
    }
    try WinActivate("ahk_id " hwnd)
    catch {
        ShowCenteredOverlay_Utils("❌ Clip Angel window not found.", 2000, BANNER_ACCENT_ERROR)
        return
    }
    if !WinWaitActive("ahk_id " hwnd, , 2) {
        ShowCenteredOverlay_Utils("❌ Clip Angel did not become active.", 2000, BANNER_ACCENT_ERROR)
        return
    }
    el := UIA.ElementFromHandle(hwnd)
    if !el {
        ShowCenteredOverlay_Utils("❌ Clip Angel UI not available.", 2000, BANNER_ACCENT_ERROR)
        return
    }
    dataGrid := ClipAngel_UiaFindFirst(el, { Type: 50036, AutomationId: "dataGridView" })
    if !dataGrid {
        ShowCenteredOverlay_Utils("❌ Clip list not found (Window tab may still have focus).", 2500,
            BANNER_ACCENT_ERROR)
        return
    }
    rows := 0
    try rows := dataGrid.FindAll({ Type: 50025 })
    catch {
        rows := 0
    }
    if !rows || rows.Length < 1 {
        ShowCenteredOverlay_Utils("❌ No clips in list.", 2000, BANNER_ACCENT_ERROR)
        return
    }
    rowTarget := rows[rows.Length]
    hasSel := rowTarget.GetPropertyValue(UIA.Property.IsSelectionItemPatternAvailable)
    try {
        if hasSel
            rowTarget.SelectionItemPattern.Select()
        else
            rowTarget.SetFocus()
    } catch {
        try rowTarget.SetFocus()
    }
    try {
        if rowTarget.GetPropertyValue(UIA.Property.IsScrollItemPatternAvailable)
            rowTarget.ScrollItemPattern.ScrollIntoView()
    } catch {
    }
    favCell := ClipAngel_FindFavoriteCell(rowTarget)
    if favCell && ClipAngel_FavoriteCellIsOn(favCell) {
        ShowCenteredOverlay_Utils("✅ Selected clip is already a favorite.", 1500, BANNER_ACCENT_SUCCESS)
        return
    }
    if !WinActive("ahk_id " hwnd) {
        try WinActivate("ahk_id " hwnd)
        if !WinWaitActive("ahk_id " hwnd, , 2) {
            ShowCenteredOverlay_Utils("❌ Clip Angel lost focus before Alt+Q.", 2000, BANNER_ACCENT_ERROR)
            return
        }
    }
    Sleep(CLIPANGEL_FAVORITE_UI_SETTLE_MS)
    ClipAngel_WaitChordModifiersReleased()
    ClipAngel_ReleaseChordModifiersForSend()
    SendInput "!q"
    ScriptSoundPlay(A_ScriptDir "\assets\sounds\favorite-set.wav")
    ShowCenteredOverlay_Utils("✅ Sent Alt+Q - marked focused clip as favorite.", 1500, BANNER_ACCENT_SUCCESS)
}
