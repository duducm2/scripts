; =============================================================================
; Utils module: clip_angel_constant_paste.ahk
; Constant Pasting: toggle loop pastes from current Clip Angel selection
; (All or Favorites; no Row-0 jump), Enter only for text, 1.5s interruptible gap.
; Directions: "down" (Shift+P) / "up" (Shift+B). Before start, Char ListView picks
; interstitial delimiter (Enter / Space / Shift+Enter / None); no timeout.
; After paste, Clip Angel moves the used clip to Row 0 — bottom-up next target is
; former N-1 at new index N.
; Efficiency: one UIA refresh per phase; SelectionPattern-first; bounded polls.
; Loaded via #include into Utils.ahk after clip_angel_favorite / activate.
; =============================================================================

CONSTANT_PASTE_GAP_MS := 1500
CONSTANT_PASTE_POLL_MS := 50
CONSTANT_PASTE_CLIPBOARD_WAIT_MS := 200
CONSTANT_PASTE_SELECT_WAIT_MS := 150
CONSTANT_PASTE_SELECT_POLL_MS := 25

global g_ClipAngelConstantPasteActive := false
global g_ClipAngelConstantPasteStopRequested := false
global g_ClipAngelConstantPasteDirection := "down"
global g_ClipAngelConstantPasteStopHint := "Shift+P"
global g_ClipAngelConstantPasteDelimiter := ""
global g_ClipAngelConstantPasteDelimiterPromptActive := false
global g_ClipAngelConstantPasteDelimiterResult := false
global g_ClipAngelConstantPasteDelimiterGui := false
global g_ClipAngelConstantPasteDelimiterLv := false
global g_ClipAngelConstantPasteDelimiterEscPollPrev := false
global g_ClipAngelConstantPasteDelimiterHotkeys := []

; Char-first delimiter options for interstitial separator between clips.
ClipAngel_ConstantPaste_DelimiterOptions() {
    return [{ char: "E", label: "Enter", send: "{Enter}" }, { char: "S", label: "Space", send: "{Space}" }, { char: "H",
        label: "Shift+Enter", send: "+{Enter}" }, { char: "N", label: "None", send: "" }
    ]
}

ClipAngel_ConstantPaste_IsActive() {
    global g_ClipAngelConstantPasteActive
    return !!g_ClipAngelConstantPasteActive
}

ClipAngel_ConstantPaste_IsDelimiterPromptActive() {
    global g_ClipAngelConstantPasteDelimiterPromptActive
    return !!g_ClipAngelConstantPasteDelimiterPromptActive
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

; Single bounded poll (early exit). No ClipWait-then-poll double wait.
ClipAngel_ConstantPaste_WaitClipboardSettle(timeoutMs := 0) {
    if (!timeoutMs)
        timeoutMs := CONSTANT_PASTE_CLIPBOARD_WAIT_MS
    deadline := A_TickCount + timeoutMs
    loop {
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
        if (A_TickCount >= deadline)
            return
        Sleep CONSTANT_PASTE_SELECT_POLL_MS
    }
}

; SelectionPattern first; focused element name; no FindAll row walk.
ClipAngel_ConstantPaste_GetSelectedRowName(hwnd, root := 0, dataGrid := 0) {
    if !hwnd
        return ""
    try {
        if !root {
            root := UIA.ElementFromHandle(hwnd)
            if !root
                return ""
        }
        if !dataGrid
            dataGrid := ClipAngel_UiaGetDataGrid(hwnd, root)
        if dataGrid {
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
        }
        try {
            focused := UIA.GetFocusedElement()
            if focused {
                n := focused.Name
                if RegExMatch(n, "i)(?:Row|Linha)\s*(\d+)", &m)
                    return "Row " m[1]
            }
        } catch {
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

; Poll until selected row index matches (or timeout).
ClipAngel_ConstantPaste_WaitSelectedIndex(hwnd, root, dataGrid, targetIdx, timeoutMs := 0) {
    if (!timeoutMs)
        timeoutMs := CONSTANT_PASTE_SELECT_WAIT_MS
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        name := ClipAngel_ConstantPaste_GetSelectedRowName(hwnd, root, dataGrid)
        if (ClipAngel_ConstantPaste_ParseRowIndex(name) = targetIdx)
            return name
        Sleep CONSTANT_PASTE_SELECT_POLL_MS
    }
    name := ClipAngel_ConstantPaste_GetSelectedRowName(hwnd, root, dataGrid)
    if (ClipAngel_ConstantPaste_ParseRowIndex(name) = targetIdx)
        return name
    return ""
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
    return ClipAngel_ConstantPaste_WaitSelectedIndex(hwnd, root, dataGrid, idx)
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
    dataGrid := ClipAngel_UiaGetDataGrid(hwnd, root)
    Send(goUp ? "{Up}" : "{Down}")
    rowAfter := ClipAngel_ConstantPaste_WaitSelectedIndex(hwnd, root, dataGrid, targetIdx)
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

; UIA root + dataGrid only (no focus). Use after move-to-top invalidates the tree.
ClipAngel_ConstantPaste_RefreshUia(hwnd, &root, &dataGrid) {
    root := 0
    dataGrid := 0
    if !hwnd
        return false
    try root := UIA.ElementFromHandle(hwnd)
    catch
        root := 0
    if !root
        return false
    dataGrid := ClipAngel_UiaGetDataGrid(hwnd, root)
    return !!dataGrid
}

ClipAngel_ConstantPaste_PrepareGrid(hwnd, &root, &dataGrid) {
    if !ClipAngel_ConstantPaste_RefreshUia(hwnd, &root, &dataGrid)
        return false
    ClipAngel_UiaEnsureGridListFocus(dataGrid, hwnd, root)
    return true
}

; Before Enter: skip full UIA rebuild when CA is foreground and grid already focused.
ClipAngel_ConstantPaste_EnsureGridReadyForPaste(hwnd, &root, &dataGrid) {
    if !hwnd
        return false
    if WinActive("ahk_id " hwnd) && dataGrid {
        try {
            if dataGrid.HasKeyboardFocus
                return true
        } catch {
        }
        ClipAngel_UiaEnsureGridListFocus(dataGrid, hwnd, root)
        try {
            if dataGrid.HasKeyboardFocus
                return true
        } catch {
        }
    }
    return ClipAngel_ConstantPaste_PrepareGrid(hwnd, &root, &dataGrid)
}

; ---------------------------------------------------------------------------
; Delimiter picker (Utility Shortcuts Char-first ListView; no timeout)
; ---------------------------------------------------------------------------

ClipAngel_ConstantPaste_DelimiterGuiHwnd() {
    global g_ClipAngelConstantPasteDelimiterGui
    if (!IsObject(g_ClipAngelConstantPasteDelimiterGui))
        return 0
    try
        return g_ClipAngelConstantPasteDelimiterGui.Hwnd
    catch
        return 0
}

ClipAngel_ConstantPaste_DelimiterUnbindHotkeys() {
    global g_ClipAngelConstantPasteDelimiterHotkeys
    hwnd := ClipAngel_ConstantPaste_DelimiterGuiHwnd()
    if (hwnd) {
        try HotIfWinActive("ahk_id " hwnd)
        catch {
        }
    }
    for handler in g_ClipAngelConstantPasteDelimiterHotkeys {
        try Hotkey(handler.key, "Off")
        catch {
        }
    }
    if (hwnd) {
        try HotIf()
        catch {
        }
    }
    g_ClipAngelConstantPasteDelimiterHotkeys := []
}

ClipAngel_ConstantPaste_DelimiterBindHotkeys() {
    global g_ClipAngelConstantPasteDelimiterHotkeys
    ClipAngel_ConstantPaste_DelimiterUnbindHotkeys()
    hwnd := ClipAngel_ConstantPaste_DelimiterGuiHwnd()
    if (!hwnd)
        return
    try HotIfWinActive("ahk_id " hwnd)
    catch {
        return
    }

    for opt in ClipAngel_ConstantPaste_DelimiterOptions() {
        ch := StrLower(opt.char)
        cb := ClipAngel_ConstantPaste_DelimiterFocusChar.Bind(ch)
        try {
            Hotkey(ch, cb, "On")
            g_ClipAngelConstantPasteDelimiterHotkeys.Push({ key: ch, handler: cb })
        } catch {
        }
    }
    try {
        Hotkey("Enter", ClipAngel_ConstantPaste_DelimiterOnEnter, "On")
        g_ClipAngelConstantPasteDelimiterHotkeys.Push({ key: "Enter", handler: ClipAngel_ConstantPaste_DelimiterOnEnter })
    } catch {
    }
    try {
        Hotkey("Escape", ClipAngel_ConstantPaste_DelimiterCancel, "On")
        g_ClipAngelConstantPasteDelimiterHotkeys.Push({ key: "Escape", handler: ClipAngel_ConstantPaste_DelimiterCancel })
    } catch {
    }

    try HotIf()
    catch {
    }
}

ClipAngel_ConstantPaste_DelimiterFocusChar(char, *) {
    global g_ClipAngelConstantPasteDelimiterPromptActive, g_ClipAngelConstantPasteDelimiterLv
    if (!g_ClipAngelConstantPasteDelimiterPromptActive || !IsObject(g_ClipAngelConstantPasteDelimiterLv))
        return
    ch := StrUpper(SubStr(char, 1, 1))
    row := 0
    loop g_ClipAngelConstantPasteDelimiterLv.GetCount() {
        try {
            if (StrUpper(g_ClipAngelConstantPasteDelimiterLv.GetText(A_Index, 1)) = ch) {
                row := A_Index
                break
            }
        } catch {
        }
    }
    if (row > 0)
        ListView_SelectRowFocused(g_ClipAngelConstantPasteDelimiterLv, row)
}

ClipAngel_ConstantPaste_DelimiterFocusedSend() {
    global g_ClipAngelConstantPasteDelimiterLv
    if (!IsObject(g_ClipAngelConstantPasteDelimiterLv))
        return false
    row := 0
    try row := g_ClipAngelConstantPasteDelimiterLv.GetNext(0, "Focused")
    catch {
        row := 0
    }
    if (row < 1) {
        try row := g_ClipAngelConstantPasteDelimiterLv.GetNext(0, "Selected")
        catch {
            row := 0
        }
    }
    if (row < 1)
        return false
    ch := ""
    try ch := StrUpper(Trim(g_ClipAngelConstantPasteDelimiterLv.GetText(row, 1)))
    catch {
        return false
    }
    for opt in ClipAngel_ConstantPaste_DelimiterOptions() {
        if (StrUpper(opt.char) = ch)
            return opt.send
    }
    return false
}

ClipAngel_ConstantPaste_DelimiterOnEnter(*) {
    global g_ClipAngelConstantPasteDelimiterPromptActive, g_ClipAngelConstantPasteDelimiterResult
    if (!g_ClipAngelConstantPasteDelimiterPromptActive)
        return
    sendStr := ClipAngel_ConstantPaste_DelimiterFocusedSend()
    if (sendStr = false)
        return
    g_ClipAngelConstantPasteDelimiterResult := sendStr
    ClipAngel_ConstantPaste_DelimiterClose()
}

ClipAngel_ConstantPaste_DelimiterOnListActivate(*) {
    ClipAngel_ConstantPaste_DelimiterOnEnter()
}

ClipAngel_ConstantPaste_DelimiterCancel(*) {
    global g_ClipAngelConstantPasteDelimiterPromptActive, g_ClipAngelConstantPasteDelimiterResult
    if (!g_ClipAngelConstantPasteDelimiterPromptActive)
        return
    g_ClipAngelConstantPasteDelimiterResult := false
    ClipAngel_ConstantPaste_DelimiterClose()
}

ClipAngel_ConstantPaste_DelimiterBindRobustEscape() {
    global g_ClipAngelConstantPasteDelimiterGui, g_OnEscapePressed, g_ClipAngelConstantPasteDelimiterEscPollPrev
    SetTimer(ClipAngel_ConstantPaste_DelimiterEscapePoll, 0)
    if (!ClipAngel_ConstantPaste_DelimiterGuiHwnd())
        return
    try g_ClipAngelConstantPasteDelimiterGui.OnEvent("Escape", ClipAngel_ConstantPaste_DelimiterCancel)
    catch {
    }
    try Hotkey("$*Escape", ClipAngel_ConstantPaste_DelimiterEscapeFromHotkey, "On")
    catch {
    }
    g_OnEscapePressed := ClipAngel_ConstantPaste_DelimiterGlobalEscapeCallback
    try Utils_EnsureGlobalEscapeHotkey()
    catch {
    }
    g_ClipAngelConstantPasteDelimiterEscPollPrev := false
    SetTimer(ClipAngel_ConstantPaste_DelimiterEscapePoll, 50)
}

ClipAngel_ConstantPaste_DelimiterUnbindRobustEscape() {
    global g_OnEscapePressed, g_ClipAngelConstantPasteDelimiterEscPollPrev
    SetTimer(ClipAngel_ConstantPaste_DelimiterEscapePoll, 0)
    g_ClipAngelConstantPasteDelimiterEscPollPrev := false
    try Hotkey("$*Escape", ClipAngel_ConstantPaste_DelimiterEscapeFromHotkey, "Off")
    catch {
    }
    if (g_OnEscapePressed = ClipAngel_ConstantPaste_DelimiterGlobalEscapeCallback)
        g_OnEscapePressed := ""
    try Utils_EnsureGlobalEscapeHotkey()
    catch {
    }
}

ClipAngel_ConstantPaste_DelimiterEscapeFromHotkey(*) {
    ClipAngel_ConstantPaste_DelimiterCancel()
}

ClipAngel_ConstantPaste_DelimiterGlobalEscapeCallback(*) {
    ClipAngel_ConstantPaste_DelimiterCancel()
}

ClipAngel_ConstantPaste_DelimiterEscapePoll() {
    global g_ClipAngelConstantPasteDelimiterPromptActive, g_ClipAngelConstantPasteDelimiterEscPollPrev
    if (!g_ClipAngelConstantPasteDelimiterPromptActive) {
        SetTimer(ClipAngel_ConstantPaste_DelimiterEscapePoll, 0)
        return
    }
    escSync := GetKeyState("Escape", "P")
    escAsync := (DllCall("user32\GetAsyncKeyState", "int", 0x1B) & 0x8000) != 0
    escDown := escSync || escAsync
    if (escDown) {
        if (!g_ClipAngelConstantPasteDelimiterEscPollPrev) {
            g_ClipAngelConstantPasteDelimiterEscPollPrev := true
            ClipAngel_ConstantPaste_DelimiterCancel()
        }
    } else {
        g_ClipAngelConstantPasteDelimiterEscPollPrev := false
    }
}

ClipAngel_ConstantPaste_DelimiterClose() {
    global g_ClipAngelConstantPasteDelimiterGui, g_ClipAngelConstantPasteDelimiterLv
    global g_ClipAngelConstantPasteDelimiterPromptActive
    if (!g_ClipAngelConstantPasteDelimiterPromptActive)
        return
    g_ClipAngelConstantPasteDelimiterPromptActive := false
    ClipAngel_ConstantPaste_DelimiterUnbindHotkeys()
    ClipAngel_ConstantPaste_DelimiterUnbindRobustEscape()
    if (IsObject(g_ClipAngelConstantPasteDelimiterGui)) {
        try g_ClipAngelConstantPasteDelimiterGui.Destroy()
        catch {
        }
    }
    g_ClipAngelConstantPasteDelimiterGui := false
    g_ClipAngelConstantPasteDelimiterLv := false
}

; Blocking: returns send string (may be "") on confirm, or false on cancel. No timeout.
ClipAngel_ConstantPaste_PromptDelimiter() {
    global g_ClipAngelConstantPasteDelimiterPromptActive, g_ClipAngelConstantPasteDelimiterResult
    global g_ClipAngelConstantPasteDelimiterGui, g_ClipAngelConstantPasteDelimiterLv

    if (g_ClipAngelConstantPasteDelimiterPromptActive)
        return false

    ClipAngel_WaitChordModifiersReleased()
    ClipAngel_ReleaseChordModifiersForSend()

    g_ClipAngelConstantPasteDelimiterResult := false
    g_ClipAngelConstantPasteDelimiterPromptActive := true

    g_ClipAngelConstantPasteDelimiterGui := Gui("+AlwaysOnTop +ToolWindow", "Constant Pasting — delimiter")
    g_ClipAngelConstantPasteDelimiterGui.SetFont("s10", "Segoe UI")
    g_ClipAngelConstantPasteDelimiterGui.Add("Text", "w420",
        "Char = jump   Enter/double-click = confirm   Esc = cancel")
    g_ClipAngelConstantPasteDelimiterLv := g_ClipAngelConstantPasteDelimiterGui.Add("ListView",
        "w420 h140 -Multi", ["Char", "Delimiter"])
    g_ClipAngelConstantPasteDelimiterLv.OnEvent("DoubleClick", ClipAngel_ConstantPaste_DelimiterOnListActivate)
    g_ClipAngelConstantPasteDelimiterGui.Add("Button", "w100", "Close").OnEvent("Click",
        ClipAngel_ConstantPaste_DelimiterCancel)
    g_ClipAngelConstantPasteDelimiterGui.OnEvent("Close", ClipAngel_ConstantPaste_DelimiterCancel)
    g_ClipAngelConstantPasteDelimiterGui.OnEvent("Escape", ClipAngel_ConstantPaste_DelimiterCancel)

    for opt in ClipAngel_ConstantPaste_DelimiterOptions()
        g_ClipAngelConstantPasteDelimiterLv.Add("", opt.char, opt.label)
    try g_ClipAngelConstantPasteDelimiterLv.ModifyCol(1, 50)
    try g_ClipAngelConstantPasteDelimiterLv.ModifyCol(2, 340)
    if (g_ClipAngelConstantPasteDelimiterLv.GetCount() > 0)
        ListView_SelectRowFocused(g_ClipAngelConstantPasteDelimiterLv, 1)

    GetActiveMonitorWorkArea_StandardBar(&ml, &mt, &mr, &mb)
    g_ClipAngelConstantPasteDelimiterGui.Show("AutoSize Hide")
    g_ClipAngelConstantPasteDelimiterGui.GetPos(, , &gw, &gh)
    cx := ml + ((mr - ml) - gw) // 2
    cy := mt + ((mb - mt) - gh) // 2
    if (cx < ml)
        cx := ml
    if (cy < mt)
        cy := mt
    g_ClipAngelConstantPasteDelimiterGui.Show("x" . cx . " y" . cy)
    try WinActivate(g_ClipAngelConstantPasteDelimiterGui.Hwnd)
    try g_ClipAngelConstantPasteDelimiterLv.Focus()
    catch {
    }

    ClipAngel_ConstantPaste_DelimiterBindHotkeys()
    ClipAngel_ConstantPaste_DelimiterBindRobustEscape()

    while (g_ClipAngelConstantPasteDelimiterPromptActive)
        Sleep 50

    return g_ClipAngelConstantPasteDelimiterResult
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
    if (ClipAngel_ConstantPaste_IsDelimiterPromptActive())
        return
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
    global g_ClipAngelConstantPasteDelimiter
    if (g_ClipAngelConstantPasteActive || ClipAngel_ConstantPaste_IsDelimiterPromptActive())
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

    delim := ClipAngel_ConstantPaste_PromptDelimiter()
    if (delim = false)
        return false
    g_ClipAngelConstantPasteDelimiter := delim

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
            ; One full prepare at iteration start.
            if !ClipAngel_ConstantPaste_PrepareGrid(hwnd, &root, &dataGrid) {
                stopReason := "clip list not ready"
                break
            }

            ; Interstitial separator before each clip after the first.
            if (pastedCount > 0 && g_ClipAngelConstantPasteDelimiter != "") {
                ClipAngel_RestorePriorFocus(priorHwnd)
                ClipAngel_ReleaseChordModifiersForSend()
                Send g_ClipAngelConstantPasteDelimiter
            }

            rowBefore := ClipAngel_ConstantPaste_GetSelectedRowName(hwnd, root, dataGrid)
            if (rowBefore = "") {
                stopReason := "no clip selected"
                break
            }

            ClipAngel_WaitChordModifiersReleased()
            ClipAngel_ReleaseChordModifiersForSend()
            ; Prime Clip Angel's "previous window", then Send Enter (same as Alt+1).
            ClipAngel_RestorePriorFocus(priorHwnd)
            if !ClipAngel_EnsureWindowActive(hwnd, 400) {
                stopReason := "could not focus Clip Angel for paste"
                break
            }
            if !ClipAngel_ConstantPaste_EnsureGridReadyForPaste(hwnd, &root, &dataGrid) {
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
            try WinShow("ahk_id " hwnd)
            catch {
            }
            if !ClipAngel_EnsureWindowActive(hwnd, 400) {
                stopReason := "could not focus Clip Angel"
                break
            }
            ; Move-to-top invalidates UIA — one refresh, then advance (no third full PrepareGrid).
            if !ClipAngel_ConstantPaste_RefreshUia(hwnd, &root, &dataGrid) {
                stopReason := "clip list not ready"
                break
            }
            ClipAngel_UiaEnsureGridListFocus(dataGrid, hwnd, root)

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
        g_ClipAngelConstantPasteDelimiter := ""
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
