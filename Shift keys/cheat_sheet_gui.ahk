; =============================================================================
; Shift keys module: cheat_sheet_gui.ahk
; Cheat sheet GUI, search, hold detection
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

; ========== Helper to decide which sheet applies ===========================
GetCheatSheetText() {
    global cheatSheets

    exe := WinGetProcessName("A") ; active process name (e.g. chrome.exe)
    title := WinGetTitle("A")
    hwnd := WinExist("A")

    ; Microsoft Store "new Outlook" runs as olk.exe; use same cheat sheets as OUTLOOK.EXE.
    if (StrLower(exe) = "olk.exe")
        exe := "OUTLOOK.EXE"

    ; (removed temporary tooltip debugging)

    ; Prefer Outlook Reminders over generic File Dialog detection
    if (exe = "OUTLOOK.EXE") {
        if RegExMatch(title, "i)Reminder")
            return cheatSheets.Has("OutlookReminder") ? cheatSheets["OutlookReminder"] : ""
    }

    ; Check for file dialog first (works in any app)
    if WinGetClass("ahk_id " hwnd) = "#32770" {
        txt := WinGetText("ahk_id " hwnd)
        if InStr(txt, "Namespace Tree Control") || InStr(txt, "Controle da Ãrvore de Namespace") {
            return cheatSheets["FileDialog"]
        }
    }

    ; Check for Settings window (both English and Portuguese)
    if (title = "Settings" || title = "ConfiguraÃ§Ãµes") {
        return cheatSheets.Has("Settings") ? cheatSheets["Settings"] : ""
    }

    ; Check for Command Palette window
    if InStr(title, "Command Palette", false) {
        return cheatSheets.Has("Command Palette") ? cheatSheets["Command Palette"] : ""
    }

    ; Check for Power BI (by process name or window title)
    if (exe = "PBIDesktop.exe" || InStr(title, "powerbi", false)) {
        return cheatSheets.Has("Power BI") ? cheatSheets["Power BI"] : ""
    }

    ; Special handling for Chrome-based apps that share chrome.exe
    if (exe = "chrome.exe") {
        chromeShortcuts := cheatSheets.Has("chrome.exe") ? cheatSheets["chrome.exe"] : ""
        appShortcuts := ""

        ; Normalize Chrome window title by removing the trailing " - Google Chrome"
        chromeTitle := RegExReplace(title, "i) - Google Chrome$", "")

        siteKey := ""
        if (hwnd) {
            try {
                if (CopilotWeb_IsCopilotHwnd(hwnd, "fast") || CopilotWeb_IsCopilotHwnd(hwnd, "full")
                || CopilotWeb_TryUiaFingerprint(hwnd))
                    siteKey := "Copilot Web"
            } catch {
            }
        }
        if (siteKey = "")
            siteKey := PickChromeAppSheetKey(chromeTitle)
        if (siteKey != "" && cheatSheets.Has(siteKey))
            appShortcuts := cheatSheets[siteKey]

        ; Combine Chrome general + app-specific shortcuts
        if (appShortcuts != "" && chromeShortcuts != "")
            return chromeShortcuts "`r`n`r`n" appShortcuts
        if (appShortcuts != "")
            return appShortcuts
        if (chromeShortcuts != "")
            return chromeShortcuts
        return ""
    }

    ; UIA Tree Inspector - check both process and window title
    if (exe = "AutoHotkey64.exe" && InStr(title, "UIATreeInspector"))
        return cheatSheets["UIATreeInspector"]

    ; Microsoft Teams â€" differentiate meeting vs chat via helper predicates
    if IsTeamsMeetingActive()
        return cheatSheets.Has("TeamsMeeting") ? cheatSheets["TeamsMeeting"] : ""
    if IsTeamsChatActive()
        return cheatSheets.Has("TeamsChat") ? cheatSheets["TeamsChat"] : ""
    if IsFileDialogActive()
        return cheatSheets["FileDialog"]

    ; Special handling for Outlook-based apps
    if (exe = "OUTLOOK.EXE") {
        ; Detect Reminders window â€" e.g. "3 Reminder(s)" or any title containing "Reminder"
        if RegExMatch(title, "i)Reminder") {
            return cheatSheets.Has("OutlookReminder") ? cheatSheets["OutlookReminder"] : cheatSheets["OUTLOOK.EXE"]
        }
        ; Detect Message inspector windows â€" e.g., " - Message (HTML)"
        if RegExMatch(title, "i) - Message \(") {
            return cheatSheets.Has("OutlookMessage") ? cheatSheets["OutlookMessage"] : cheatSheets["OUTLOOK.EXE"]
        }
        ; Detect Appointment, Meeting, or Event inspector windows
        if RegExMatch(title, "i)(Appointment|Meeting|Event)") {
            return cheatSheets.Has("OutlookAppointment") ? cheatSheets["OutlookAppointment"] : cheatSheets[
                "OUTLOOK.EXE"]
        }
        ; Fallback to generic Outlook sheet
        if cheatSheets.Has("OUTLOOK.EXE")
            return cheatSheets["OUTLOOK.EXE"]
    }

    ; Direct match by process name (generic fallback)
    if cheatSheets.Has(exe)
        return cheatSheets[exe]

    ; Try case-insensitive match for process name
    for key, value in cheatSheets {
        if (StrLower(key) = StrLower(exe))
            return value
    }

    ; Nothing found > blank > caller will show fallback message
    return ""
}

; Mirrors the former sequential if-chain: later assignments override earlier ones; Shopee/Google only when still unset.
PickChromeAppSheetKey(chromeTitle) {
    key := ""
    if IsChromePdfViewerActive()
        key := "Chrome PDF Viewer"
    if InStr(chromeTitle, "WhatsApp")
        key := "WhatsApp"
    if InStr(chromeTitle, "Gmail")
        key := "Gmail"
    if InStr(chromeTitle, "chatgpt")
        key := "ChatGPT"
    if InStr(chromeTitle, "Mobills")
        key := "Mobills"
    if InStr(chromeTitle, "Google Keep") || InStr(chromeTitle, "keep.google.com")
        key := "Google Keep"
    if InStr(chromeTitle, "YouTube")
        key := "YouTube"
    if InStr(chromeTitle, "UIATreeInspector")
        key := "UIATreeInspector"
    if InStr(chromeTitle, "Settle Up")
        key := "Settle Up"
    if InStr(chromeTitle, "Miro")
        key := "Miro"
    if InStr(chromeTitle, "Wikipedia", false) || InStr(chromeTitle, "wikipedia.org", false)
        key := "Wikipedia"
    if IsMercadoLivreActive()
        key := "Mercado Livre"
    if (key = "" && IsShopeeActive())
        key := "Shopee"
    if InStr(chromeTitle, "gemini", false)
        key := "Gemini"
    if (key = "" && (InStr(chromeTitle, "M365 Copilot", false) || InStr(chromeTitle, "Chat | M365 Copilot", false)))
        key := "Copilot Web"
    if InStr(chromeTitle, "Google Maps")
        key := "Google Maps"
    if (key = "" && (chromeTitle = "Google" || InStr(chromeTitle, " - Google Search")))
        key := "Google"
    return key
}

; ========== Shared variables for cheat sheet state ========================
global g_helpGui := 0
global g_helpShown := false
global g_globalGui := 0
global g_globalShown := false
global g_helpSearchEdit := 0
global g_helpLv := 0
global g_globalSearchEdit := 0
global g_globalLv := 0
global g_cheatSheetAppFullProcessed := ""
global g_cheatSheetGlobalFullProcessed := ""
global g_cheatSheetAppRows := []
global g_cheatSheetGlobalRows := []
global g_searchAllGui := 0
global g_cheatSheetSuppressFilter := false

; ListView uses system white background; black text for readability.
global CHEAT_SHEET_LV_TEXT := 0x000000    ; black (COLORREF BGR)

; #region agent log
CheatSheet_AgentDebugLog(hypothesisId, location, message, dataJson := "{}", runId := "post-fix") {
    try {
        logPath := A_ScriptDir "\debug-e6792c.log"
        line := Format(
            '{{"sessionId":"e6792c","timestamp":{1},"location":"{2}","message":"{3}","hypothesisId":"{4}","runId":"{5}","data":{6}}}`n',
            A_TickCount, location, message, hypothesisId, runId, dataJson)
        FileAppend line, logPath
    } catch {
    }
}
; #endregion

GetGlobalCheatSheetRawText() {
    global GLOBAL_CHEAT_SHEET_RAW
    provider := GetGlobalAIProviderLabel()
    return StrReplace(GLOBAL_CHEAT_SHEET_RAW, "{AI_PROVIDER}", provider)
}

; Returns Map of context label -> array of matching processed lines. Empty query => empty map.
SearchCheatSheets(query, includeGlobal := true) {
    global cheatSheets
    q := Trim(query)
    results := Map()
    if (q = "")
        return results

    for key, text in cheatSheets {
        if (text = "")
            continue
        proc := ProcessCheatSheetText(NormalizeMojibake(text))
        lines := StrSplit(proc, "`n")
        hits := []
        for line in lines {
            if CheatSheet_LineMatchesQuery(line, q)
                hits.Push(line)
        }
        if (hits.Length)
            results[key] := hits
    }

    if includeGlobal {
        proc := ProcessCheatSheetText(NormalizeMojibake(GetGlobalCheatSheetRawText()))
        lines := StrSplit(proc, "`n")
        hits := []
        for line in lines {
            if CheatSheet_LineMatchesQuery(line, q)
                hits.Push(line)
        }
        if (hits.Length)
            results["(Global shortcuts)"] := hits
    }
    return results
}

CheatSheet_ApplyListViewTheme(lv) {
    global CHEAT_SHEET_LV_TEXT
    if !IsObject(lv)
        return
    ; Light theme only: black text on default white ListView (no SetWindowTheme — can block the hotkey thread).
    textRet := 0
    try textRet := SendMessage(0x1024, 0, CHEAT_SHEET_LV_TEXT, lv) ; LVM_SETTEXTCOLOR
    ; #region agent log
    CheatSheet_AgentDebugLog("C", "cheat_sheet_gui.ahk:ApplyListViewTheme", "text color only",
        Format('{{"textRet":{1}}}', textRet), "post-fix")
    ; #endregion
}

CheatSheet_ConfigureSheetListViewColumns(lv) {
    lv.ModifyCol(1, 200)
    lv.ModifyCol(2, 280)
    lv.ModifyCol(3, "AutoHdr")
}

CheatSheet_ConfigureSheetListView(lv) {
    CheatSheet_ConfigureSheetListViewColumns(lv)
}

; Parse processed sheet into ListView rows (Section | Shortcut | Description).
CheatSheet_ParseSheetRows(processedText) {
    rows := []
    section := ""
    for line in StrSplit(processedText, "`n", "`r") {
        rawLine := line
        stripped := RegExReplace(line, "^(>>>\s*|---\s*)", "")
        stripped := Trim(stripped, "`r`n `t")
        if (stripped = "")
            continue
        if RegExMatch(stripped, "^\s*===\s*(.+?)\s*===\s*$", &hm) {
            section := Trim(hm[1])
            continue
        }
        if RegExMatch(stripped, "^\s*(.+)\s*\([^)]+\)\s*$") && !InStr(stripped, "[") {
            section := stripped
            continue
        }
        shortcut := ""
        description := ""
        if RegExMatch(stripped, "i)^\[ZMK\s([^\]]+)\]\s*(.+?)\s*>\s*(.*)$", &zm) {
            shortcut := Trim(zm[1]) . " · " . Trim(zm[2])
            description := zm[3]
        } else if RegExMatch(stripped, "^\[(.+?)\]\s*>\s*(.*)$", &m) {
            shortcut := Trim(m[1])
            description := m[2]
        } else if RegExMatch(stripped, "\[(.+?)\]", &m) {
            shortcut := Trim(m[1])
            pos := InStr(stripped, "]")
            description := Trim(SubStr(stripped, pos + 1))
        } else {
            description := stripped
        }
        rows.Push({
            section: section,
            shortcut: shortcut,
            description: description,
            rawLine: rawLine
        })
    }
    return rows
}

CheatSheet_FilterSheetRows(rows, query) {
    q := Trim(query)
    if (q = "")
        return rows
    filtered := []
    for row in rows {
        if CheatSheet_LineMatchesQuery(row.rawLine, q)
            filtered.Push(row)
    }
    return filtered
}

CheatSheet_RefreshSheetListView(lv, rows) {
    lvOk := IsObject(lv)
    rowIn := rows.Length
    if !lvOk {
        ; #region agent log
        CheatSheet_AgentDebugLog("A", "cheat_sheet_gui.ahk:RefreshSheetListView", "lv not object", "{}")
        ; #endregion
        return
    }
    lv.Delete()
    for row in rows
        lv.Add("", row.section, row.shortcut, row.description)
    lvCount := 0
    try lvCount := lv.GetCount()
    ; #region agent log
    CheatSheet_AgentDebugLog("D", "cheat_sheet_gui.ahk:RefreshSheetListView", "after refresh",
        Format('{{"lvOk":1,"rowIn":{1},"lvCount":{2},"sampleShortcut":"{3}"}}',
            rowIn, lvCount, rowIn ? StrReplace(rows[1].shortcut, '"', "'") : ""))
    ; #endregion
}

CheatSheet_ShowSheetListView(gui, lv) {
    chromePx := 52
    guiW := 1000
    lvH := 520
    wx := 0, wy := 0, wr := A_ScreenWidth, wb := A_ScreenHeight
    try GetActiveMonitorWorkArea_StandardBar(&wx, &wy, &wr, &wb)
    maxLvH := wb - wy - chromePx - 12
    if (maxLvH >= 200)
        lvH := maxLvH
    lv.Move(, , guiW, lvH)
    guiH := lvH + chromePx
    guiX := wx + ((wr - wx) - guiW) / 2
    guiY := wy + ((wb - wy) - guiH) / 2
    guiX := Max(wx, Min(guiX, wr - guiW))
    guiY := Max(wy, Min(guiY, wb - guiH))
    ; #region agent log
    CheatSheet_AgentDebugLog("E", "cheat_sheet_gui.ahk:ShowSheetListView", "before show",
        Format('{{"guiW":{1},"guiH":{2}}}', guiW, guiH), "post-fix")
    ; #endregion
    gui.Show("NoActivate x" Round(guiX) " y" Round(guiY) " w" guiW " h" guiH)
    ; #region agent log
    CheatSheet_AgentDebugLog("E", "cheat_sheet_gui.ahk:ShowSheetListView", "after show", "{}", "post-fix")
    ; #endregion
}

CheatSheet_OnSheetListCopy(lv, guiEvent) {
    row := 0
    try
        row := guiEvent.EventInfo
    if !row
        row := lv.GetNext(0, "F")
    if !row
        return
    shortcut := lv.GetText(row, 2)
    description := lv.GetText(row, 3)
    if (shortcut != "" && description != "")
        A_Clipboard := shortcut . " > " . description
    else if (description != "")
        A_Clipboard := description
    else
        A_Clipboard := shortcut
}

CheatSheet_InitSheetOverlayGui(gui, &filterEdit, &lv, onFilter, onEscape, onCopy) {
    gui.BackColor := "FFFFFF"
    gui.SetFont("s10 c000000", "Consolas")
    filterEdit := gui.Add("Edit", "xm w1000 Section Limit20 BackgroundFFFFFF c000000", "")
    lv := gui.Add("ListView", "xm w1000 h520 Grid -Multi +ReadOnly c000000 BackgroundFFFFFF",
        ["Section", "Shortcut", "Description"])
    filterEdit.OnEvent("Change", onFilter)
    lv.OnEvent("DoubleClick", onCopy)
    gui.OnEvent("Escape", onEscape)
    CheatSheet_ConfigureSheetListViewColumns(lv)
    CheatSheet_ApplyListViewTheme(lv)
}

CheatSheet_DestroySheetOverlayIfStale(&gui, &filterEdit, &lv, &shown) {
    if !IsObject(gui)
        return
    if IsObject(filterEdit) && IsObject(lv)
        return
    try {
        if shown
            gui.Hide()
        gui.Destroy()
    } catch {
    }
    gui := 0
    filterEdit := 0
    lv := 0
    shown := false
    ; #region agent log
    CheatSheet_AgentDebugLog("A", "cheat_sheet_gui.ahk:DestroySheetOverlayIfStale", "rebuilt stale gui", "{}")
    ; #endregion
}

CheatSheet_EnsureAppSheetGui() {
    global g_helpGui, g_helpSearchEdit, g_helpLv, g_helpShown
    CheatSheet_DestroySheetOverlayIfStale(&g_helpGui, &g_helpSearchEdit, &g_helpLv, &g_helpShown)
    if !IsObject(g_helpGui) {
        g_helpGui := Gui("+AlwaysOnTop -Caption +ToolWindow +Border")
        CheatSheet_InitSheetOverlayGui(g_helpGui, &g_helpSearchEdit, &g_helpLv, CheatSheet_OnAppFilterChanged,
            CheatSheet_OnEscapeApp, (ctrl, guiEvent) => CheatSheet_OnSheetListCopy(ctrl, guiEvent))
    }
}

CheatSheet_EnsureGlobalSheetGui() {
    global g_globalGui, g_globalSearchEdit, g_globalLv, g_globalShown
    CheatSheet_DestroySheetOverlayIfStale(&g_globalGui, &g_globalSearchEdit, &g_globalLv, &g_globalShown)
    if !IsObject(g_globalGui) {
        g_globalGui := Gui("+AlwaysOnTop -Caption +ToolWindow +Border")
        CheatSheet_InitSheetOverlayGui(g_globalGui, &g_globalSearchEdit, &g_globalLv, CheatSheet_OnGlobalFilterChanged,
            CheatSheet_OnEscapeGlobal, (ctrl, guiEvent) => CheatSheet_OnSheetListCopy(ctrl, guiEvent))
    }
}

CheatSheet_OnAppFilterChanged(*) {
    global g_cheatSheetSuppressFilter, g_helpSearchEdit, g_helpLv, g_cheatSheetAppRows, g_helpGui, g_helpShown
    if (g_cheatSheetSuppressFilter || !IsObject(g_helpLv) || !IsObject(g_helpSearchEdit))
        return
    q := Trim(g_helpSearchEdit.Value)
    rows := CheatSheet_FilterSheetRows(g_cheatSheetAppRows, q)
    CheatSheet_RefreshSheetListView(g_helpLv, rows)
    CheatSheet_DeferFocusSearch(g_helpSearchEdit)
}

CheatSheet_OnGlobalFilterChanged(*) {
    global g_cheatSheetSuppressFilter, g_globalSearchEdit, g_globalLv, g_cheatSheetGlobalRows, g_globalGui,
        g_globalShown
    if (g_cheatSheetSuppressFilter || !IsObject(g_globalLv) || !IsObject(g_globalSearchEdit))
        return
    q := Trim(g_globalSearchEdit.Value)
    rows := CheatSheet_FilterSheetRows(g_cheatSheetGlobalRows, q)
    CheatSheet_RefreshSheetListView(g_globalLv, rows)
    CheatSheet_DeferFocusSearch(g_globalSearchEdit)
}

ShowSearchAllCheatSheetsGui() {
    global g_searchAllGui
    static filterEdit := 0, lv := 0

    if !IsObject(g_searchAllGui) {
        g_searchAllGui := Gui("+AlwaysOnTop +Resize +MinSize800x400", "Search cheat sheets")
        g_searchAllGui.SetFont("s10", "Consolas")
        g_searchAllGui.Add("Text", "xm Section",
            "Filter (description text only; space = AND; max 20 chars):")
        filterEdit := g_searchAllGui.Add("Edit", "xs w780 Limit20")
        lv := g_searchAllGui.Add("ListView", "xm w780 h520 Grid", ["Context", "Line"])
        filterEdit.OnEvent("Change", (*) => CheatSheet_RefreshSearchAllList(filterEdit, lv))
        lv.OnEvent("DoubleClick", (ctrl, guiEvent) => CheatSheet_OnSearchAllCopy(ctrl, guiEvent))
        g_searchAllGui.Add("Text", "xm", "Double-click a row to copy the line.")
        g_searchAllGui.OnEvent("Escape", CheatSheet_OnEscapeSearchAll)
    }
    filterEdit.Value := ""
    CheatSheet_RefreshSearchAllList(filterEdit, lv)
    g_searchAllGui.Show("w800 h620")
    CheatSheet_DeferFocusSearch(filterEdit)
}

CheatSheet_RefreshSearchAllList(filterEdit, lv) {
    lv.Delete()
    q := Trim(filterEdit.Value)
    if (q = "") {
        return
    }
    m := SearchCheatSheets(q, true)
    for ctx, lines in m {
        for line in lines {
            lv.Add("", ctx, line)
        }
    }
}

CheatSheet_OnSearchAllCopy(lv, guiEvent) {
    row := 0
    try
        row := guiEvent.EventInfo
    if !row
        row := lv.GetNext(0, "F")
    if !row
        return
    line := lv.GetText(row, 2)
    A_Clipboard := line
}

; ========== GUI creation & showing ========================================
ToggleShortcutHelp() {
    global g_helpGui, g_helpShown, g_helpSearchEdit, g_helpLv
    global g_cheatSheetAppFullProcessed, g_cheatSheetAppRows

    ; Toggle off if currently shown (or gui exists but flag lost after a partial show)
    if (IsObject(g_helpGui) && g_helpShown) {
        g_helpGui.Hide()
        g_helpShown := false
        return
    }

    ; #region agent log
    CheatSheet_AgentDebugLog("A", "cheat_sheet_gui.ahk:ToggleShortcutHelp", "entry state",
        Format('{{"helpGui":{1},"helpLv":{2},"helpShown":0}}',
            IsObject(g_helpGui) ? 1 : 0, IsObject(g_helpLv) ? 1 : 0))
    ; #endregion

    text := NormalizeMojibake(GetCheatSheetText())
    if (text = "") {
        exe := WinGetProcessName("A")
        msg := "No cheat-sheet registered for: " exe
        g_cheatSheetAppFullProcessed := msg
        g_cheatSheetAppRows := [{
            section: "",
            shortcut: "",
            description: msg,
            rawLine: msg
        }]
    } else {
        g_cheatSheetAppFullProcessed := ProcessCheatSheetText(text)
        g_cheatSheetAppRows := CheatSheet_ParseSheetRows(g_cheatSheetAppFullProcessed)
    }

    ; #region agent log
    CheatSheet_AgentDebugLog("B", "cheat_sheet_gui.ahk:ToggleShortcutHelp", "parsed app rows",
        Format('{{"appRowCount":{1},"procLen":{2}}}', g_cheatSheetAppRows.Length, StrLen(g_cheatSheetAppFullProcessed))
    )
    ; #endregion

    try {
        global g_cheatSheetSuppressFilter
        CheatSheet_EnsureAppSheetGui()
        ; #region agent log
        CheatSheet_AgentDebugLog("A", "cheat_sheet_gui.ahk:ToggleShortcutHelp", "ensure done", "{}")
        ; #endregion
        g_cheatSheetSuppressFilter := true
        g_helpSearchEdit.Value := ""
        g_cheatSheetSuppressFilter := false
        CheatSheet_RefreshSheetListView(g_helpLv, g_cheatSheetAppRows)
        CheatSheet_ShowSheetListView(g_helpGui, g_helpLv)
        g_helpShown := true
    } catch as e {
        ; #region agent log
        CheatSheet_AgentDebugLog("A", "cheat_sheet_gui.ahk:ToggleShortcutHelp", "error",
            Format('{{"msg":"{1}"}}', StrReplace(e.Message, '"', "'")))
        ; #endregion
        g_helpShown := false
    }
}

; ========== Global shortcuts cheat sheet (Win+Alt+Shift+key) ===============
ShowGlobalShortcutsHelp() {
    global g_globalGui, g_globalShown, g_globalSearchEdit, g_globalLv
    global g_cheatSheetGlobalFullProcessed, g_cheatSheetGlobalRows

    if (IsObject(g_globalGui) && g_globalShown) {
        g_globalGui.Hide()
        g_globalShown := false
        return
    }

    ; #region agent log
    CheatSheet_AgentDebugLog("A", "cheat_sheet_gui.ahk:ShowGlobalShortcutsHelp", "entry state",
        Format('{{"globalGui":{1},"globalLv":{2},"globalShown":0}}',
            IsObject(g_globalGui) ? 1 : 0, IsObject(g_globalLv) ? 1 : 0))
    ; #endregion

    normalizedText := NormalizeMojibake(GetGlobalCheatSheetRawText())
    processedText := ProcessCheatSheetText(normalizedText)
    g_cheatSheetGlobalFullProcessed := processedText
    g_cheatSheetGlobalRows := CheatSheet_ParseSheetRows(processedText)
    ; #region agent log
    CheatSheet_AgentDebugLog("B", "cheat_sheet_gui.ahk:ShowGlobalShortcutsHelp", "parsed global rows",
        Format('{{"globalRowCount":{1},"procLen":{2}}}', g_cheatSheetGlobalRows.Length, StrLen(processedText)))
    ; #endregion
    try {
        global g_cheatSheetSuppressFilter
        CheatSheet_EnsureGlobalSheetGui()
        g_cheatSheetSuppressFilter := true
        g_globalSearchEdit.Value := ""
        g_cheatSheetSuppressFilter := false
        CheatSheet_RefreshSheetListView(g_globalLv, g_cheatSheetGlobalRows)
        CheatSheet_ShowSheetListView(g_globalGui, g_globalLv)
        g_globalShown := true
    } catch as e {
        ; #region agent log
        CheatSheet_AgentDebugLog("A", "cheat_sheet_gui.ahk:ShowGlobalShortcutsHelp", "error",
            Format('{{"msg":"{1}"}}', StrReplace(e.Message, '"', "'")))
        ; #endregion
        g_globalShown := false
    }
}

; ========== Hotkey with hold detection ====================================
; Win + Alt + Shift + A with hold detection
#!+a::
{
    global g_helpGui, g_helpShown, g_globalGui, g_globalShown, g_searchAllGui

    ; First check if any cheat sheet is currently open - if so, close it
    if (IsObject(g_helpGui) && g_helpShown) {
        g_helpGui.Hide()
        g_helpShown := false
        return
    }

    if (IsObject(g_globalGui) && g_globalShown) {
        g_globalGui.Hide()
        g_globalShown := false
        return
    }

    if (IsObject(g_searchAllGui)) {
        g_searchAllGui.Hide()
        return
    }

    ; No cheat sheet is open, determine which one to show based on hold time
    static pressTime := 0
    pressTime := A_TickCount

    ; Wait for key release or timeout (increased to accommodate 1s+ holds)
    KeyWait "a", "T1"  ; Wait max 1.5s for key release

    holdTime := A_TickCount - pressTime

    if (holdTime >= 700) {
        ; Long hold (1s+) - show global shortcuts
        ShowGlobalShortcutsHelp()
    } else {
        ; Quick press - show app-specific shortcuts
        ToggleShortcutHelp()
    }
}

; Win+Alt+Shift+/ — search all registered cheat sheets (ListView; double-click row copies line)
#!+/::
{
    ShowSearchAllCheatSheetsGui()
}
