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
                if (GeminiEnterprise_IsEnterpriseHwnd(hwnd, "fast") || GeminiEnterprise_IsEnterpriseHwnd(hwnd, "full")
                || GeminiEnterprise_TryUiaFingerprint(hwnd))
                    siteKey := "Gemini Enterprise"
            } catch {
            }
        }
        if (siteKey = "" && hwnd) {
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
    if (chromeTitle = "Tasks" || InStr(chromeTitle, "Tasks") = 1)
        key := "Tasks"
    if InStr(chromeTitle, "Wikipedia", false) || InStr(chromeTitle, "wikipedia.org", false)
        key := "Wikipedia"
    if IsMercadoLivreActive()
        key := "Mercado Livre"
    if (key = "" && IsShopeeActive())
        key := "Shopee"
    if InStr(chromeTitle, "Gemini Enterprise", false)
        key := "Gemini Enterprise"
    if (key = "" && IsConsumerGeminiChromeTitle(chromeTitle))
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
global CHEAT_SHEET_WIDTH_FRAC := 0.8      ; overlay width as fraction of monitor work area
global CHEAT_SHEET_HEIGHT_FRAC := 0.8     ; overlay height as fraction of monitor work area
global CHEAT_SHEET_WORKAREA_MARGIN := 12  ; px inset from work-area edges (each side)

CheatSheet_GetOverlayFramePad() {
    ; +Border adds non-client pixels beyond Show w/h (client area).
    return DllCall("GetSystemMetrics", "int", 32, "int")  ; SM_CXDLGFRAME
}

CheatSheet_GetDpiScaleAt(x, y) {
    try {
        pt := (y & 0xFFFFFFFF) << 32 | (x & 0xFFFFFFFF)
        hMon := DllCall("User32.dll\MonitorFromPoint", "int64", pt, "uint", 2, "ptr")
        dpiX := 0, dpiY := 0
        if (DllCall("Shcore.dll\GetDpiForMonitor", "ptr", hMon, "int", 0, "uint*", &dpiX, "uint*", &dpiY, "hresult") =
        0
        && dpiX > 0)
            return dpiX / 96.0
    } catch {
    }
    return 1.0
}

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
    try SendMessage(0x1024, 0, CHEAT_SHEET_LV_TEXT, lv) ; LVM_SETTEXTCOLOR
}

CheatSheet_ConfigureSheetListViewColumns(lv, guiW := 1000) {
    col1 := Min(200, Max(80, Round(guiW * 0.20)))
    col2 := Min(280, Max(100, Round(guiW * 0.28)))
    lv.ModifyCol(1, col1)
    lv.ModifyCol(2, col2)
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
    if !IsObject(lv)
        return
    lv.Delete()
    for row in rows
        lv.Add("", row.section, row.shortcut, row.description)
}

CheatSheet_ComputeOverlayBounds(&wx, &wy, &wr, &wb, &guiW, &guiH, &lvH, &outerX, &outerY) {
    global CHEAT_SHEET_WIDTH_FRAC, CHEAT_SHEET_HEIGHT_FRAC, CHEAT_SHEET_WORKAREA_MARGIN
    chromePx := 52  ; filter + spacing in Gui.Show/Move units
    margin := CHEAT_SHEET_WORKAREA_MARGIN
    wx := 0, wy := 0, wr := A_ScreenWidth, wb := A_ScreenHeight
    try GetActiveMonitorWorkArea_StandardBar(&wx, &wy, &wr, &wb)
    workW := wr - wx
    workH := wb - wy
    dpiScale := CheatSheet_GetDpiScaleAt(wx + workW // 2, wy + workH // 2)
    framePadPhysical := Max(2, Round(CheatSheet_GetOverlayFramePad() * dpiScale))
    maxOuterW := workW - 2 * margin
    maxOuterH := workH - 2 * margin
    minClientPhysical := Max(320, Round((chromePx + 200) * dpiScale))
    outerW := Max(minClientPhysical + 2 * framePadPhysical, Min(maxOuterW, Round(workW * CHEAT_SHEET_WIDTH_FRAC)))
    outerH := Max(minClientPhysical + 2 * framePadPhysical, Min(maxOuterH, Round(workH * CHEAT_SHEET_HEIGHT_FRAC)))
    outerX := wx + (workW - outerW) / 2
    outerY := wy + (workH - outerH) / 2
    outerX := Max(wx + margin, Min(outerX, wr - outerW - margin))
    outerY := Max(wy + margin, Min(outerY, wb - outerH - margin))
    ; Work area is physical px; Gui.Show w/h are logical — divide before Show/Move.
    guiW := Max(chromePx + 200, Round((outerW - 2 * framePadPhysical) / dpiScale))
    guiH := Max(chromePx + 200, Round((outerH - 2 * framePadPhysical) / dpiScale))
    lvH := guiH - chromePx
}

CheatSheet_ShowSheetListView(gui, lv, filterEdit := 0) {
    wx := 0, wy := 0, wr := 0, wb := 0, guiW := 0, guiH := 0, lvH := 0, outerX := 0, outerY := 0
    CheatSheet_ComputeOverlayBounds(&wx, &wy, &wr, &wb, &guiW, &guiH, &lvH, &outerX, &outerY)
    mx := IsObject(gui) ? gui.MarginX : 10
    bodyW := Max(200, guiW - 2 * mx)
    if IsObject(filterEdit)
        filterEdit.Move(, , bodyW)
    if IsObject(lv) {
        lv.Move(, , bodyW, lvH)
        CheatSheet_ConfigureSheetListViewColumns(lv, bodyW)
    }
    gui.Show("x" Round(outerX) " y" Round(outerY) " w" guiW " h" guiH)
    if IsObject(filterEdit)
        CheatSheet_DeferFocusSearch(filterEdit)
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
    gui.MarginX := 10
    gui.MarginY := 6
    gui.BackColor := "FFFFFF"
    gui.SetFont("s10 c000000", "Consolas")
    filterEdit := gui.Add("Edit", "xm w400 Section Limit20 BackgroundFFFFFF c000000", "")
    lv := gui.Add("ListView", "xm w400 h520 Grid -Multi +ReadOnly c000000 BackgroundFFFFFF",
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

    ; Toggle off if currently shown
    if (IsObject(g_helpGui) && g_helpShown) {
        g_helpGui.Hide()
        g_helpShown := false
        return
    }

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

    try {
        global g_cheatSheetSuppressFilter
        CheatSheet_EnsureAppSheetGui()
        g_cheatSheetSuppressFilter := true
        g_helpSearchEdit.Value := ""
        g_cheatSheetSuppressFilter := false
        CheatSheet_RefreshSheetListView(g_helpLv, g_cheatSheetAppRows)
        CheatSheet_HideOpeningIndicator()
        CheatSheet_ShowSheetListView(g_helpGui, g_helpLv, g_helpSearchEdit)
        g_helpShown := true
    } catch {
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

    normalizedText := NormalizeMojibake(GetGlobalCheatSheetRawText())
    processedText := ProcessCheatSheetText(normalizedText)
    g_cheatSheetGlobalFullProcessed := processedText
    g_cheatSheetGlobalRows := CheatSheet_ParseSheetRows(processedText)
    try {
        global g_cheatSheetSuppressFilter
        CheatSheet_EnsureGlobalSheetGui()
        g_cheatSheetSuppressFilter := true
        g_globalSearchEdit.Value := ""
        g_cheatSheetSuppressFilter := false
        CheatSheet_RefreshSheetListView(g_globalLv, g_cheatSheetGlobalRows)
        CheatSheet_HideOpeningIndicator()
        CheatSheet_ShowSheetListView(g_globalGui, g_globalLv, g_globalSearchEdit)
        g_globalShown := true
    } catch {
        g_globalShown := false
    }
}

CheatSheet_ShowOpeningIndicator() {
    try StandardLoadingBar_Show("⏳ Opening cheat sheet...", BANNER_ACCENT_INTERMEDIATE, { passive: false })
}

CheatSheet_HideOpeningIndicator() {
    try StandardLoadingBar_Hide(0)
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

    ; Loading Indication while waiting for tap vs hold and while resolving/building the overlay.
    CheatSheet_ShowOpeningIndicator()
    try {
        static pressTime := 0
        pressTime := A_TickCount

        ; Wait for key release or timeout (increased to accommodate 1s+ holds)
        KeyWait "a", "T1"  ; Wait max 1.5s for key release

        holdTime := A_TickCount - pressTime

        if (holdTime >= 700) {
            try StandardLoadingBar_Update("⏳ Opening global shortcuts...", BANNER_ACCENT_INTERMEDIATE)
            ShowGlobalShortcutsHelp()
        } else {
            try StandardLoadingBar_Update("⏳ Opening app shortcuts...", BANNER_ACCENT_INTERMEDIATE)
            ToggleShortcutHelp()
        }
    } finally {
        CheatSheet_HideOpeningIndicator()
    }
}

; Win+Alt+Shift+/ — search all registered cheat sheets (ListView; double-click row copies line)
#!+/::
{
    ShowSearchAllCheatSheetsGui()
}
