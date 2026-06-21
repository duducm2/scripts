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
global g_helpCheatCtrl := 0
global g_globalSearchEdit := 0
global g_globalCheatCtrl := 0
global g_cheatSheetAppFullProcessed := ""
global g_cheatSheetGlobalFullProcessed := ""
global g_searchAllGui := 0

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

CheatSheet_ResizeBody(editCtrl, gui, fontLinePx := 18, minH := 200, lineCountSource := "", chromeAboveBodyPx := 48) {
    text := lineCountSource != "" ? lineCountSource : editCtrl.Value
    lineCnt := StrLen(text) ? StrSplit(text, "`n").Length : 1
    controlHeight := lineCnt * fontLinePx + 10
    if (controlHeight < minH)
        controlHeight := minH
    ; Same work-rect as standard banners (Utils.GetActiveMonitorWorkArea_StandardBar).
    GetActiveMonitorWorkArea_StandardBar(&ml, &mt, &mr, &mb)
    workH := mb - mt
    margin := 6
    maxHeight := workH - chromeAboveBodyPx - margin
    if (maxHeight < 80)
        maxHeight := 80
    if (controlHeight > maxHeight)
        controlHeight := maxHeight
    ; Pixel height for body; Custom RichEdit50W may not use rN row metrics like a built-in Edit.
    editCtrl.Move(, , 1000, controlHeight)
    gui.Show("AutoSize Hide")
    ; If chrome estimate was low, total GUI height can still exceed work area — shrink body once.
    gui.GetPos(, , &gw, &gh)
    maxGuiH := workH - margin
    if (gh > maxGuiH) {
        newH := controlHeight - (gh - maxGuiH)
        if (newH < 80)
            newH := 80
        editCtrl.Move(, , 1000, newH)
        gui.Show("AutoSize Hide")
    }
    CenterGuiOnActiveMonitor(gui)
    gui.Show()
}

CheatSheet_OnAppFilterChanged(*) {
    global g_helpSearchEdit, g_helpCheatCtrl, g_cheatSheetAppFullProcessed, g_helpGui, g_helpShown
    if (!IsObject(g_helpCheatCtrl) || !IsObject(g_helpSearchEdit))
        return
    q := Trim(g_helpSearchEdit.Value)
    body := g_cheatSheetAppFullProcessed
    displayBody := ""
    if (q = "") {
        displayBody := body
    } else {
        displayBody := CheatSheet_BuildFilteredBodyWithSections(body, q)
    }
    CheatSheet_RichSetProcessedBody(g_helpCheatCtrl, displayBody)
    CheatSheet_ResizeBody(g_helpCheatCtrl, g_helpGui, 18, 200, displayBody, 48)
    CheatSheet_DeferFocusSearch(g_helpSearchEdit)
}

CheatSheet_OnGlobalFilterChanged(*) {
    global g_globalSearchEdit, g_globalCheatCtrl, g_cheatSheetGlobalFullProcessed, g_globalGui, g_globalShown
    if (!IsObject(g_globalCheatCtrl) || !IsObject(g_globalSearchEdit))
        return
    q := Trim(g_globalSearchEdit.Value)
    body := g_cheatSheetGlobalFullProcessed
    displayBody := ""
    if (q = "") {
        displayBody := body
    } else {
        displayBody := CheatSheet_BuildFilteredBodyWithSections(body, q)
    }
    CheatSheet_RichSetProcessedBody(g_globalCheatCtrl, displayBody)
    CheatSheet_ResizeBody(g_globalCheatCtrl, g_globalGui, 16, 180, displayBody, 46)
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
    global g_helpGui, g_helpShown, g_helpSearchEdit, g_helpCheatCtrl, g_cheatSheetAppFullProcessed

    ; Toggle off if currently shown
    if (IsObject(g_helpGui) && g_helpShown) {
        g_helpGui.Hide()
        g_helpShown := false
        return
    }

    text := NormalizeMojibake(GetCheatSheetText())
    usedFallback := false
    if (text = "") {
        exe := WinGetProcessName("A")
        text := "No cheat-sheet registered for:`n" exe
        usedFallback := true
    }

    if !IsObject(g_helpGui) {
        CheatSheet_EnsureRichDll()
        g_helpGui := Gui("+AlwaysOnTop -Caption +ToolWindow +Border +Owner +LastFound")
        g_helpGui.BackColor := "000000"
        g_helpGui.SetFont("s11 cFFFF00", "Consolas")
        g_helpSearchEdit := g_helpGui.Add("Edit", "xm w1000 Section Limit20", "")
        g_helpCheatCtrl := g_helpGui.Add("Custom",
            "ClassRichEdit50W xs+0 y+4 +0x44 +Multi -E0x200 +VScroll -HScroll -Border Background000000 w1000 r12")
        g_helpSearchEdit.OnEvent("Change", CheatSheet_OnAppFilterChanged)
        g_helpGui.OnEvent("Escape", CheatSheet_OnEscapeApp)
    }

    processedText := ProcessCheatSheetText(text)
    g_cheatSheetAppFullProcessed := processedText
    g_helpSearchEdit.Value := ""
    CheatSheet_RichSetProcessedBody(g_helpCheatCtrl, processedText)
    CheatSheet_ResizeBody(g_helpCheatCtrl, g_helpGui, 18, 200, processedText, 48)
    g_helpShown := true
    CheatSheet_DeferFocusSearch(g_helpSearchEdit)
}

; ========== Global shortcuts cheat sheet (Win+Alt+Shift+key) ===============
ShowGlobalShortcutsHelp() {
    global g_globalGui, g_globalShown, g_globalSearchEdit, g_globalCheatCtrl, g_cheatSheetGlobalFullProcessed

    if (IsObject(g_globalGui) && g_globalShown) {
        g_globalGui.Hide()
        g_globalShown := false
        return
    }

    if !IsObject(g_globalGui) {
        CheatSheet_EnsureRichDll()
        g_globalGui := Gui("+AlwaysOnTop -Caption +ToolWindow +Border +Owner +LastFound")
        g_globalGui.BackColor := "000000"
        g_globalGui.SetFont("s9 c00BFFF", "Consolas")
        g_globalSearchEdit := g_globalGui.Add("Edit", "xm w1000 Section Limit20", "")
        g_globalCheatCtrl := g_globalGui.Add("Custom",
            "ClassRichEdit50W xs+0 y+4 +0x44 +Multi +VScroll -HScroll -Border Background000000 w1000 r12")
        g_globalSearchEdit.OnEvent("Change", CheatSheet_OnGlobalFilterChanged)
        g_globalGui.OnEvent("Escape", CheatSheet_OnEscapeGlobal)
    }

    normalizedText := NormalizeMojibake(GetGlobalCheatSheetRawText())
    processedText := ProcessCheatSheetText(normalizedText)
    g_cheatSheetGlobalFullProcessed := processedText
    g_globalSearchEdit.Value := ""
    CheatSheet_RichSetProcessedBody(g_globalCheatCtrl, processedText)
    CheatSheet_ResizeBody(g_globalCheatCtrl, g_globalGui, 16, 180, processedText, 46)
    g_globalShown := true
    CheatSheet_DeferFocusSearch(g_globalSearchEdit)
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
