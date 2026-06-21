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

; Raw text for long-hold global cheat sheet (also used by SearchCheatSheets).
GLOBAL_CHEAT_SHEET_RAW := "
(
    [Win+Alt+Shift] - PRIMARY triple modifier (most used for system-wide shortcuts)
        [Ctrl+Alt+Win] - SECONDARY triple modifier
    
    === AVAILABLE SECONDARY (Ctrl+Alt+Win) SLOTS ===
    [Ctrl+Alt+Win+N] > TEMPORARY — M365 Copilot auto-continue: send "continue", wait for Stop generating, loop (toggle off with same chord)
    [Ctrl+Alt+Win+O] > Evidence search loop — CSV row substring → PDF find; stop saves not-found rows to data/evidence_not_found.csv + 10s report (VSCodeEvidenceSearch.ahk; toggle)
    Letters available: T, U

    Shift+CAW: A/S/D/F/Q/W/E/R (+B debug, +1/Z/G fallbacks) used for window management; other Shift+letters unassigned.
    [Ctrl+Alt+Win+1] > Available
    [Ctrl+Alt+Win+G] > RESERVED — Handy: cancel dictation (define in Handy only; not bound in AHK)
    [Ctrl+Alt+Win+L] > {AI_PROVIDER} D2C direct submit (Utils.ahk; ZMK hold on L key)
    [Ctrl+Alt+Win+V] > Maximize active window (WindowManagement.ahk; ZMK hold on minimize/close key)
    [Ctrl+Alt+Win+X] > Snap 50/50: half-width active window + pair recent window in other half (WindowManagement.ahk)
    [Ctrl+Alt+Win+Z] > Window tools [1]: maximize lone visible window per monitor (WindowManagement.ahk; also Win+Alt+Shift+W → 1)
    [Ctrl+Alt+Win+6] > Window tools [2]: hidden background window list (WindowManagement.ahk; also Win+Alt+Shift+W → 2)
    [Ctrl+Alt+Win+Y] > Window tools [3]: tile background windows (WindowManagement.ahk; also Win+Alt+Shift+W → 3)
    [Ctrl+Alt+Win+P] > Window tools [4]: exit F11 fullscreen (WindowManagement.ahk; also Win+Alt+Shift+W → 4)
    [Ctrl+Alt+Win+0] > Project Quick Selector (opens project folder in Cursor)
    [Ctrl+Alt+Win+1] > Cursor AI quick action (Project Selector + Selection Mode)
    [Ctrl+Alt+Win+2] > Quick Update to Your Scripts (HotStrings macro)
    [Ctrl+Alt+Win+3] > Toggle Outlook and Teams (HotStrings macro)
    [Ctrl+Alt+Win+5] > Clean the Clipboard (HotStrings macro)
    [Ctrl+Alt+Win+7] > Mark Last Clip as Favorite (HotStrings macro; same as Ctrl+Alt+Win+J if 7 chord fails on keyboard)
    [Ctrl+Alt+Win+J] > Mark Last Clip as Favorite (HotStrings macro; alternate for keyboards that ghost Ctrl+Alt+Win+7)
    [Ctrl+Alt+Win+8] > Moves Desktop to Recycle Bin (HotStrings macro)
    [Ctrl+Alt+Win+9] > Handy: Cohere Portuguese (model slot 4)
    [Ctrl+Alt+Win+B] > Handy: Cohere English (model slot 3)
    
    === MAIN KEY COMBINATIONS ===
    [Symbol Layer] Win+Alt+Shift - Primary combination
    [Window Management] Ctrl+Alt+Win - Secondary combination
    
    [Alt+P] Ope clip angel
    
    [Win+Alt+Shift+L] > Outlook Copilot shortcut modal (1–9); global hotkey
    
    === CURSOR ===
    [Win+Alt+Shift+N] > Context file browser (paste path)
    
    [Win+Alt+Shift+J] > Fast Copy: tap on/off (count Ctrl+C / PrtSc / Alt+PrtSc, then paste N); hold 700ms+ repeats last N (Clip Angel)
    
    === SPOTIFY ===
    [Win+Alt+Shift+S] > Opens or activates Spotify
    
    === CLIP ANGEL ===
    [Win+Alt+Shift+1] > Send top list item from Clip Angel
    
    === AI CHAT (Chrome) ===
    [Win+Alt+Shift+I] > Opens {AI_PROVIDER}
    [Win+Alt+Shift+8] > Get word pronunciation, definition, and Portuguese translation ({AI_PROVIDER})
    [Win+Alt+Shift+O] > Read aloud the last message in {AI_PROVIDER}
    [Win+Alt+Shift+P] > Copy the last message in {AI_PROVIDER}
    [Win+Alt+Shift+7] > Copy selected text and read aloud ({AI_PROVIDER})
    
    === HANDY DICTATION ===
    [Win+Alt+Shift+0] > Start/stop dictation (transcription to clipboard)
    [Ctrl+Alt+Win+G] > Cancel dictation (Handy — user-defined; reserved in cheat sheet, not in AHK)
    [Ctrl+Alt+Win+9] > Handy: Cohere Portuguese (picker slot 4; same as Win+Alt+Shift+C then 4)
    [Ctrl+Alt+Win+B] > Handy: Cohere English (picker slot 3; same as Win+Alt+Shift+C then 3)
    [Win+Alt+Shift+C] > AI model picker (Handy): 1 Parakeet V2, 2 Parakeet V3, 3 Cohere English, 4 Cohere Portuguese
    
    === YOUTUBE ===
    [Win+Alt+Shift+H] > Activates Youtube
    
    === GOOGLE ===
    [Win+Alt+Shift+F] > Opens Google
    
    === CURSOR ===
    [Win+Alt+Shift+,] > Opens or activates Cursor
    [Win+Alt+Shift+C] > Handy AI model picker (see HANDY DICTATION)
    
    === OUTLOOK ===
    [Win+Alt+Shift+B] > Open mail
    [Win+Alt+Shift+V] > Open Reminder
    [Win+Alt+Shift+G] > Open calendar
    [Win+Alt+Shift+D] > Voice aloud the email
    
    === MICROSOFT TEAMS ===
    [Win+Alt+Shift+R] > New conversation
    [Win+Alt+Shift+5] > Toggle Mute (meeting)
    [Win+Alt+Shift+4] > Toggle camera (meeting)
    [Win+Alt+Shift+T] > Screen share (meeting)
    [Win+Alt+Shift+2] > Exit meeting
    [Win+Alt+Shift+E] > Select the chats window
    [Win+Alt+Shift+3] > Select the meeting window
    
    === WHATSAPP ===
    [Win+Alt+Shift+Z] > Opens WhatsApp
    
    === WINDOWS ===
    [Win+Alt+Shift+6] > Minimizes windows
    [Win+Alt+Shift+M] > Maximizes the current window
    [Win+Alt+Shift+W] > Window tools menu: [1] maximize lone; [2] hidden background list; [3] tile background (≤12 total, ≤3/monitor); [4] exit F11 fullscreen — direct CAW: Z=[1], 6=[2], Y=[3], P=[4]
    [Win+Alt+Shift+Y] > Focus Mode: Black out all monitors except the one with the active window (toggle)
    
    === WINDOW MANAGEMENT (Ctrl+Alt+Win) ===
    [Ctrl+Alt+Win+X] > Snap 50/50: half-width active window + pair recent window in other half
    [Ctrl+Alt+Win+Z] > Window tools [1]: maximize lone visible window per monitor (also Win+Alt+Shift+W → 1)
    [Ctrl+Alt+Win+6] > Window tools [2]: hidden background window list (also Win+Alt+Shift+W → 2)
    [Ctrl+Alt+Win+Y] > Window tools [3]: tile background windows (also Win+Alt+Shift+W → 3)
    [Ctrl+Alt+Win+P] > Window tools [4]: exit F11 fullscreen (also Win+Alt+Shift+W → 4)
    [Ctrl+Alt+Win+A] > Move window to monitor 1 (left-most)
    [Ctrl+Alt+Win+S] > Move window to monitor 2
    [Ctrl+Alt+Win+D] > Move window to monitor 3
    [Ctrl+Alt+Win+F] > Move window to monitor 4
    [Ctrl+Alt+Win+Shift+A] > Close window on monitor 1
    [Ctrl+Alt+Win+Shift+S] > Close window on monitor 2
    [Ctrl+Alt+Win+Shift+D] > Close window on monitor 3
    [Ctrl+Alt+Win+Shift+F] > Close window on monitor 4
    [Ctrl+Alt+Win+Q] > Cycle windows on monitor 1
    [Ctrl+Alt+Win+W] > Cycle windows on monitor 2
    [Ctrl+Alt+Win+E] > Cycle windows on monitor 3
    [Ctrl+Alt+Win+R] > Cycle windows on monitor 4
    [Ctrl+Alt+Win+Shift+Q] > Minimize window on monitor 1
    [Ctrl+Alt+Win+Shift+W] > Minimize window on monitor 2
    [Ctrl+Alt+Win+Shift+E] > Minimize window on monitor 3
    [Ctrl+Alt+Win+Shift+R] > Minimize window on monitor 4
    
    === COMMAND PALETTE BOOKMARKS ===
    [Ctrl+Alt+Win+M] > Add bookmark (Command Palette Bookmark extension)
    
    === GENERAL ===
    [Win+Alt+Shift+U] > Quick string shortcuts
    [Ctrl+Alt+Win+4] > Send AI Text Optimizer prompt to {AI_PROVIDER} (same as Win+Alt+Shift+U then L, 4)
    [Win+Alt+Shift+Q] > Jump mouse on the middle
    [Win+Alt+Shift+X] > Peek PDF (tap) / Set PDF path (hold 700ms+)
    [Win+Alt+Shift+→] > Show square selector (right direction)
    [Win+Alt+Shift+←] > Show square selector (left direction)
    [Win+Alt+Shift+↓] > Show square selector (down direction)
    [Win+Alt+Shift+↑] > Show square selector (up direction)
    [Win+Alt+Shift+.] > Clip Angel (copy, paste, and quit)
    
    === COMMAND PALETTE ===
    [Win+Ctrl+Alt+Y] > Command Palette - File search
    [Shift+D] > Command Palette (active): exclude current bookmark (confirm)
    
    === SHORTCUTS ===
    [Win+Alt+Shift+A] > Show app-specific shortcuts (quick press)
    [Win+Alt+Shift+A] > Show global shortcuts (hold 700ms+)
    [Win+Alt+Shift+/] > Search all cheat sheets (cross-context)
    
    === WIKIPEDIA ===
    [Win+Alt+Shift+K] > Opens or activates Wikipedia
)"

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
