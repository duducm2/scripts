; =============================================================================
; WindowManagement module: cursor_window_select.ahk
; Cursor window selection (within project selector)
; Extracted verbatim from WindowManagement.ahk; loaded via #include into the
; WindowManagement.ahk process, which remains the entry point / source of truth.
; =============================================================================

; =============================================================================
; Cursor Window Selection (within Project Selector)
; =============================================================================

; Handler for Cursor window selection
HandleCursorWindowSelection(targetHwnd, allCursorWindows) {
    global g_ProjectSelectorActive

    ; Only process if selector is active
    if (!g_ProjectSelectorActive) {
        return
    }

    ; Iterate through all windows and close those that don't match target
    for hwnd in allCursorWindows {
        if (hwnd != targetHwnd) {
            try {
                WinClose("ahk_id " . hwnd)
            } catch {
                ; Silently ignore if window close fails
            }
        }
    }

    ; Activate the target window
    try {
        WinActivate("ahk_id " . targetHwnd)
        WinWaitActive("ahk_id " . targetHwnd, , 1)
    } catch {
        ShowNotification_WM("Error: Target window not found.")
    }

    ; Cleanup the selector
    CleanupProjectSelector()
}

; Factory function to create a handler for Cursor window selection
CreateCursorWindowSelectionHandler(char) {
    ; Return a function that captures the char value at creation time
    return (*) => HandleCursorWindowSelectionByChar(char)
}

; Handler for character key press in Cursor window selector sub-menu
HandleCursorWindowSelectionByChar(char) {
    global g_CursorWindowMap, g_ProjectSelectorActive

    ; Only process if selector is active (cursor window selector inherits from project selector state)
    if (!g_ProjectSelectorActive) {
        return
    }

    ; Get the HWND for this character
    targetHwnd := g_CursorWindowMap.Get(char, "")
    if (targetHwnd = "") {
        ; Try lowercase if uppercase
        targetHwnd := g_CursorWindowMap.Get(StrLower(char), "")
    }

    if (targetHwnd != "") {
        allCursorWindows := []
        if (WM_UsesAutomationDaemon()) {
            try {
                for w in WMIPC_GetCursorWindows()
                    allCursorWindows.Push(Integer(w["hwnd"]))
            } catch {
            }
        }
        if (allCursorWindows.Length = 0)
            allCursorWindows := WinGetList("ahk_exe Cursor.exe")
        HandleCursorWindowSelection(targetHwnd, allCursorWindows)

        ; Also cleanup the cursor window selector GUI if it exists
        global g_CursorWindowSelectorGui
        if (IsObject(g_CursorWindowSelectorGui)) {
            try {
                g_CursorWindowSelectorGui.Destroy()
                g_CursorWindowSelectorGui := false
            } catch {
                ; Ignore
            }
        }

        ; Disable cursor window hotkeys
        global g_CursorWindowHotkeyHandlers
        for handler in g_CursorWindowHotkeyHandlers {
            try {
                charToDisable := handler.char
                ; Handle special VK codes
                if (charToDisable = ",") {
                    Hotkey("vkBC", "Off")
                } else if (charToDisable = ".") {
                    Hotkey("vkBE", "Off")
                } else {
                    Hotkey(charToDisable, "Off")
                    ; Also disable uppercase for lowercase letters
                    if (RegExMatch(charToDisable, "^[a-z]$")) {
                        Hotkey(StrUpper(charToDisable), "Off")
                    }
                }
            } catch {
                ; Silently ignore errors
            }
        }
        g_CursorWindowHotkeyHandlers := []
        g_CursorWindowMap := Map()
    }
}

; Handler for Escape key in Cursor window selector sub-menu
HandleCursorWindowSelectorEscape(*) {
    global g_CursorWindowSelectorGui, g_CursorWindowHotkeyHandlers, g_CursorWindowMap

    ; Close and destroy GUI
    if (IsObject(g_CursorWindowSelectorGui)) {
        try {
            g_CursorWindowSelectorGui.Destroy()
            g_CursorWindowSelectorGui := false
        } catch {
            ; Ignore
        }
    }

    ; Disable all character hotkeys
    for handler in g_CursorWindowHotkeyHandlers {
        try {
            char := handler.char
            ; Handle special VK codes
            if (char = ",") {
                Hotkey("vkBC", "Off")
            } else if (char = ".") {
                Hotkey("vkBE", "Off")
            } else {
                Hotkey(char, "Off")
                ; Also disable uppercase for lowercase letters
                if (RegExMatch(char, "^[a-z]$")) {
                    Hotkey(StrUpper(char), "Off")
                }
            }
        } catch {
            ; Silently ignore errors
        }
    }

    ; Restore Escape callback to project selector (main selector still open)
    g_OnEscapePressed := HandleProjectEscape

    ; Clear handlers and map
    g_CursorWindowHotkeyHandlers := []
    g_CursorWindowMap := Map()
}

; Show Cursor window selector sub-menu GUI
ShowCursorWindowSelectorSubMenu() {
    global g_CursorWindowSelectorGui, g_CursorWindowMap, g_CursorWindowHotkeyHandlers
    global g_ProjectCharSequence, g_ProjectSelectorActive, g_Projects, g_ProjectCategories
    global IS_WORK_ENVIRONMENT

    ; Only show if project selector is active
    if (!g_ProjectSelectorActive) {
        return
    }

    ; Get all Cursor windows (daemon cache or legacy)
    cursorWindows := []
    if (WM_UsesAutomationDaemon()) {
        try {
            for w in WMIPC_GetCursorWindows()
                cursorWindows.Push(Integer(w["hwnd"]))
        } catch {
        }
    }
    if (cursorWindows.Length = 0)
        cursorWindows := WinGetList("ahk_exe Cursor.exe")

    if (cursorWindows.Length = 0) {
        ShowNotification_WM("No Cursor windows found.")
        return
    }

    if (cursorWindows.Length = 1) {
        try {
            WinActivate("ahk_id " . cursorWindows[1])
            WinWaitActive("ahk_id " . cursorWindows[1], , 1)
        } catch {
            ShowNotification_WM("Error: Target window not found.")
        }
        CleanupProjectSelector()
        return
    }

    ; Build project index to character mapping (same as ShowProjectSelector)
    projectIndexToChar := Map()
    projectIndexToCategory := Map()

    ; Build map of project index to category
    loop g_Projects.Length {
        projectIndex := A_Index
        project := g_Projects[projectIndex]
        category := project.HasProp("category") ? project.category : "Personal"
        projectIndexToCategory[projectIndex] := category
    }

    charIndex := 1

    ; Assign characters sequentially within each category (same logic as ShowProjectSelector)
    for category in g_ProjectCategories {
        ; Find all project indices in this category
        categoryProjectIndices := []
        for projectIndex, cat in projectIndexToCategory {
            if (cat = category) {
                categoryProjectIndices.Push(projectIndex)
            }
        }

        ; Assign characters to projects in this category
        for projectIndex in categoryProjectIndices {
            project := g_Projects[projectIndex]

            ; Skip empty placeholders
            if (project.name = "" && project.path = "" && project.workPath = "") {
                charIndex++
                continue
            }

            ; Check if we have a character available
            if (charIndex > g_ProjectCharSequence.Length) {
                break
            }

            char := g_ProjectCharSequence[charIndex]

            ; Skip character "3" - it's reserved for preview window activation
            if (char = "3") {
                charIndex++
                if (charIndex > g_ProjectCharSequence.Length) {
                    break
                }
                char := g_ProjectCharSequence[charIndex]
            }

            projectIndexToChar[projectIndex] := char
            charIndex++
        }
    }

    ; Helper function to check if a window title matches a project path
    WindowMatchesProject(winTitle, projectPath) {
        if (projectPath = "") {
            return false
        }
        matchSegments := ExtractProjectMatchSegments(projectPath)
        for segment in matchSegments {
            if (InStr(winTitle, segment)) {
                return true
            }
        }
        return false
    }

    ; Helper function to get matching project index for a window
    GetMatchingProjectIndex(winTitle) {
        loop g_Projects.Length {
            projectIndex := A_Index
            project := g_Projects[projectIndex]

            ; Skip empty placeholders
            if (project.name = "" && project.path = "" && project.workPath = "") {
                continue
            }

            ; Select path based on environment
            projectPath := IS_WORK_ENVIRONMENT ? project.workPath : project.path

            ; If work environment but no workPath set, fall back to personal path
            if (IS_WORK_ENVIRONMENT && projectPath = "") {
                projectPath := project.path
            }

            ; Check if window matches this project
            if (WindowMatchesProject(winTitle, projectPath)) {
                return projectIndex
            }
        }
        return 0
    }

    ; Build list of windows with their assigned keys
    windowsWithKeys := []
    usedKeys := Map()
    usedProjectIndices := Map()

    ; First pass: assign keys to windows that match projects
    for hwnd in cursorWindows {
        try {
            winTitle := WinGetTitle("ahk_id " . hwnd)
            if (winTitle = "") {
                winTitle := "Untitled"
            }

            ; Check if this window matches a project
            matchingProjectIndex := GetMatchingProjectIndex(winTitle)

            if (matchingProjectIndex > 0 && projectIndexToChar.Has(matchingProjectIndex)) {
                char := projectIndexToChar[matchingProjectIndex]
                ; Only use this key once
                if (!usedKeys.Has(char) && !usedProjectIndices.Has(matchingProjectIndex)) {
                    windowsWithKeys.Push({ hwnd: hwnd, title: winTitle, char: char, projectIndex: matchingProjectIndex })
                    usedKeys[char] := true
                    usedProjectIndices[matchingProjectIndex] := true
                } else {
                    ; Mark as unassigned for now, will assign in second pass
                    windowsWithKeys.Push({ hwnd: hwnd, title: winTitle, char: "", projectIndex: matchingProjectIndex })
                }
            } else {
                ; No project match, will assign in second pass
                windowsWithKeys.Push({ hwnd: hwnd, title: winTitle, char: "", projectIndex: 0 })
            }
        } catch {
            ; Skip windows we can't access
            continue
        }
    }

    ; Second pass: assign remaining keys to unmatched windows
    charIndex := 1
    for window in windowsWithKeys {
        ; Skip if already assigned
        if (window.char != "") {
            continue
        }

        ; Find next available character
        while (charIndex <= g_ProjectCharSequence.Length) {
            char := g_ProjectCharSequence[charIndex]

            ; Skip character "3" - reserved for preview windows
            if (char = "3") {
                charIndex++
                continue
            }

            ; Check if this character is already used
            if (!usedKeys.Has(char)) {
                window.char := char
                usedKeys[char] := true
                charIndex++
                break
            }

            charIndex++
        }
    }

    ; Filter out windows without assigned keys
    filteredWindows := []
    for window in windowsWithKeys {
        if (window.char != "") {
            filteredWindows.Push(window)
        }
    }

    ; Clear window map
    g_CursorWindowMap := Map()

    ; Get active monitor for positioning
    activeWin := 0
    try {
        activeWin := WinGetID("A")
    } catch {
        activeWin := 0
    }

    ; Default to primary monitor work area
    MonitorGetWorkArea(1, &monitorLeft, &monitorTop, &monitorRight, &monitorBottom)
    monitorWidth := monitorRight - monitorLeft
    monitorHeight := monitorBottom - monitorTop

    ; If we have an active window, find which monitor contains its center
    if (activeWin && activeWin != 0) {
        rect := Buffer(16, 0)
        if (DllCall("GetWindowRect", "ptr", activeWin, "ptr", rect)) {
            ; Calculate window center
            winLeft := NumGet(rect, 0, "int")
            winTop := NumGet(rect, 4, "int")
            winRight := NumGet(rect, 8, "int")
            winBottom := NumGet(rect, 12, "int")

            centerX := winLeft + (winRight - winLeft) // 2
            centerY := winTop + (winBottom - winTop) // 2

            ; Find which monitor contains the window center
            monitorCount := MonitorGetCount()
            loop monitorCount {
                idx := A_Index
                MonitorGetWorkArea(idx, &l, &t, &r, &b)
                if (centerX >= l && centerX <= r && centerY >= t && centerY <= b) {
                    monitorLeft := l
                    monitorTop := t
                    monitorRight := r
                    monitorBottom := b
                    monitorWidth := r - l
                    monitorHeight := b - t
                    break
                }
            }
        }
    }

    ; Create GUI - non-activating so it doesn't steal focus, standard background
    g_CursorWindowSelectorGui := Gui("+AlwaysOnTop +ToolWindow +E0x08000000", "Focus Cursor Window")
    fontSize := (monitorHeight < 800) ? 9 : 10
    g_CursorWindowSelectorGui.SetFont("s" . fontSize, "Segoe UI")
    g_CursorWindowSelectorGui.MarginX := 15
    g_CursorWindowSelectorGui.MarginY := 10

    ; Build display text
    displayText := "=== FOCUS CURSOR WINDOW ===`n`n"

    for window in filteredWindows {
        ; Map character to HWND
        g_CursorWindowMap[window.char] := window.hwnd

        ; Add to display
        displayText .= "[" . window.char . "] " . window.title . "`n"
    }

    displayText .= "`n[ESC] Cancel"

    ; Calculate text dimensions
    baseWidth := 400
    lineHeight := fontSize + 6
    lineCount := StrSplit(displayText, "`n").Length
    textControlHeight := lineCount * lineHeight + 10

    ; Add text control
    g_CursorWindowSelectorGui.Add("Text", "w" . (baseWidth - 30), displayText)

    ; Add close button
    closeBtn := g_CursorWindowSelectorGui.Add("Button", "w80 Center", "Close")
    closeBtn.OnEvent("Click", (*) => HandleCursorWindowSelectorEscape())

    ; Calculate total height
    totalHeight := 20 + textControlHeight + 40 + 10

    ; Calculate center position
    marginX := 20
    marginY := 20
    guiX := monitorLeft + (monitorWidth - baseWidth) // 2
    guiY := monitorTop + (monitorHeight - totalHeight) // 2

    ; Ensure GUI stays within monitor bounds
    if (guiX < monitorLeft + marginX)
        guiX := monitorLeft + marginX
    if (guiY < monitorTop + marginY)
        guiY := monitorTop + marginY
    if (guiX + baseWidth > monitorLeft + monitorWidth - marginX)
        guiX := monitorLeft + monitorWidth - baseWidth - marginX
    if (guiY + totalHeight > monitorTop + monitorHeight - marginY)
        guiY := monitorTop + monitorHeight - totalHeight - marginY

    ; Show GUI
    g_CursorWindowSelectorGui.Show("NA w" . baseWidth . " h" . totalHeight . " x" . guiX . " y" . guiY)

    ; Clear handlers array
    g_CursorWindowHotkeyHandlers := []

    ; Enable hotkeys for assigned characters
    for window in filteredWindows {
        char := window.char

        ; Create handler
        handler := CreateCursorWindowSelectionHandler(char)

        ; Store handler for cleanup
        g_CursorWindowHotkeyHandlers.Push({ char: char, handler: handler })

        ; Enable hotkey
        try {
            ; Handle special characters that need VK codes
            if (char = ",") {
                Hotkey("vkBC", handler, "On")
            } else if (char = ".") {
                Hotkey("vkBE", handler, "On")
            } else {
                Hotkey(char, handler, "On")
                ; Also enable uppercase for lowercase letters
                if (RegExMatch(char, "^[a-z]$")) {
                    Hotkey(StrUpper(char), handler, "On")
                }
            }
        } catch {
            ; Silently ignore if we can't create hotkey for this character
        }
    }

    ; Switch Escape callback to cursor window selector (project selector still open)
    g_OnEscapePressed := HandleCursorWindowSelectorEscape
}

; Handler for Cursor window selection trigger (character "c" in project selector)
HandleCursorWindowSelectionTrigger(*) {
    global g_ProjectSelectorActive

    ; Only process if selector is active
    if (!g_ProjectSelectorActive) {
        return
    }

    ; Show the cursor window selector sub-menu
    ShowCursorWindowSelectorSubMenu()
}

; Show project selector GUI
ShowProjectSelector() {
    global g_ProjectSelectorGui, g_ProjectSelectorActive, g_Projects
    global g_ProjectHotkeyHandlers
    global g_HotstringSelectorGui, g_HotstringSelectorActive
    global g_HS_SelectorOpenFile_WM, g_HS_SelectorCloseRequestFile_WM

    ; Close existing GUI if open
    if (g_ProjectSelectorActive && IsObject(g_ProjectSelectorGui)) {
        CleanupProjectSelector()
        Sleep 50
    }

    ; In-process mutual exclusion: if the Hotstring Selector is active, close it first
    try {
        if (IsSet(g_HotstringSelectorActive) && g_HotstringSelectorActive && IsObject(g_HotstringSelectorGui)) {
            CleanupHotstringSelector()
            Sleep 50
        }
    } catch {
        ; Ignore failures – project selector should still open
    }

    ; Cross-process mutual exclusion: if a Hotstring Selector sentinel exists in another process,
    ; request it to close via hs_selector_close_request.
    try {
        if (FileExist(g_HS_SelectorOpenFile_WM)) {
            try FileAppend("", g_HS_SelectorCloseRequestFile_WM)
            catch {
            }
            Sleep 50
        }
    } catch {
        ; Ignore IPC failures – project selector should still open
    }

    ; Check if we have projects configured
    if (g_Projects.Length = 0) {
        ShowNotification_WM("No projects configured.")
        return
    }

    ; Get monitor dimensions early for responsive sizing
    activeWin := 0
    try {
        activeWin := WinGetID("A")
    } catch {
        activeWin := 0
    }

    ; Default to primary monitor work area
    MonitorGetWorkArea(1, &monitorLeft, &monitorTop, &monitorRight, &monitorBottom)
    monitorWidth := monitorRight - monitorLeft
    monitorHeight := monitorBottom - monitorTop

    ; If we have an active window, find which monitor contains its center
    if (activeWin && activeWin != 0) {
        rect := Buffer(16, 0)
        if (DllCall("GetWindowRect", "ptr", activeWin, "ptr", rect)) {
            ; Calculate window center
            winLeft := NumGet(rect, 0, "int")
            winTop := NumGet(rect, 4, "int")
            winRight := NumGet(rect, 8, "int")
            winBottom := NumGet(rect, 12, "int")

            centerX := winLeft + (winRight - winLeft) // 2
            centerY := winTop + (winBottom - winTop) // 2

            ; Find which monitor contains the window center
            monitorCount := MonitorGetCount()
            loop monitorCount {
                idx := A_Index
                MonitorGetWorkArea(idx, &l, &t, &r, &b)
                if (centerX >= l && centerX <= r && centerY >= t && centerY <= b) {
                    monitorLeft := l
                    monitorTop := t
                    monitorRight := r
                    monitorBottom := b
                    monitorWidth := r - l
                    monitorHeight := b - t
                    break
                }
            }
        }
    }

    ; Create GUI - non-activating so it doesn't steal focus (match Hotstring U aesthetics)
    g_ProjectSelectorGui := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x08000000", "Project Selector")
    ; Use a distinct dark theme background (slightly bluish gray for contrast)
    g_ProjectSelectorGui.BackColor := "BF092F"
    g_ProjectSelectorGui.MarginX := 14
    g_ProjectSelectorGui.MarginY := 10
    fontSize := (monitorHeight < 800) ? 9 : 9
    g_ProjectSelectorGui.SetFont("s" . fontSize . " cCDD6F4", "Segoe UI")

    ; Get categorized projects
    categorized := GetCategorizedProjects()

    resolved := ProjectSelector_ResolveProjectCharMap()
    projectIndexToChar := resolved.projectIndexToChar
    projectIndexToCategory := resolved.projectIndexToCategory

    ; Build display text with category headers (match Hotstring U: single-line "— Category —")
    ; Also build RichEdit lines so we can emphasize mnemonic letters in both [key] and title.
    displayText := ""
    richLines := []

    ; Display each category with header
    for category in g_ProjectCategories {
        ; Find all project indices in this category that have characters assigned
        categoryProjectIndices := []
        for projectIndex, char in projectIndexToChar {
            if (projectIndexToCategory.Has(projectIndex) && projectIndexToCategory[projectIndex] = category) {
                categoryProjectIndices.Push(projectIndex)
            }
        }

        ; Skip if no projects in this category
        if (categoryProjectIndices.Length = 0) {
            continue
        }

        ; Add category header (compact single line)
        displayText .= "— " . category . " —`n"
        richLines.Push({ text: "— " . category . " —" })

        ; Display projects in this category
        for projectIndex in categoryProjectIndices {
            project := g_Projects[projectIndex]

            ; Skip empty placeholders (shouldn't happen, but safety check)
            if (project.name = "" && project.path = "" && project.workPath = "") {
                continue
            }

            ; Get assigned character
            if (projectIndexToChar.Has(projectIndex)) {
                char := projectIndexToChar[projectIndex]
                displayText .= "[" . char . "] " . project.name . "`n"
                richLines.Push({ text: "[" . char . "] " . project.name, key: char })
            }
        }

        displayText .= "`n"  ; Space between categories
        richLines.Push({ text: "" })
    }

    ; Commands section so [c], [3], [L], [K], [ESC] are always visible and grouped
    displayText .= "— Commands —`n"
    displayText .= "[c] Focus Cursor Window`n"
    displayText .= "[3] Activate Preview Windows`n"
    displayText .= "[L] Selection Mode`n"
    displayText .= "[K] Copy from Gemini`n"
    displayText .= "[ESC] Close"

    richLines.Push({ text: "— Commands —" })
    richLines.Push({ text: "[c] Focus Cursor Window", key: "c" })
    richLines.Push({ text: "[3] Activate Preview Windows", key: "3" })
    richLines.Push({ text: "[L] Selection Mode", key: "L" })
    richLines.Push({ text: "[K] Copy from Gemini", key: "K" })
    richLines.Push({ text: "[ESC] Close", key: "E" })

    ; Calculate text dimensions: use 18px per line so Edit control fits all lines without scroll (14px was too small for font)
    baseWidth := 400
    textControlWidth := baseWidth - 20
    lineCount := StrSplit(displayText, "`n").Length
    lineHeight := 18
    textControlHeight := lineCount * lineHeight
    minHeight := 150
    ; No maxHeight cap so all content (projects + Commands) is visible without scroll
    if (textControlHeight < minHeight)
        textControlHeight := minHeight

    ; Title and separator (compact, match U)
    g_ProjectSelectorGui.SetFont("s11 cCDD6F4 Bold", "Segoe UI")
    g_ProjectSelectorGui.Add("Text", "w" . textControlWidth . " Center", "Project Selector")
    g_ProjectSelectorGui.Add("Text", "w" . textControlWidth . " h1 Background45475A")
    g_ProjectSelectorGui.SetFont("s" . fontSize . " cCDD6F4", "Segoe UI")

    ; Content: scrollable RichEdit (so mnemonic letter can be larger)
    MnemonicRich_EnsureDll()
    richCtrl := g_ProjectSelectorGui.Add("Custom",
        "ClassRichEdit50W w" . textControlWidth . " h" . textControlHeight
        . " +0x44 -E0x200 +VScroll -HScroll -Border Background1E1E2E")
    try MnemonicRich_Render(richCtrl, richLines, fontSize, 6, "Segoe UI", "CDD6F4", "1E1E2E")

    ; Footer hint (match U)
    g_ProjectSelectorGui.SetFont("s9 c89B4FA", "Segoe UI")
    g_ProjectSelectorGui.Add("Text", "w" . textControlWidth . " Center", "Press Escape to close.")

    ; Total height: margins + title + separator + gap + content + hint + spacing (no button, match U)
    totalHeight := 10 + 20 + 1 + 4 + textControlHeight + 6 + 18 + 10

    ; Calculate center position for the GUI
    marginX := 20
    marginY := 20
    guiX := monitorLeft + (monitorWidth - baseWidth) // 2
    guiY := monitorTop + (monitorHeight - totalHeight) // 2

    ; Ensure the GUI stays within monitor bounds
    if (guiX < monitorLeft + marginX)
        guiX := monitorLeft + marginX
    if (guiY < monitorTop + marginY)
        guiY := monitorTop + marginY
    if (guiX + baseWidth > monitorLeft + monitorWidth - marginX)
        guiX := monitorLeft + monitorWidth - baseWidth - marginX
    if (guiY + totalHeight > monitorTop + monitorHeight - marginY)
        guiY := monitorTop + monitorHeight - totalHeight - marginY

    ; Show GUI centered on the active window's monitor
    g_ProjectSelectorGui.Show("NA w" . baseWidth . " h" . totalHeight . " x" . guiX . " y" . guiY)

    ; Set active flag and register Escape with Utils so Escape closes the selector (same process); file sentinel for cross-process Escape from Shift keys
    g_ProjectSelectorActive := true
    g_OnEscapePressed := HandleProjectEscape
    try {
        DirCreate(A_ScriptDir "\.cursor")
        FileAppend "", g_WM_SelectorOpenFile
    } catch {
    }
    g_WM_SelectorCloseCheckTimer := SetTimer(WM_CheckSelectorCloseRequest, 120)

    ; Clear handlers array
    g_ProjectHotkeyHandlers := []

    ; Enable hotkeys using the same character mapping as display
    for projectIndex, char in projectIndexToChar {
        handler := CreateProjectHandler(projectIndex)

        ; Store handler for cleanup
        g_ProjectHotkeyHandlers.Push({ char: char, handler: handler })

        ; Enable hotkey (handle special VK codes for comma and period)
        try {
            if (char = ",") {
                Hotkey("vkBC", handler, "On")  ; VK code for comma
            } else if (char = ".") {
                Hotkey("vkBE", handler, "On")  ; VK code for period
            } else {
                Hotkey(char, handler, "On")
                ; Also enable uppercase for lowercase letters
                if (RegExMatch(char, "^[a-z]$")) {
                    Hotkey(StrUpper(char), handler, "On")
                }
            }
        } catch {
            ; Silently ignore if we can't create hotkey
        }
    }

    ; Enable hotkey for Cursor window selection (character "c")
    try {
        cursorWindowHandler := HandleCursorWindowSelectionTrigger
        g_ProjectHotkeyHandlers.Push({ char: "c", handler: cursorWindowHandler })
        Hotkey("c", cursorWindowHandler, "On")
        Hotkey("C", cursorWindowHandler, "On")  ; Also enable uppercase
    } catch {
        ; Silently ignore if we can't create hotkey
    }

    ; Enable hotkey for preview window activation (character "3")
    try {
        previewHandler := HandlePreviewWindowSelection
        g_ProjectHotkeyHandlers.Push({ char: "3", handler: previewHandler })
        Hotkey("3", previewHandler, "On")
    } catch {
        ; Silently ignore if we can't create hotkey
    }

    ; Enable hotkey for Selection Mode (character "L")
    try {
        selectionModeHandler := HandleSelectionModeTrigger
        g_ProjectHotkeyHandlers.Push({ char: "l", handler: selectionModeHandler })
        Hotkey("l", selectionModeHandler, "On")
        Hotkey("L", selectionModeHandler, "On")
    } catch {
    }

    ; Enable hotkey for Copy from Gemini mode (character "K")
    try {
        copyFromGeminiHandler := HandleCopyFromGeminiModeTrigger
        g_ProjectHotkeyHandlers.Push({ char: "k", handler: copyFromGeminiHandler })
        Hotkey("k", copyFromGeminiHandler, "On")
        Hotkey("K", copyFromGeminiHandler, "On")
    } catch {
    }

    SetTimer(ProjectSelector_AutoCloseIfIdle, -3000)
}
; Ctrl+Alt+Win+0: Project Quick Selector (toggle: close if open, open if closed)
^!#0:: {
    global g_ProjectSelectorActive, g_ProjectSelectorGui
    if (g_ProjectSelectorActive && IsObject(g_ProjectSelectorGui)) {
        CleanupProjectSelector()
    } else {
        ShowProjectSelector()
    }
}

; Ctrl+Alt+Win+1: close-M1 (Shift) + project selector (no Shift). * allows extra modifiers (CapsLock, etc.) so the chord
; still matches on picky stacks; use ^!+#g / ^!+#z from the IDE monitor if this still does not fire.
*^!+#1:: CloseWindowOnMonitor(1)
*^!+#SC002:: CloseWindowOnMonitor(1)  ; US QWERTY top-row 1 scan code if character "1" binding differs
^!#1:: {
    global g_ProjectSelectorActive, g_ProjectSelectorGui

    if (!g_ProjectSelectorActive || !IsObject(g_ProjectSelectorGui)) {
        ShowProjectSelector()
    }

    if (!g_ProjectSelectorActive || !IsObject(g_ProjectSelectorGui)) {
        try ShowNotification_WM("Project selector could not be opened.")
        return
    }

    HandleSelectionModeTrigger()
}
