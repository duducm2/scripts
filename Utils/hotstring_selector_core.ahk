; =============================================================================
; Utils module: hotstring_selector_core.ahk
; Hotstring selector system core (#!+U Utility Shortcuts)
; =============================================================================

; Global state
global g_HotstringSelectorGui := false
global g_HotstringSelectorLv := false
global g_HotstringSelectorHint := false
global g_HotstringSelectorFilterLabel := false
global g_HotstringSelectorFilterCtrl := false
global g_HotstringSelectorFilterEnterBtn := false
global g_HotstringSelectorBtnAdd := false
global g_HotstringSelectorBtnEdit := false
global g_HotstringSelectorBtnDelete := false
global g_HotstringSelectorBtnClose := false
global g_HotstringSelectorGuiReady := false
global g_UtilitySelectorFilterQuery := ""
global g_UtilitySelectorFilterTyping := false
global g_UtilitySelectorSuppressFilterKillFocus := false
global g_HotstringSelectorActive := false
global g_HotstringHotkeyHandlers := []
global g_HotstringPromptCharMap := Map()
global g_HotstringGeminiArmed := false
global g_HotstringGeminiAutoSubmit := true
global g_HotstringGeminiSubmitTimer := false
global g_HotstringGeminiRestoreHwnd := 0
global g_UtilitySelectorRestoreHwnd := 0
global g_UtilitySelectorHotkeysBound := false
global g_UtilitySelectorNoActivate := false
global g_UtilitySelectorRows := []

global g_UtilitySelectorMode := "top"
global g_UtilitySelectorCategory := ""

global g_UtilityTopCategories := ["Prompts", "Projects", "Macros", "Hotstrings", "Sequences", "Finance",
    "Memory Palace",
    "Push"]
global g_UtilityTopCategoryById := Map("r", "Prompts", "p", "Projects", "m", "Macros", "h", "Hotstrings", "s",
    "Sequences", "f", "Finance", "n", "Memory Palace", "g", "Push")

global g_HotstringCharSequence := ["1", "2", "3", "4", "5", "q", "w", "e", "r", "t", "a", "s", "d", "f", "g", "z", "x",
    "c", "v", "b", "6", "7", "8", "9", "0", "y", "u", "i", "o", "p", "h", "j", "k", "l", "n", "m", ",", "."]

global g_ReservedEmptyChar := ""

; Assign macro characters (explicit first, then sequential). Used by the Macros category view.
BuildMacroCharMap() {
    global g_Macros, g_MacroCharMap, g_HotstringCharSequence, g_ReservedEmptyChar
    g_MacroCharMap := Map()
    if (!IsSet(g_Macros) || g_Macros.Length = 0)
        return g_MacroCharMap

    for macroEntry in g_Macros {
        if (macroEntry.HasProp("char") && macroEntry.char != "" && (g_ReservedEmptyChar = "" || macroEntry.char !=
            g_ReservedEmptyChar)) {
            charIndexInSequence := 0
            for idx, seqChar in g_HotstringCharSequence {
                if (seqChar = macroEntry.char) {
                    charIndexInSequence := idx
                    break
                }
            }
            if (charIndexInSequence > 0 && !g_MacroCharMap.Has(macroEntry.char))
                g_MacroCharMap[macroEntry.char] := macroEntry.func
        }
    }

    charIndex := 1
    for macroEntry in g_Macros {
        alreadyAssigned := false
        for assignedChar, assignedFunc in g_MacroCharMap {
            if (assignedFunc = macroEntry.func) {
                alreadyAssigned := true
                break
            }
        }
        if (alreadyAssigned)
            continue
        while (charIndex <= g_HotstringCharSequence.Length) {
            char := g_HotstringCharSequence[charIndex]
            if (g_ReservedEmptyChar != "" && char = g_ReservedEmptyChar) {
                charIndex++
                continue
            }
            if (!g_MacroCharMap.Has(char)) {
                g_MacroCharMap[char] := macroEntry.func
                charIndex++
                break
            }
            charIndex++
        }
    }
    return g_MacroCharMap
}

UtilitySelector_RebuildPromptCharMap() {
    global g_HotstringPromptCharMap
    g_HotstringPromptCharMap := Map()
    for prompt in PromptData_Sorted() {
        if (prompt.char != "")
            g_HotstringPromptCharMap[prompt.char] := true
    }
}

UtilitySelector_MacroRows() {
    global g_Macros, g_MacroCharMap
    if (!IsObject(g_MacroCharMap) || g_MacroCharMap.Count = 0)
        BuildMacroCharMap()
    rows := []
    if (!IsSet(g_Macros))
        return rows
    funcToChar := Map()
    for c, fn in g_MacroCharMap
        funcToChar[fn] := c
    for macro in g_Macros {
        ch := funcToChar.Has(macro.func) ? funcToChar[macro.func] : ""
        rows.Push({ char: ch, title: macro.title, func: macro.func })
    }
    return rows
}

UtilitySelector_ProjectRows() {
    rows := []
    for project in ProjectData_Load(false, true) {
        if (project.name = "" && project.path = "" && project.workPath = "")
            continue
        rows.Push(project)
    }
    return rows
}

UtilitySelector_ProjectCountCached() {
    global g_Projects, g_ProjectDataCacheReady
    if (!g_ProjectDataCacheReady)
        return UtilitySelector_ProjectRows().Length
    n := 0
    for project in g_Projects {
        if (project.name = "" && project.path = "" && project.workPath = "")
            continue
        n++
    }
    return n
}

UtilitySelector_MacroCountCached() {
    global g_Macros
    if (!IsSet(g_Macros) || !IsObject(g_Macros))
        return 0
    return g_Macros.Length
}

GetPreviewText(text, maxLength := 60) {
    preview := RegExReplace(text, "`r?`n", " ")
    preview := RegExReplace(preview, "\s+", " ")
    preview := Trim(preview)
    if (StrLen(preview) <= maxLength)
        return preview
    return SubStr(preview, 1, maxLength) . "..."
}

FindAndActivatePowerBIFile(filePath) {
    if (!FileExist(filePath))
        return false
    SplitPath(filePath, , , , &fileNameNoExt)
    fileNameNoExt := Trim(fileNameNoExt)
    fileNameLower := StrLower(fileNameNoExt)
    try {
        for hwnd in WinGetList("ahk_exe PBIDesktop.exe") {
            try {
                winTitleLower := StrLower(Trim(WinGetTitle("ahk_id " hwnd)))
                if (winTitleLower = fileNameLower || InStr(winTitleLower, fileNameLower) = 1) {
                    WinActivate("ahk_id " hwnd)
                    WinWaitActive("ahk_id " hwnd, , 2)
                    return true
                }
            } catch {
                continue
            }
        }
    } catch {
    }
    try {
        Run(filePath)
        return true
    } catch {
        return false
    }
}

FindAndActivateMiroWindow(url, titleKeywords) {
    keywordsLower := StrLower(titleKeywords)
    try {
        for hwnd in WinGetList("ahk_exe chrome.exe") {
            try {
                winTitleLower := StrLower(Trim(WinGetTitle("ahk_id " hwnd)))
                if (InStr(winTitleLower, "miro") && InStr(winTitleLower, keywordsLower)) {
                    try {
                        if (WinGetMinMax("ahk_id " hwnd) = -1)
                            WinRestore("ahk_id " hwnd)
                        WinActivate("ahk_id " hwnd)
                        WinWaitActive("ahk_id " hwnd, , 2)
                        WinSetAlwaysOnTop("On", "ahk_id " hwnd)
                        Sleep 50
                        WinSetAlwaysOnTop("Off", "ahk_id " hwnd)
                    } catch {
                    }
                    return true
                }
            } catch {
                continue
            }
        }
    } catch {
    }
    try {
        Run("chrome.exe --new-window " . url)
        WinWait("ahk_exe chrome.exe", , 10)
        Sleep(1000)
        loop 10 {
            for hwnd in WinGetList("ahk_exe chrome.exe") {
                try {
                    winTitleLower := StrLower(Trim(WinGetTitle("ahk_id " hwnd)))
                    if (InStr(winTitleLower, "miro") && InStr(winTitleLower, keywordsLower)) {
                        if (WinGetMinMax("ahk_id " hwnd) = -1)
                            WinRestore("ahk_id " hwnd)
                        WinActivate("ahk_id " hwnd)
                        WinWaitActive("ahk_id " hwnd, , 2)
                        WinSetAlwaysOnTop("On", "ahk_id " hwnd)
                        Sleep 50
                        WinSetAlwaysOnTop("Off", "ahk_id " hwnd)
                        return true
                    }
                } catch {
                    continue
                }
            }
            Sleep(500)
        }
        try {
            chromeWindows := WinGetList("ahk_exe chrome.exe")
            if (chromeWindows.Length > 0) {
                hwnd := chromeWindows[1]
                if (WinGetMinMax("ahk_id " hwnd) = -1)
                    WinRestore("ahk_id " hwnd)
                WinActivate("ahk_id " hwnd)
                WinWaitActive("ahk_id " hwnd, , 2)
                WinSetAlwaysOnTop("On", "ahk_id " hwnd)
                Sleep 50
                WinSetAlwaysOnTop("Off", "ahk_id " hwnd)
                return true
            }
        } catch {
        }
        return true
    } catch {
        return false
    }
}
