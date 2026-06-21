; =============================================================================
; WindowManagement module: project_selector_01.ahk
; Project quick selector GUI and handlers (#!+L)
; Extracted verbatim from WindowManagement.ahk; loaded via #include into the
; WindowManagement.ahk process, which remains the entry point / source of truth.
; =============================================================================

; =============================================================================
; Project Quick Selector
; Hotkey: Win+Alt+Shift+L
; Displays a numbered list of projects and opens the selected folder in Cursor.
; =============================================================================

; Character sequence for assignment: 1 2 3 4 5 q w e r t a s d f g z x c v b 6 7 8 9 0 y u i o p h j k l n m , .
global g_ProjectCharSequence := ["1", "2", "3", "4", "5", "q", "w", "e", "r", "t", "a", "s", "d", "f", "g", "z", "x",
    "c", "v", "b", "6", "7", "8", "9", "0", "y", "u", "i", "o", "p", "h", "j", "k", "l", "n", "m", ",", "."]

; Category display order (General first, Personal second, Work last)
global g_ProjectCategories := ["General", "Personal", "Work"]

ProjectSelector_IsValidChar(char) {
    global g_ProjectCharSequence
    if (char = "" || !IsObject(g_ProjectCharSequence))
        return false
    for c in g_ProjectCharSequence {
        if (c = char)
            return true
    }
    return false
}

ProjectSelector_ResolveProjectCharMap() {
    global g_Projects, g_ProjectCategories, g_ProjectCharSequence
    projectIndexToChar := Map()
    taken := Map()

    projectIndexToCategory := Map()
    loop g_Projects.Length {
        idx := A_Index
        project := g_Projects[idx]
        category := project.HasProp("category") ? project.category : "Personal"
        projectIndexToCategory[idx] := category
    }

    ; Pass 1: explicit hotkeys
    for category in g_ProjectCategories {
        for projectIndex, cat in projectIndexToCategory {
            if (cat != category)
                continue
            project := g_Projects[projectIndex]
            if (project.name = "" && project.path = "" && project.workPath = "")
                continue
            if (project.HasProp("char") && project.char != "") {
                ch := project.char
                if (ch = "3")
                    continue
                if (ProjectSelector_IsValidChar(ch) && !taken.Has(ch)) {
                    projectIndexToChar[projectIndex] := ch
                    taken[ch] := true
                }
            }
        }
    }

    ; Pass 2: sequential assignment for remaining projects
    charIndex := 1
    for category in g_ProjectCategories {
        for projectIndex, cat in projectIndexToCategory {
            if (cat != category)
                continue
            if (projectIndexToChar.Has(projectIndex))
                continue
            project := g_Projects[projectIndex]

            ; Skip empty placeholders but keep charIndex aligned with placeholders
            if (project.name = "" && project.path = "" && project.workPath = "") {
                charIndex++
                continue
            }

            while (charIndex <= g_ProjectCharSequence.Length) {
                ch := g_ProjectCharSequence[charIndex]
                charIndex++
                if (ch = "3")
                    continue
                if (taken.Has(ch))
                    continue
                projectIndexToChar[projectIndex] := ch
                taken[ch] := true
                break
            }
        }
    }

    return { projectIndexToChar: projectIndexToChar, projectIndexToCategory: projectIndexToCategory }
}

; Global project list - add your projects here
; Each project should have: name, path, workPath, and category ("General", "Personal", or "Work")
global g_Projects := [
    ; General category
    { name: "Scripts", path: "C:\Users\eduev\Meu Drive\17 - Projects\scripts", workPath: "C:\Users\fie7ca\Documents\scripts",
        category: "General", char: "s" }, { name: "14-my-Notes", path: "C:\Users\eduev\Meu Drive\17 - Projects\notes",
            workPath: "C:\Users\fie7ca\OneDrive - Bosch Group\14-my-notes",
            category: "General", char: "n" }, { name: "", path: "", workPath: "", category: "General" }, { name: "",
                path: "",
                workPath: "", category: "General" }, { name: "", path: "", workPath: "", category: "General" },
                ; Personal category
                { name: "ZMK Sofle", path: "C:\Users\eduev\Documents\ZMK\zmk-sofle", workPath: "", category: "Personal",
                    char: "z" }, { name: "AI ExperIment",
                        path: "C:\Users\eduev\Documents\Web projects\ai-experiments", workPath: "",
                        category: "Personal", char: "i" }, { name: "my-personal-rePo", path: "C:\Users\eduev\Meu Drive\17 - Projects\my-personal-repo",
                            workPath: "C:\Users\fie7ca\OneDrive - Bosch Group\13 - General workspace\my-personal-repo",
                            category: "Personal", char: "p" }, { name: "",
                                path: "", workPath: "", category: "Personal" }, { name: "", path: "", workPath: "",
                                    category: "Personal" },
                                ; Work category
                                { name: "GS_E&S_CIP Dashboard research and design workspace folder", path: "C:\Users\fie7ca\OneDrive - Bosch Group\13 - General workspace\GS_E&S_CIP Dashboard research and design workspace folder",
                                    workPath: "C:\Users\fie7ca\OneDrive - Bosch Group\13 - General workspace\GS_E&S_CIP Dashboard research and design workspace folder",
                                    category: "Work", char: "d" }, { name: "GS_UX core team_UX and CIP Integration",
                                        path: "",
                                        workPath: "C:\Users\fie7ca\OneDrive - Bosch Group\13 - General workspace\GS_UX core team_UX and CIP Integration",
                                        category: "Work", char: "u" }, { name: "🪂 A vante", path: "", workPath: "C:\Users\fie7ca\OneDrive - Bosch Group\General - GS_BDU_Team\00_UX_GS_Team\AM_Planning\Avante",
                                            category: "Work", char: "v" }, { name: "🪂 Avante – CapacitY", path: "",
                                                workPath: "C:\Users\fie7ca\OneDrive - Bosch Group\General - GS_BDU_Team\00_UX_GS_Team\AM_Planning\Avante\Capacity",
                                                category: "Work", char: "y" }, { name: "E&S Opex CIM Journey Mapping",
                                                    path: "",
                                                    workPath: "C:\Users\fie7ca\OneDrive - Bosch Group\13 - General workspace\opex-cim-journey-mapping",
                                                    category: "Work", char: "o" }, { name: "boiler-plate", path: "",
                                                        workPath: "C:\Users\fie7ca\OneDrive - Bosch Group\13 - General workspace\boiler-plate",
                                                        category: "Work", char: "0" }, { name: "astra", path: "",
                                                            workPath: "C:\Users\fie7ca\OneDrive - Bosch Group\SO UX - LA (Internal) - Projeto Astra",
                                                            category: "Work", char: "a" }, { name: "Piloto PT B2B",
                                                                path: "C:\Users\fie7ca\OneDrive - Bosch Group\SO UX - LA (Internal) - Data Insights SO - Piloto PT B2B",
                                                                workPath: "C:\Users\fie7ca\OneDrive - Bosch Group\SO UX - LA (Internal) - Data Insights SO - Piloto PT B2B",
                                                                category: "Work", char: "b" }, { name: "Python ScripTs",
                                                                    path: "C:\Users\eduev\Meu Drive\17 - Projects\My-Python-Scripts",
                                                                    workPath: "C:\Users\fie7ca\OneDrive - Bosch Group\17 - Python Scripts",
                                                                    category: "Work", char: "t" }
]
; TODO: Fill in workPath for each project above when configuring work environment
; Global variables for project selector
global g_ProjectSelectorGui := false
global g_ProjectSelectorActive := false
global g_ProjectHotkeyHandlers := []  ; Store hotkey handlers for cleanup

; Global variables for Cursor window selector (used within project selector)
global g_CursorWindowMap := Map()  ; Maps character to window HWND
global g_CursorWindowHotkeyHandlers := []  ; Store hotkey handlers for cleanup
global g_CursorWindowSelectorGui := false

; Global variable for Selection Mode
global g_SelectionModeActive := false
global g_SelectionModeHotkeyHandlers := []  ; Store hotkey handlers for selection mode cleanup

; Global variables for Copy from Gemini mode (K in project selector)
global g_CopyFromGeminiModeActive := false
global g_CopyFromGeminiHotkeyHandlers := []

; File-based IPC so Shift keys (or other process) can request project selector close on Escape
global g_WM_SelectorOpenFile := A_ScriptDir "\.cursor\wm_selector_open"
global g_WM_SelectorCloseRequestFile := A_ScriptDir "\.cursor\wm_selector_close_request"
global g_WM_SelectorCloseCheckTimer := ""

; Cross-process IPC for Hotstring Selector (Utils.ahk)
global g_HS_SelectorOpenFile_WM := A_ScriptDir "\.cursor\hs_selector_open"
global g_HS_SelectorCloseRequestFile_WM := A_ScriptDir "\.cursor\hs_selector_close_request"

WM_CheckSelectorCloseRequest() {
    global g_ProjectSelectorActive, g_WM_SelectorCloseRequestFile
    if (!g_ProjectSelectorActive)
        return
    if (FileExist(g_WM_SelectorCloseRequestFile)) {
        try FileDelete(g_WM_SelectorCloseRequestFile)
        catch {
        }
        CleanupProjectSelector()
    }
}

; Activate a Cursor project by path: find or launch window, then focus the AI text field. Returns true on success.
; Ensures the target project window is explicitly activated before focus/paste, regardless of current active window.
ActivateCursorProject(projectPath) {
    ; #region agent log
    _DebugLog_WM("WindowManagement.ahk:ActivateCursorProject", "entry", '{"pathLen":' . StrLen(projectPath) .
    ',"dirExists":' . (DirExist(projectPath) ? 1 : 0) . '}', "H3")
    ; #endregion
    if (projectPath = "" || !DirExist(projectPath)) {
        return false
    }
    targetHwnd := FindAndActivateCursorWindow(projectPath)
    ; #region agent log
    _DebugLog_WM("WindowManagement.ahk:ActivateCursorProject", "after FindAndActivate", '{"targetHwnd":' . targetHwnd .
        '}', "H3")
    ; #endregion
    if (!targetHwnd) {
        cursorPath := IS_WORK_ENVIRONMENT ?
            "C:\Users\fie7ca\AppData\Local\Programs\cursor\Cursor.exe" :
                "C:\Users\eduev\AppData\Local\Programs\cursor\Cursor.exe"
        try {
            Run cursorPath . ' "' . projectPath . '"'
        } catch {
            return false
        }
        ; Wait for the new window to appear and match our project
        loop 30 {
            Sleep 200
            targetHwnd := GetCursorHwndForProject(projectPath)
            if (targetHwnd)
                break
        }
        if (!targetHwnd) {
            return false
        }
    }
    ; Explicitly activate the target window so paste goes to the correct project (works regardless of current active window).
    try {
        WinActivate("ahk_id " targetHwnd)
        WinWaitActive("ahk_id " targetHwnd, , 3)
    } catch {
        ShowNotification_WM("Error: Target window not found.")
        return false
    }
    Sleep 300
    focusOk := FocusCursorAITextField(targetHwnd)
    ; #region agent log
    _DebugLog_WM("WindowManagement.ahk:ActivateCursorProject", "after FocusCursorAITextField", '{"focusOk":' . (focusOk ?
        1 : 0) . '}', "H4")
    ; #endregion
    if (focusOk) {
        try {
            ScriptSoundPlay(A_ScriptDir . "\sounds\into-cursor-textfield.wav")
        } catch {
        }
        ; #region agent log
        _DebugLog_WM("WindowManagement.ahk:ActivateCursorProject", "return true", "{}", "H3")
        ; #endregion
        return true
    }
    ; #region agent log
    _DebugLog_WM("WindowManagement.ahk:ActivateCursorProject", "return false (focus failed)", "{}", "H4")
    ; #endregion
    return false
}

; Get categorized projects for display
GetCategorizedProjects() {
    global g_Projects
    categorized := Map()
    categorized["General"] := []
    categorized["Personal"] := []
    categorized["Work"] := []

    if (!IsSet(g_Projects) || g_Projects.Length = 0) {
        return categorized
    }

    for project in g_Projects {
        category := project.HasProp("category") ? project.category : "Personal"
        if (category = "General" || category = "Personal" || category = "Work") {
            categorized[category].Push(project)
        }
    }

    return categorized
}
; One-shot: close project selector if still open (no project/command chosen in time)
ProjectSelector_AutoCloseIfIdle() {
    global g_ProjectSelectorActive
    if (g_ProjectSelectorActive)
        CleanupProjectSelector()
}

; Cleanup project selector: destroy GUI, disable hotkeys, reset state
CleanupProjectSelector() {
    global g_ProjectSelectorActive, g_ProjectSelectorGui, g_ProjectHotkeyHandlers, g_SelectionModeActive,
        g_CopyFromGeminiModeActive, g_WM_SelectorOpenFile, g_WM_SelectorCloseRequestFile, g_WM_SelectorCloseCheckTimer

    SetTimer(ProjectSelector_AutoCloseIfIdle, 0)
    g_ProjectSelectorActive := false
    SetTimer(WM_CheckSelectorCloseRequest, 0)
    g_WM_SelectorCloseCheckTimer := ""
    try FileDelete(g_WM_SelectorOpenFile)
    catch {
    }
    try FileDelete(g_WM_SelectorCloseRequestFile)
    catch {
    }
    if (g_SelectionModeActive) {
        CleanupSelectionMode()
    }
    if (g_CopyFromGeminiModeActive) {
        CleanupCopyFromGeminiMode()
    }

    ; Disable all character hotkeys
    for handler in g_ProjectHotkeyHandlers {
        try {
            char := handler.char
            ; Handle special VK codes for comma and period
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

    ; Unregister Escape callback so Utils forwards Escape again
    g_OnEscapePressed := ""

    ; Clear handlers array
    g_ProjectHotkeyHandlers := []

    ; Close and destroy GUI
    if (IsObject(g_ProjectSelectorGui)) {
        try {
            g_ProjectSelectorGui.Destroy()
        } catch {
            ; Ignore
        }
        g_ProjectSelectorGui := false
    }
}

; Return the hwnd of a Cursor window whose title matches the project path, or 0. Does not activate.
GetCursorHwndForProject(projectPath) {
    if (WM_UsesAutomationDaemon()) {
        try {
            r := WMIPC_ResolveProjectWindow(projectPath)
            if (r.Has("hwnd") && Integer(r["hwnd"]) != 0)
                return Integer(r["hwnd"])
        } catch {
        }
    }
    matchSegments := ExtractProjectMatchSegments(projectPath)
    try {
        for hwnd in WinGetList("ahk_exe Cursor.exe") {
            try {
                winTitle := WinGetTitle("ahk_id " hwnd)
                if (InStr(StrLower(winTitle), "preview"))
                    continue
                for segment in matchSegments {
                    if (InStr(winTitle, segment))
                        return hwnd
                }
            } catch {
                continue
            }
        }
    } catch {
    }
    return 0
}

; Return the hwnd of a VS Code window whose title matches the project path, or 0. Does not activate.
GetVSCodeHwndForProject(projectPath) {
    matchSegments := ExtractProjectMatchSegments(projectPath)
    try {
        for hwnd in WinGetList("ahk_exe Code.exe") {
            try {
                winTitle := WinGetTitle("ahk_id " hwnd)
                if (InStr(StrLower(winTitle), "preview"))
                    continue
                for segment in matchSegments {
                    if (InStr(winTitle, segment))
                        return hwnd
                }
            } catch {
                continue
            }
        }
    } catch {
    }
    return 0
}

; Find and activate the last used VS Code window for a project path.
; Returns the activated window's hwnd, or 0 if not found / activation failed.
FindAndActivateVSCodeWindow(projectPath) {
    matchSegments := ExtractProjectMatchSegments(projectPath)
    codeWindows := []

    try {
        for hwnd in WinGetList("ahk_exe Code.exe") {
            try {
                winTitle := WinGetTitle("ahk_id " hwnd)
                winTitleLower := StrLower(winTitle)
                if (InStr(winTitleLower, "preview"))
                    continue
                for segment in matchSegments {
                    if (InStr(winTitle, segment)) {
                        codeWindows.Push({ hwnd: hwnd, title: winTitle })
                        break
                    }
                }
            } catch {
                continue
            }
        }
    } catch {
    }

    if (codeWindows.Length = 0)
        return 0

    try {
        activeHwnd := WinGetID("A")
        for window in codeWindows {
            if (window.hwnd = activeHwnd) {
                WMAutomation_SuppressCursorCentering("vscode_activate_existing", 1600)
                WinActivate("ahk_id " window.hwnd)
                WM_MaybeCenterMouse(window.hwnd, "vscode_activate_existing")
                return window.hwnd
            }
        }
    } catch {
    }

    targetWindow := codeWindows[1]
    try {
        WMAutomation_SuppressCursorCentering("vscode_activate_target", 1600)
        WinActivate("ahk_id " targetWindow.hwnd)
        WinWaitActive("ahk_id " targetWindow.hwnd, , 2)
        WM_MaybeCenterMouse(targetWindow.hwnd, "vscode_activate_target")
        return targetWindow.hwnd
    } catch {
        return 0
    }
}

; Asynchronously activate the VS Code window that opens after launching a folder.
; This avoids blocking the hotkey handler on first-open.
global g_VSCodeLaunchActivate := { active: false, projectPath: "", startedAt: 0, timeoutMs: 0 }

VSCode_ScheduleActivateAfterLaunch(projectPath, timeoutMs := 8000) {
    global g_VSCodeLaunchActivate
    g_VSCodeLaunchActivate.active := true
    g_VSCodeLaunchActivate.projectPath := projectPath
    g_VSCodeLaunchActivate.startedAt := A_TickCount
    g_VSCodeLaunchActivate.timeoutMs := timeoutMs
    SetTimer(VSCode_TryActivateAfterLaunch, 150)
}

VSCode_TryActivateAfterLaunch() {
    global g_VSCodeLaunchActivate
    if (!g_VSCodeLaunchActivate.active) {
        SetTimer(VSCode_TryActivateAfterLaunch, 0)
        return
    }
    if ((A_TickCount - g_VSCodeLaunchActivate.startedAt) > g_VSCodeLaunchActivate.timeoutMs) {
        g_VSCodeLaunchActivate.active := false
        SetTimer(VSCode_TryActivateAfterLaunch, 0)
        return
    }
    try {
        hwnd := GetVSCodeHwndForProject(g_VSCodeLaunchActivate.projectPath)
        if (hwnd && Integer(hwnd) != 0) {
            WMAutomation_SuppressCursorCentering("vscode_activate_after_launch", 1600)
            WinActivate("ahk_id " hwnd)
            WM_MaybeCenterMouse(hwnd, "vscode_activate_after_launch")
            g_VSCodeLaunchActivate.active := false
            SetTimer(VSCode_TryActivateAfterLaunch, 0)
            return
        }
    } catch {
    }
}

; Find and activate the last used Cursor window for a project path.
; Returns the activated window's hwnd, or 0 if not found / activation failed.
FindAndActivateCursorWindow(projectPath) {
    matchSegments := ExtractProjectMatchSegments(projectPath)
    cursorWindows := []

    if (WM_UsesAutomationDaemon()) {
        try {
            for w in WMIPC_GetCursorWindows() {
                title := w.Has("title") ? w["title"] : ""
                if (!title || InStr(StrLower(title), "preview"))
                    continue
                for segment in matchSegments {
                    if (InStr(title, segment)) {
                        cursorWindows.Push({ hwnd: Integer(w["hwnd"]), title: title })
                        break
                    }
                }
            }
        } catch {
        }
    }

    if (cursorWindows.Length = 0) {
        try {
            for hwnd in WinGetList("ahk_exe Cursor.exe") {
                try {
                    winTitle := WinGetTitle("ahk_id " hwnd)
                    winTitleLower := StrLower(winTitle)
                    if (InStr(winTitleLower, "preview"))
                        continue
                    for segment in matchSegments {
                        if (InStr(winTitle, segment)) {
                            cursorWindows.Push({ hwnd: hwnd, title: winTitle })
                            break
                        }
                    }
                } catch {
                    continue
                }
            }
        } catch {
        }
    }

    if (cursorWindows.Length = 0)
        return 0

    try {
        activeHwnd := WinGetID("A")
        for window in cursorWindows {
            if (window.hwnd = activeHwnd) {
                WMAutomation_SuppressCursorCentering("cursor_activate_existing", 1600)
                WinActivate("ahk_id " window.hwnd)
                WM_MaybeCenterMouse(window.hwnd, "cursor_activate_existing")
                return window.hwnd
            }
        }
    } catch {
    }

    targetWindow := cursorWindows[1]
    try {
        WMAutomation_SuppressCursorCentering("cursor_activate_target", 1600)
        WinActivate("ahk_id " targetWindow.hwnd)
        WinWaitActive("ahk_id " targetWindow.hwnd, , 2)
        WM_MaybeCenterMouse(targetWindow.hwnd, "cursor_activate_target")
        return targetWindow.hwnd
    } catch {
        return 0
    }
}

; Handle project selection - activates existing Cursor window or launches new one
HandleProjectSelection(index) {
    global g_ProjectSelectorActive, g_Projects
    global IS_WORK_ENVIRONMENT, VS_CODE_EXE_WORK

    ; Only process if selector is active
    if (!g_ProjectSelectorActive) {
        return
    }

    ; Validate index
    if (index < 1 || index > g_Projects.Length) {
        return
    }

    ; Get project
    project := g_Projects[index]

    ; Skip empty placeholders (no name or path)
    if (project.name = "" && project.path = "" && project.workPath = "") {
        return
    }

    ; Cleanup first (closes GUI, disables hotkeys)
    CleanupProjectSelector()

    ; Select path based on environment
    projectPath := IS_WORK_ENVIRONMENT ? project.workPath : project.path

    ; If work environment but no workPath set, fall back to personal path
    if (IS_WORK_ENVIRONMENT && projectPath = "") {
        projectPath := project.path
    }

    ; Validate project path exists
    if (projectPath = "" || !DirExist(projectPath)) {
        ShowNotification_WM("Project folder not found: " . projectPath)
        return
    }

    if (IS_WORK_ENVIRONMENT) {
        ; Work: prefer VS Code
        if (FindAndActivateVSCodeWindow(projectPath)) {
            return
        }
        if (!IsSet(VS_CODE_EXE_WORK) || VS_CODE_EXE_WORK = "" || !FileExist(VS_CODE_EXE_WORK)) {
            ShowNotification_WM("VS Code not found: " . (IsSet(VS_CODE_EXE_WORK) ? VS_CODE_EXE_WORK : ""))
            return
        }
        try {
            Run '"' . VS_CODE_EXE_WORK . '" "' . projectPath . '"'
            VSCode_ScheduleActivateAfterLaunch(projectPath, 9000)
        } catch Error as e {
            ShowNotification_WM("Failed to launch VS Code: " . e.Message)
        }
        return
    }

    ; Personal: keep Cursor behavior
    if (FindAndActivateCursorWindow(projectPath)) {
        return
    }
    cursorPath := "C:\Users\eduev\AppData\Local\Programs\cursor\Cursor.exe"
    try {
        Run cursorPath . ' "' . projectPath . '"'
    } catch Error as e {
        ShowNotification_WM("Failed to launch Cursor: " . e.Message)
    }
}
; Factory function to create a handler that properly captures the index
CreateProjectHandler(index) {
    return (*) => HandleProjectSelection(index)
}

; Handler for Escape key in project selector
HandleProjectEscape(*) {
    global g_ProjectSelectorActive
    if (g_ProjectSelectorActive) {
        CleanupProjectSelector()
    }
}
