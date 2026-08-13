; =============================================================================
; WindowManagement module: project_selector_02.ahk
; Project selector selection mode and preview handlers
; Extracted verbatim from WindowManagement.ahk; loaded via #include into the
; WindowManagement.ahk process, which remains the entry point / source of truth.
; =============================================================================

; Handler for project selection in Selection Mode
HandleSelectionModeProjectSelection(index) {
    global g_SelectionModeActive, g_Projects

    if (!g_SelectionModeActive) {
        return
    }
    ProjectData_Load()
    if (index < 1 || index > g_Projects.Length) {
        return
    }
    project := g_Projects[index]
    if (project.name = "" && project.path = "" && project.workPath = "") {
        return
    }

    ; Activate the project in Cursor and rely on ActivateCursorProject/FocusCursorAITextField
    ; to handle AI sidebar visibility (only open if hidden, never toggle closed).
    g_SelectionModeActive := false
    projectPath := IS_WORK_ENVIRONMENT ? project.workPath : project.path
    if (IS_WORK_ENVIRONMENT && projectPath = "") {
        projectPath := project.path
    }
    if (projectPath = "" || !DirExist(projectPath)) {
        ShowNotification_WM("Project folder not found: " . projectPath)
        CleanupSelectionMode()
        CleanupProjectSelector()
        return
    }

    ; Best-effort: even if focusing the AI field reports a soft failure,
    ; the Cursor window may still be usable. Suppress noisy failure toast.
    ActivateCursorProject(projectPath)
    CleanupSelectionMode()
    CleanupProjectSelector()
}

; Factory function to create a handler for selection mode project selection
CreateSelectionModeProjectHandler(index) {
    return (*) => HandleSelectionModeProjectSelection(index)
}

; Handler for Selection Mode trigger (L key in project selector)
HandleSelectionModeTrigger(*) {
    global g_ProjectSelectorActive, g_SelectionModeActive

    ; Only process if project selector is active
    if (!g_ProjectSelectorActive) {
        return
    }

    ShowNotification_WM("Entering Selection Mode - Select Project")
    g_SelectionModeActive := true
    ProjectSelector_BindModalHotkeys()
}

; Cleanup selection mode: disable hotkeys and reset state
CleanupSelectionMode() {
    global g_SelectionModeActive, g_SelectionModeHotkeyHandlers

    ; Disable active flag
    g_SelectionModeActive := false

    ; Disable all selection mode character hotkeys
    for handler in g_SelectionModeHotkeyHandlers {
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

    ; Clear handlers array
    g_SelectionModeHotkeyHandlers := []
}

; Cleanup Copy from Gemini mode: disable hotkeys and reset state
CleanupCopyFromGeminiMode() {
    global g_CopyFromGeminiModeActive, g_CopyFromGeminiHotkeyHandlers

    g_CopyFromGeminiModeActive := false
    for handler in g_CopyFromGeminiHotkeyHandlers {
        try {
            char := handler.char
            if (char = ",") {
                Hotkey("vkBC", "Off")
            } else if (char = ".") {
                Hotkey("vkBE", "Off")
            } else {
                Hotkey(char, "Off")
                if (RegExMatch(char, "^[a-z]$")) {
                    Hotkey(StrUpper(char), "Off")
                }
            }
        } catch {
        }
    }
    g_CopyFromGeminiHotkeyHandlers := []
}

; Handler for project selection in Copy from Gemini mode. Delegates to GeminiToCursorBridge module.
HandleCopyFromGeminiProjectSelection(index) {
    global g_CopyFromGeminiModeActive, g_Projects
    ; #region agent log
    _DebugLog_WM("WindowManagement.ahk:HandleCopyFromGeminiProjectSelection", "entry", '{"index":' . index . '}', "H1")
    ; #endregion

    if (!g_CopyFromGeminiModeActive) {
        return
    }
    if (index < 1 || index > g_Projects.Length) {
        return
    }
    project := g_Projects[index]
    if (project.name = "" && project.path = "" && project.workPath = "") {
        return
    }

    g_CopyFromGeminiModeActive := false
    projectPath := IS_WORK_ENVIRONMENT ? project.workPath : project.path
    if (IS_WORK_ENVIRONMENT && projectPath = "") {
        projectPath := project.path
    }
    if (projectPath = "" || !DirExist(projectPath)) {
        ShowNotification_WM("Project folder not found: " . projectPath)
        CleanupCopyFromGeminiMode()
        CleanupProjectSelector()
        return
    }
    ; #region agent log
    pathLast := ""
    try {
        pNorm := RTrim(projectPath, "\")
        parts := StrSplit(pNorm, "\")
        pathLast := parts.Length ? parts[parts.Length] : ""
    } catch {
        pathLast := "?"
    }
    _DebugLog_WM("WindowManagement.ahk:CopyFromGeminiSelection", "calling bridge", '{"index":' . index .
        ',"pathLast":"' . pathLast . '","pathLen":' . StrLen(projectPath) . '}', "WM1")
    ; #endregion

    ; Close selector before bridge so the modal cannot steal focus when we activate the Cursor window.
    CleanupCopyFromGeminiMode()
    CleanupProjectSelector()

    result := CopyFromGeminiToCursor(projectPath, IS_WORK_ENVIRONMENT)
    if (!result.ok) {
        if (result.reason = "no_script")
            ShowNotification_WM("Gemini.ahk not running")
        else if (result.reason = "no_gemini_window")
            ShowNotification_WM("Open Gemini in Chrome first")
        else if (result.reason = "gemini_activate_failed")
            ShowNotification_WM("Could not activate Gemini window")
        else if (result.reason = "send_failed")
            ShowNotification_WM("Could not trigger Gemini copy")
        else if (result.reason = "validation_failed")
            ShowNotification_WM("Copy from Gemini: clipboard not updated")
        else if (result.reason = "cursor_activate_failed")
            ShowNotification_WM("Failed to open project or focus AI field")
        else
            ShowNotification_WM("Copy from Gemini timed out")
        CleanupCopyFromGeminiMode()
        CleanupProjectSelector()
        return
    }
    CleanupCopyFromGeminiMode()
    CleanupProjectSelector()
}

; Factory for Copy from Gemini mode project handler
CreateCopyFromGeminiProjectHandler(index) {
    return (*) => HandleCopyFromGeminiProjectSelection(index)
}

; Handler for Copy from Gemini mode trigger (K key in project selector)
HandleCopyFromGeminiModeTrigger(*) {
    global g_ProjectSelectorActive, g_CopyFromGeminiModeActive, g_Projects, g_ProjectCharSequence
    global g_CopyFromGeminiHotkeyHandlers, g_ProjectHotkeyHandlers

    ; #region agent log
    _DebugLog_WM("WindowManagement.ahk:HandleCopyFromGeminiModeTrigger", "K pressed", '{"selectorActive":' . (
        g_ProjectSelectorActive ? 1 : 0) . '}', "H0")
    ; #endregion
    if (!g_ProjectSelectorActive) {
        return
    }
    ShowNotification_WM("Copy from Gemini - Select Project")
    g_CopyFromGeminiModeActive := true

    ; Disable existing project hotkeys (keep special keys c, 3, l, k, Escape)
    for handler in g_ProjectHotkeyHandlers {
        try {
            char := handler.char
            if (char = "l" || char = "L" || char = "k" || char = "K" || char = "c" || char = "C" || char = "3") {
                continue
            }
            if (char = ",") {
                Hotkey("vkBC", "Off")
            } else if (char = ".") {
                Hotkey("vkBE", "Off")
            } else {
                Hotkey(char, "Off")
                if (RegExMatch(char, "^[a-z]$")) {
                    Hotkey(StrUpper(char), "Off")
                }
            }
        } catch {
        }
    }

    resolved := ProjectSelector_ResolveProjectCharMap()
    projectIndexToChar := resolved.projectIndexToChar

    g_CopyFromGeminiHotkeyHandlers := []
    for projectIndex, char in projectIndexToChar {
        handler := CreateCopyFromGeminiProjectHandler(projectIndex)
        g_CopyFromGeminiHotkeyHandlers.Push({ char: char, handler: handler })
        try {
            if (char = ",") {
                Hotkey("vkBC", handler, "On")
            } else if (char = ".") {
                Hotkey("vkBE", handler, "On")
            } else {
                Hotkey(char, handler, "On")
                if (RegExMatch(char, "^[a-z]$")) {
                    Hotkey(StrUpper(char), handler, "On")
                }
            }
        } catch {
        }
    }
}

; Handler for preview window activation (character "3")
HandlePreviewWindowSelection(*) {
    global g_ProjectSelectorActive, g_Projects

    ; Only process if selector is active
    if (!g_ProjectSelectorActive) {
        return
    }

    ; Cleanup first (closes GUI, disables hotkeys)
    CleanupProjectSelector()

    ; Small delay to ensure cleanup is complete
    Sleep 100

    previewWindows := []
    previewSource := []  ; list of {hwnd, title} from daemon or legacy
    if (WM_UsesAutomationDaemon()) {
        try {
            for w in WMIPC_GetPreviewWindows()
                previewSource.Push({ hwnd: Integer(w["hwnd"]), title: w.Has("title") ? w["title"] : "" })
        } catch {
        }
    }
    if (previewSource.Length = 0) {
        try {
            for hwnd in WinGetList("ahk_exe Cursor.exe") {
                try {
                    previewSource.Push({ hwnd: hwnd, title: WinGetTitle("ahk_id " hwnd) })
                } catch {
                }
            }
        } catch {
        }
    }
    try {
        for item in previewSource {
            hwnd := item.hwnd
            winTitle := item.title
            winTitleLower := StrLower(winTitle)
            if (!InStr(winTitleLower, "preview"))
                continue

            ; Extract workspace name from window title
            ; Format: "Preview filename - WorkspaceName (Workspace) - Cursor"
            ; We want to extract "WorkspaceName"
            workspaceName := ""
            if (RegExMatch(winTitle, "Preview .+? - (.+?) \(Workspace\)", &match)) {
                workspaceName := match[1]
            }

            ; Check if this preview window matches any project
            windowMatched := false
            for project in g_Projects {
                ; Skip empty placeholders
                if (project.name = "" && project.path = "" && project.workPath = "") {
                    continue
                }

                ; Select path based on environment
                projectPath := IS_WORK_ENVIRONMENT ? project.workPath : project.path
                if (IS_WORK_ENVIRONMENT && projectPath = "") {
                    projectPath := project.path
                }

                ; First, try matching by workspace name against project name
                if (workspaceName != "" && project.name != "" && InStr(workspaceName, project.name)) {
                    previewWindows.Push({ hwnd: hwnd, title: winTitle })
                    windowMatched := true
                    break
                }

                ; Also try matching workspace name directly in project path
                if (workspaceName != "" && projectPath != "" && InStr(projectPath, workspaceName)) {
                    previewWindows.Push({ hwnd: hwnd, title: winTitle })
                    windowMatched := true
                    break
                }

                ; Also try matching project name in window title (fallback)
                if (project.name != "" && InStr(winTitle, project.name)) {
                    previewWindows.Push({ hwnd: hwnd, title: winTitle })
                    windowMatched := true
                    break
                }

                if (projectPath = "") {
                    continue
                }

                ; Extract match segments and check if window title matches
                matchSegments := ExtractProjectMatchSegments(projectPath)
                for segment in matchSegments {
                    ; Try exact match first
                    if (InStr(winTitle, segment)) {
                        previewWindows.Push({ hwnd: hwnd, title: winTitle })
                        windowMatched := true
                        break  ; Found a match, no need to check other segments
                    }
                    ; Also try matching segment with "(Workspace)" suffix (for titles like "Trustmate Workspace (Workspace)")
                    if (InStr(winTitle, segment . " (Workspace)")) {
                        previewWindows.Push({ hwnd: hwnd, title: winTitle })
                        windowMatched := true
                        break
                    }
                    ; Also try matching just the last word if segment contains spaces (e.g., "Workspace" from "Trustmate Workspace")
                    if (InStr(segment, " ")) {
                        segmentParts := StrSplit(segment, " ")
                        lastPart := segmentParts[segmentParts.Length]
                        if (InStr(winTitle, lastPart) && InStr(winTitle, segmentParts[1])) {
                            ; Both first and last parts are in title, likely a match
                            previewWindows.Push({ hwnd: hwnd, title: winTitle })
                            windowMatched := true
                            break
                        }
                    }
                }

                ; If we found a match, break from project loop
                if (windowMatched)
                    break
            }
        }
    } catch {
        ShowNotification_WM("No preview windows found.")
        return
    }

    if (previewWindows.Length = 0) {
        try {
            for item in previewSource {
                winTitle := item.title
                winTitleLower := StrLower(winTitle)
                if (!InStr(winTitleLower, "preview"))
                    continue

                ; Extract workspace name
                workspaceName := ""
                if (RegExMatch(winTitle, "Preview .+? - (.+?) \(Workspace\)", &match)) {
                    workspaceName := match[1]
                }

                if (workspaceName != "")
                    previewWindows.Push({ hwnd: item.hwnd, title: winTitle })
            }
        } catch {
        }

        ; If still no preview windows found
        if (previewWindows.Length = 0) {
            ShowNotification_WM("No preview windows found for any project.")
            return
        }
    }

    ; Find the last used preview window
    ; First, check if any of them is currently active
    try {
        activeHwnd := WinGetID("A")
        for window in previewWindows {
            if (window.hwnd = activeHwnd) {
                ; This window is already active, just center mouse
                WMAutomation_SuppressCursorCentering("preview_activate_existing", 1600)
                WinActivate("ahk_id " window.hwnd)
                WM_MaybeCenterMouse(window.hwnd, "preview_activate_existing")
                return
            }
        }
    } catch {
        ; Could not get active window, continue
    }

    ; If no active window matches, get the first window in the list
    ; WinGetList returns windows in z-order (most recently used first)
    if (previewWindows.Length > 0) {
        targetWindow := previewWindows[1]
        try {
            WMAutomation_SuppressCursorCentering("preview_activate_target", 1600)
            WinActivate("ahk_id " targetWindow.hwnd)
            WinWaitActive("ahk_id " targetWindow.hwnd, , 2)
            WM_MaybeCenterMouse(targetWindow.hwnd, "preview_activate_target")
        } catch {
            ShowNotification_WM("Failed to activate preview window.")
        }
    }
}
