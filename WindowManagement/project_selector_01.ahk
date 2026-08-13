; =============================================================================
; WindowManagement module: project_selector_01.ahk
; Project quick selector GUI and handlers (Ctrl+Alt+Win+0)
; Extracted verbatim from WindowManagement.ahk; loaded via #include into the
; WindowManagement.ahk process, which remains the entry point / source of truth.
; =============================================================================

; =============================================================================
; Project Quick Selector
; Hotkey: Ctrl+Alt+Win+0 (see WindowManagement\cursor_window_select.ahk)
; ListView of projects (char, name, paths) with add/edit/delete; opens the selected folder.
; =============================================================================

; Project list lives in assets/data/projects.ini (loaded by Utils\project_data_cursor.ahk).

ProjectSelector_IsValidChar(char) {
    return ProjectData_IsValidChar(char)
}

; Map each project to its stored unique char. No category grouping or auto-fill.
ProjectSelector_ResolveProjectCharMap() {
    global g_Projects
    ProjectData_Load()
    projectIndexToChar := Map()
    taken := Map()
    loop g_Projects.Length {
        idx := A_Index
        project := g_Projects[idx]
        if (project.name = "" && project.path = "" && project.workPath = "")
            continue
        if (!project.HasProp("char") || project.char = "")
            continue
        ch := project.char
        if (ProjectSelector_IsValidChar(ch) && !taken.Has(ch)) {
            projectIndexToChar[idx] := ch
            taken[ch] := true
        }
    }
    return { projectIndexToChar: projectIndexToChar }
}

; Global variables for project selector
global g_ProjectSelectorGui := false
global g_ProjectSelectorLv := false
global g_ProjectSelectorActive := false
global g_ProjectSelectorHotkeysBound := false
global g_ProjectHotkeyHandlers := []  ; Store hotkey handlers for cleanup
global g_ProjectPathPickGui := false
global g_ProjectPathPickPersonalEdit := false
global g_ProjectPathPickWorkEdit := false
global g_ProjectPathPickResult := ""

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

WM_CheckSelectorCloseRequest() {
    global g_ProjectSelectorActive, g_WM_SelectorCloseRequestFile, g_WM_SelectorOpenFile
    if (!g_ProjectSelectorActive) {
        try FileDelete(g_WM_SelectorOpenFile)
        catch {
        }
        try FileDelete(g_WM_SelectorCloseRequestFile)
        catch {
        }
        return
    }
    if (FileExist(g_WM_SelectorCloseRequestFile)) {
        try FileDelete(g_WM_SelectorCloseRequestFile)
        catch {
        }
        CleanupProjectSelector()
    }
}

; Remove sentinel files left by a crash or force-kill so global Escape is not permanently swallowed.
WM_CleanupStaleEscapeSentinels() {
    global g_ProjectSelectorActive, g_WM_SelectorOpenFile, g_WM_SelectorCloseRequestFile
    global g_WM_MinimizedListActive, g_WM_MinimizedListOpenFile, g_WM_MinimizedListCloseRequestFile
    if (!g_ProjectSelectorActive) {
        try FileDelete(g_WM_SelectorOpenFile)
        catch {
        }
        try FileDelete(g_WM_SelectorCloseRequestFile)
        catch {
        }
    }
    if (!g_WM_MinimizedListActive) {
        try FileDelete(g_WM_MinimizedListOpenFile)
        catch {
        }
        try FileDelete(g_WM_MinimizedListCloseRequestFile)
        catch {
        }
    }
}

SetTimer(WM_CleanupStaleEscapeSentinels, -1)

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
            ScriptSoundPlay(A_ScriptDir . "\assets\sounds\into-cursor-textfield.wav")
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

ProjectSelector_HotkeyName(char) {
    if (char = ",")
        return "vkBC"
    if (char = ".")
        return "vkBE"
    return char
}

ProjectSelector_UnbindModalHotkeys() {
    global g_ProjectSelectorGui, g_ProjectHotkeyHandlers, g_ProjectSelectorHotkeysBound
    hwnd := 0
    try {
        if (IsObject(g_ProjectSelectorGui))
            hwnd := g_ProjectSelectorGui.Hwnd
    } catch {
        hwnd := 0
    }
    if (hwnd) {
        try HotIfWinActive("ahk_id " hwnd)
        catch {
        }
    }
    for handler in g_ProjectHotkeyHandlers {
        try Hotkey(handler.key, "Off")
        catch {
        }
    }
    if (hwnd) {
        try HotIf()
        catch {
        }
    }
    g_ProjectHotkeyHandlers := []
    g_ProjectSelectorHotkeysBound := false
}

ProjectSelector_BindOneChar(char, handler) {
    global g_ProjectHotkeyHandlers
    key := ProjectSelector_HotkeyName(char)
    try {
        Hotkey(key, handler, "On")
        g_ProjectHotkeyHandlers.Push({ char: char, key: key, handler: handler })
    } catch {
    }
    if (RegExMatch(char, "^[a-z]$")) {
        upperKey := StrUpper(char)
        try {
            Hotkey(upperKey, handler, "On")
            g_ProjectHotkeyHandlers.Push({ char: char, key: upperKey, handler: handler })
        } catch {
        }
    }
}

ProjectSelector_BindModalHotkeys() {
    global g_ProjectSelectorGui, g_ProjectSelectorHotkeysBound, g_SelectionModeActive, g_ProjectHotkeyHandlers
    ProjectSelector_UnbindModalHotkeys()
    hwnd := 0
    try {
        if (IsObject(g_ProjectSelectorGui))
            hwnd := g_ProjectSelectorGui.Hwnd
    } catch {
        hwnd := 0
    }
    if (!hwnd)
        return
    try HotIfWinActive("ahk_id " hwnd)
    catch {
        return
    }

    resolved := ProjectSelector_ResolveProjectCharMap()
    for projectIndex, char in resolved.projectIndexToChar {
        handler := g_SelectionModeActive
            ? CreateSelectionModeProjectHandler(projectIndex)
                : CreateProjectHandler(projectIndex)
        ProjectSelector_BindOneChar(char, handler)
    }

    try {
        Hotkey("Insert", ProjectSelector_OnAdd, "On")
        g_ProjectHotkeyHandlers.Push({ char: "Insert", key: "Insert", handler: ProjectSelector_OnAdd })
    } catch {
    }
    try {
        Hotkey("F2", ProjectSelector_OnEdit, "On")
        g_ProjectHotkeyHandlers.Push({ char: "F2", key: "F2", handler: ProjectSelector_OnEdit })
    } catch {
    }
    try {
        Hotkey("Delete", ProjectSelector_OnDelete, "On")
        g_ProjectHotkeyHandlers.Push({ char: "Delete", key: "Delete", handler: ProjectSelector_OnDelete })
    } catch {
    }
    try {
        Hotkey("Enter", ProjectSelector_OnListActivate, "On")
        g_ProjectHotkeyHandlers.Push({ char: "Enter", key: "Enter", handler: ProjectSelector_OnListActivate })
    } catch {
    }
    try {
        Hotkey("Escape", HandleProjectEscape, "On")
        g_ProjectHotkeyHandlers.Push({ char: "Escape", key: "Escape", handler: HandleProjectEscape })
    } catch {
    }

    try HotIf()
    catch {
    }
    g_ProjectSelectorHotkeysBound := true
}

ProjectSelector_PopulateLv() {
    global g_ProjectSelectorLv, g_Projects
    if (!IsObject(g_ProjectSelectorLv))
        return
    ProjectData_Load()
    g_ProjectSelectorLv.Delete()
    for project in g_Projects {
        ch := project.HasProp("char") ? project.char : ""
        g_ProjectSelectorLv.Add("", ch, project.name, project.path, project.workPath)
    }
    try g_ProjectSelectorLv.ModifyCol(1, 50)
    try g_ProjectSelectorLv.ModifyCol(2, 220)
    try g_ProjectSelectorLv.ModifyCol(3, 280)
    try g_ProjectSelectorLv.ModifyCol(4, 280)
}

ProjectSelector_SelectedIndex() {
    global g_ProjectSelectorLv
    if (!IsObject(g_ProjectSelectorLv))
        return 0
    row := 0
    try row := g_ProjectSelectorLv.GetNext(0)
    catch {
        return 0
    }
    return row ? Integer(row) : 0
}

ProjectSelector_SelectorHwnd() {
    global g_ProjectSelectorGui
    hwnd := 0
    try {
        if (IsObject(g_ProjectSelectorGui))
            hwnd := g_ProjectSelectorGui.Hwnd
    } catch {
        hwnd := 0
    }
    return hwnd
}

; AlwaysOnTop parent covers InputBox/DirSelect/MsgBox; drop it for the duration of those dialogs.
ProjectSelector_DialogsBegin() {
    global g_ProjectSelectorGui
    try {
        if (IsObject(g_ProjectSelectorGui))
            g_ProjectSelectorGui.Opt("-AlwaysOnTop")
    } catch {
    }
}

ProjectSelector_DialogsEnd() {
    global g_ProjectSelectorGui
    try {
        if (IsObject(g_ProjectSelectorGui))
            g_ProjectSelectorGui.Opt("+AlwaysOnTop")
    } catch {
    }
}

ProjectSelector_InputBox(prompt, title, width := 420, defaultVal := "") {
    return InputBox(prompt, title, "w" . width, defaultVal)
}

ProjectSelector_RefocusGui() {
    global g_ProjectSelectorGui, g_ProjectSelectorLv
    try {
        if (IsObject(g_ProjectSelectorGui))
            WinActivate("ahk_id " g_ProjectSelectorGui.Hwnd)
    } catch {
    }
    try {
        if (IsObject(g_ProjectSelectorLv))
            g_ProjectSelectorLv.Focus()
    } catch {
    }
}

ProjectSelector_AvailableChars(excludeIndex := 0) {
    global g_Projects, g_ProjectCharSequence
    ProjectData_Load()
    taken := Map()
    loop g_Projects.Length {
        if (A_Index = excludeIndex)
            continue
        project := g_Projects[A_Index]
        if (project.HasProp("char") && project.char != "")
            taken[project.char] := true
    }
    avail := []
    for c in g_ProjectCharSequence {
        if (!taken.Has(c))
            avail.Push(c)
    }
    return avail
}

ProjectSelector_PromptChar(currentChar := "", excludeIndex := 0) {
    avail := ProjectSelector_AvailableChars(excludeIndex)
    if (avail.Length = 0 && (currentChar = "" || !ProjectSelector_IsValidChar(currentChar))) {
        ShowNotification_WM("No free characters left. Delete a project first.")
        return ""
    }
    hint := ""
    for c in avail {
        hint .= (hint = "" ? "" : " ") . c
    }
    prompt := "Unique character from the assignment pool."
    if (hint != "")
        prompt .= "`nAvailable: " . hint
    result := ProjectSelector_InputBox(prompt, "Project character", 520, currentChar)
    if (result.Result != "OK")
        return ""
    ch := Trim(result.Value)
    ch := StrLower(ch)
    if (StrLen(ch) != 1 || !ProjectSelector_IsValidChar(ch)) {
        ShowNotification_WM("Character must be one of the assignment pool keys.")
        return ""
    }
    for c in avail {
        if (c = ch)
            return ch
    }
    if (ch = currentChar)
        return ch
    ShowNotification_WM("Character '" . ch . "' is already assigned.")
    return ""
}

ProjectSelector_OnListActivate(*) {
    global g_SelectionModeActive
    idx := ProjectSelector_SelectedIndex()
    if (idx < 1)
        return
    if (g_SelectionModeActive)
        HandleSelectionModeProjectSelection(idx)
    else
        HandleProjectSelection(idx)
}

ProjectSelector_PathPickClose() {
    global g_ProjectPathPickGui, g_ProjectPathPickPersonalEdit, g_ProjectPathPickWorkEdit
    if (IsObject(g_ProjectPathPickGui)) {
        try g_ProjectPathPickGui.Destroy()
        catch {
        }
    }
    g_ProjectPathPickGui := false
    g_ProjectPathPickPersonalEdit := false
    g_ProjectPathPickWorkEdit := false
}

ProjectSelector_PathPickOk(*) {
    global g_ProjectPathPickResult
    if (g_ProjectPathPickResult != "")
        return
    g_ProjectPathPickResult := "ok"
}

ProjectSelector_PathPickCancel(*) {
    global g_ProjectPathPickResult
    if (g_ProjectPathPickResult != "")
        return
    g_ProjectPathPickResult := "cancel"
}

ProjectSelector_NormalizePath(path) {
    p := Trim(path)
    loop 2 {
        if (StrLen(p) < 2)
            break
        first := SubStr(p, 1, 1)
        last := SubStr(p, -1)
        if ((first = '"' && last = '"') || (first = "'" && last = "'"))
            p := Trim(SubStr(p, 2, StrLen(p) - 2))
        else
            break
    }
    return p
}

ProjectSelector_PathPickBrowse(which, *) {
    global g_ProjectPathPickGui, g_ProjectPathPickPersonalEdit, g_ProjectPathPickWorkEdit
    start := ""
    try {
        if (which = "work")
            start := ProjectSelector_NormalizePath(g_ProjectPathPickWorkEdit.Value)
        else
            start := ProjectSelector_NormalizePath(g_ProjectPathPickPersonalEdit.Value)
    } catch {
        start := ""
    }
    if (start != "" && !DirExist(start))
        start := ""
    try {
        if (IsObject(g_ProjectPathPickGui))
            g_ProjectPathPickGui.Opt("-AlwaysOnTop")
    } catch {
    }
    folder := DirSelect(start, 0, (which = "work")
        ? "Select work folder"
        : "Select personal folder")
    try {
        if (IsObject(g_ProjectPathPickGui))
            g_ProjectPathPickGui.Opt("+AlwaysOnTop")
    } catch {
    }
    if (folder = "") {
        try {
            if (IsObject(g_ProjectPathPickGui))
                WinActivate("ahk_id " g_ProjectPathPickGui.Hwnd)
        } catch {
        }
        return
    }
    try {
        if (which = "work")
            g_ProjectPathPickWorkEdit.Value := folder
        else
            g_ProjectPathPickPersonalEdit.Value := folder
    } catch {
    }
    try {
        if (IsObject(g_ProjectPathPickGui))
            WinActivate("ahk_id " g_ProjectPathPickGui.Hwnd)
    } catch {
    }
}

; Paste full paths and/or Browse. Returns {path, workPath} or "" if cancelled.
ProjectSelector_PromptPaths(personalDefault := "", workDefault := "") {
    global g_ProjectPathPickResult, g_ProjectPathPickGui, g_ProjectPathPickPersonalEdit, g_ProjectPathPickWorkEdit
    ProjectSelector_PathPickClose()
    g_ProjectPathPickResult := ""

    g_ProjectPathPickGui := Gui("+AlwaysOnTop +ToolWindow", "Project folders")
    g_ProjectPathPickGui.SetFont("s10", "Segoe UI")
    g_ProjectPathPickGui.Add("Text", "w700",
        "Paste a full path (quotes optional) or Browse. At least one existing folder is required.")
    g_ProjectPathPickGui.Add("Text", "xm w700", "Personal path:")
    g_ProjectPathPickPersonalEdit := g_ProjectPathPickGui.Add("Edit", "w580 Section", personalDefault)
    g_ProjectPathPickGui.Add("Button", "ys w100", "Browse").OnEvent("Click", ProjectSelector_PathPickBrowse.Bind(
        "personal"))
    g_ProjectPathPickGui.Add("Text", "xm w700", "Work path:")
    g_ProjectPathPickWorkEdit := g_ProjectPathPickGui.Add("Edit", "w580 Section", workDefault)
    g_ProjectPathPickGui.Add("Button", "ys w100", "Browse").OnEvent("Click", ProjectSelector_PathPickBrowse.Bind("work"
    ))
    g_ProjectPathPickGui.Add("Button", "xm w100 Section Default", "OK").OnEvent("Click", ProjectSelector_PathPickOk)
    g_ProjectPathPickGui.Add("Button", "ys w100", "Cancel").OnEvent("Click", ProjectSelector_PathPickCancel)
    g_ProjectPathPickGui.OnEvent("Close", ProjectSelector_PathPickCancel)
    g_ProjectPathPickGui.OnEvent("Escape", ProjectSelector_PathPickCancel)
    g_ProjectPathPickGui.Show()
    try g_ProjectPathPickPersonalEdit.Focus()
    catch {
    }

    while (g_ProjectPathPickResult = "")
        Sleep 50

    personal := ""
    work := ""
    try personal := ProjectSelector_NormalizePath(g_ProjectPathPickPersonalEdit.Value)
    catch {
    }
    try work := ProjectSelector_NormalizePath(g_ProjectPathPickWorkEdit.Value)
    catch {
    }
    result := g_ProjectPathPickResult
    g_ProjectPathPickResult := ""
    ProjectSelector_PathPickClose()
    if (result != "ok")
        return ""
    return { path: personal, workPath: work }
}

ProjectSelector_OnAdd(*) {
    global g_Projects
    ProjectData_Load()
    ProjectSelector_DialogsBegin()
    nameBox := ProjectSelector_InputBox("Project name:", "Add project")
    if (nameBox.Result != "OK") {
        ProjectSelector_DialogsEnd()
        ProjectSelector_RefocusGui()
        return
    }
    name := Trim(nameBox.Value)
    if (name = "") {
        ProjectSelector_DialogsEnd()
        ShowNotification_WM("Name is required.")
        ProjectSelector_RefocusGui()
        return
    }
    ch := ProjectSelector_PromptChar()
    if (ch = "") {
        ProjectSelector_DialogsEnd()
        ProjectSelector_RefocusGui()
        return
    }
    paths := ProjectSelector_PromptPaths()
    ProjectSelector_DialogsEnd()
    if (!IsObject(paths)) {
        ProjectSelector_RefocusGui()
        return
    }
    personalPath := paths.path
    workPath := paths.workPath
    if ((personalPath = "" || !DirExist(personalPath)) && (workPath = "" || !DirExist(workPath))) {
        ShowNotification_WM("At least one existing folder is required.")
        ProjectSelector_RefocusGui()
        return
    }
    list := []
    for project in g_Projects
        list.Push(project)
    list.Push({ name: name, char: ch, path: personalPath, workPath: workPath })
    if (!ProjectData_Save(list)) {
        ShowNotification_WM("Failed to save project.")
        ProjectSelector_RefocusGui()
        return
    }
    ProjectSelector_PopulateLv()
    ProjectSelector_BindModalHotkeys()
    ProjectSelector_RefocusGui()
}

ProjectSelector_OnEdit(*) {
    global g_Projects, g_ProjectSelectorLv
    idx := ProjectSelector_SelectedIndex()
    if (idx < 1) {
        ShowNotification_WM("Select a project to edit.")
        return
    }
    ProjectData_Load()
    if (idx > g_Projects.Length)
        return
    project := g_Projects[idx]
    ProjectSelector_DialogsBegin()
    nameBox := ProjectSelector_InputBox("Project name:", "Edit project", 420, project.name)
    if (nameBox.Result != "OK") {
        ProjectSelector_DialogsEnd()
        ProjectSelector_RefocusGui()
        return
    }
    name := Trim(nameBox.Value)
    if (name = "") {
        ProjectSelector_DialogsEnd()
        ShowNotification_WM("Name is required.")
        ProjectSelector_RefocusGui()
        return
    }
    currentChar := project.HasProp("char") ? project.char : ""
    ch := ProjectSelector_PromptChar(currentChar, idx)
    ProjectSelector_DialogsEnd()
    if (ch = "") {
        ProjectSelector_RefocusGui()
        return
    }
    list := []
    loop g_Projects.Length {
        item := g_Projects[A_Index]
        if (A_Index = idx)
            list.Push({ name: name, char: ch, path: item.path, workPath: item.workPath })
        else
            list.Push(item)
    }
    if (!ProjectData_Save(list)) {
        ShowNotification_WM("Failed to save project.")
        ProjectSelector_RefocusGui()
        return
    }
    ProjectSelector_PopulateLv()
    ProjectSelector_BindModalHotkeys()
    try g_ProjectSelectorLv.Modify(idx, "Select Focus Vis")
    catch {
    }
    ProjectSelector_RefocusGui()
}

ProjectSelector_OnDelete(*) {
    global g_Projects
    idx := ProjectSelector_SelectedIndex()
    if (idx < 1) {
        ShowNotification_WM("Select a project to delete.")
        return
    }
    ProjectData_Load()
    if (idx > g_Projects.Length)
        return
    project := g_Projects[idx]
    label := project.name != "" ? project.name : "(unnamed)"
    hwnd := ProjectSelector_SelectorHwnd()
    msgOpts := "YesNo Icon! Default2"
    if (hwnd)
        msgOpts .= " Owner" . hwnd
    ProjectSelector_DialogsBegin()
    confirmed := (MsgBox("Delete project '" . label . "'?`nThis frees its assigned character.",
        "Delete project", msgOpts) = "Yes")
    ProjectSelector_DialogsEnd()
    if (!confirmed) {
        ProjectSelector_RefocusGui()
        return
    }
    list := []
    loop g_Projects.Length {
        if (A_Index != idx)
            list.Push(g_Projects[A_Index])
    }
    if (!ProjectData_Save(list)) {
        ShowNotification_WM("Failed to save project list.")
        ProjectSelector_RefocusGui()
        return
    }
    ProjectSelector_PopulateLv()
    ProjectSelector_BindModalHotkeys()
    ProjectSelector_RefocusGui()
}

; Cleanup project selector: destroy GUI, disable hotkeys, reset state
CleanupProjectSelector() {
    global g_ProjectSelectorActive, g_ProjectSelectorGui, g_ProjectSelectorLv, g_ProjectHotkeyHandlers,
        g_SelectionModeActive, g_CopyFromGeminiModeActive, g_WM_SelectorOpenFile,
        g_WM_SelectorCloseRequestFile, g_WM_SelectorCloseCheckTimer, g_OnEscapePressed

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

    ProjectSelector_UnbindModalHotkeys()

    ; Unregister Escape callback so Utils forwards Escape again
    g_OnEscapePressed := ""

    ; Close and destroy GUI
    if (IsObject(g_ProjectSelectorGui)) {
        try {
            g_ProjectSelectorGui.Destroy()
        } catch {
            ; Ignore
        }
        g_ProjectSelectorGui := false
    }
    g_ProjectSelectorLv := false
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
    ProjectData_Load()

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
    global g_ProjectSelectorActive, g_OnEscapePressed
    if (g_ProjectSelectorActive) {
        CleanupProjectSelector()
        return true
    }
    g_OnEscapePressed := ""
    return false
}
