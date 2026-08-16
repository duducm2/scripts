; =============================================================================
; Utils module: prompt_editor_gui.ahk
; Single-form Add/Edit prompt dialog (identity + personal/work context file lists)
; =============================================================================

global g_PromptEditorGui := false
global g_PromptEditorResult := { saved: false }
global g_PromptEditorName := false
global g_PromptEditorCategory := false
global g_PromptEditorChar := false
global g_PromptEditorFile := false
global g_PromptEditorFilePath := ""
global g_PromptEditorSource := "file"
global g_PromptEditorAuthor := ""
global g_PromptEditorPersonalLv := false
global g_PromptEditorWorkLv := false
global g_PromptEditorIsEdit := false
global g_PromptEditorListIndex := 0
global g_PromptEditorPersonalPaths := []
global g_PromptEditorWorkPaths := []

PromptEditor_Show(existingPrompt := false, listIndex := 0) {
    global g_PromptEntries, g_PromptEditorGui, g_PromptEditorResult
    global g_PromptEditorIsEdit, g_PromptEditorListIndex, g_PromptEditorAuthor
    global g_PromptEditorFilePath, g_PromptEditorSource
    global g_PromptEditorPersonalPaths, g_PromptEditorWorkPaths

    PromptData_Load()
    g_PromptEditorIsEdit := IsObject(existingPrompt)
    g_PromptEditorListIndex := listIndex
    g_PromptEditorResult := { saved: false }
    g_PromptEditorAuthor := (g_PromptEditorIsEdit && existingPrompt.HasProp("author")) ? existingPrompt.author : ""
    g_PromptEditorFilePath := (g_PromptEditorIsEdit && existingPrompt.HasProp("filePath")) ? existingPrompt.filePath :
        ""
    g_PromptEditorSource := (g_PromptEditorIsEdit && existingPrompt.HasProp("source")) ? existingPrompt.source : "file"
    g_PromptEditorPersonalPaths := PromptData_ParsePathList((g_PromptEditorIsEdit && existingPrompt.HasProp(
        "personal_context_files")) ? existingPrompt.personal_context_files : [])
    g_PromptEditorWorkPaths := PromptData_ParsePathList((g_PromptEditorIsEdit && existingPrompt.HasProp(
        "work_context_files")) ? existingPrompt.work_context_files : [])

    currentChar := (g_PromptEditorIsEdit && existingPrompt.HasProp("char")) ? existingPrompt.char : ""
    avail := UtilitySelector_AvailableChars(PromptData_CharSequence, PromptData_IsValidChar, g_PromptEntries, "char",
        listIndex)
    if (currentChar != "" && PromptData_IsValidChar(currentChar)) {
        already := false
        for c in avail {
            if (c = currentChar) {
                already := true
                break
            }
        }
        if (!already)
            avail.InsertAt(1, currentChar)
    }
    if (avail.Length = 0) {
        UtilitySelector_Notify("No free characters left. Delete an item first.")
        return g_PromptEditorResult
    }

    ownerOpt := "+AlwaysOnTop +ToolWindow"
    hwnd := UtilitySelector_SelectorHwnd()
    if (hwnd)
        ownerOpt .= " +Owner" . hwnd

    title := g_PromptEditorIsEdit ? "Edit prompt" : "Add prompt"
    g_PromptEditorGui := Gui(ownerOpt, title)
    g_PromptEditorGui.SetFont("s10", "Segoe UI")
    PromptEditor_BuildControls(existingPrompt, avail, currentChar)
    g_PromptEditorGui.OnEvent("Close", PromptEditor_OnCancel)
    g_PromptEditorGui.OnEvent("Escape", PromptEditor_OnCancel)

    UtilitySelector_DialogsBegin()
    PromptEditor_CenterOnSelector()
    try g_PromptEditorName.Focus()
    catch {
    }
    PromptEditor_BindEditorHotkeys(true)
    try WinWaitClose("ahk_id " g_PromptEditorGui.Hwnd)
    catch {
    }
    PromptEditor_BindEditorHotkeys(false)
    UtilitySelector_DialogsEnd()
    g_PromptEditorGui := false
    return g_PromptEditorResult
}

PromptEditor_BuildControls(existingPrompt, avail, currentChar) {
    global g_PromptEditorGui, g_PromptEditorName, g_PromptEditorCategory, g_PromptEditorChar
    global g_PromptEditorFile, g_PromptEditorFilePath, g_PromptEditorIsEdit
    global g_PromptEditorPersonalLv, g_PromptEditorWorkLv

    colW := 360
    g_PromptEditorGui.Add("Text", "xm w80", "Name")
    nameVal := (IsObject(existingPrompt) && existingPrompt.HasProp("name")) ? existingPrompt.name : ""
    g_PromptEditorName := g_PromptEditorGui.Add("Edit", "yp w" . (colW * 2 - 80), nameVal)

    g_PromptEditorGui.Add("Text", "xm w80", "Category")
    catChoices := PromptEditor_CategoryChoices(IsObject(existingPrompt) ? existingPrompt.category : "General")
    g_PromptEditorCategory := g_PromptEditorGui.Add("ComboBox", "yp w220", catChoices)
    if (IsObject(existingPrompt) && existingPrompt.HasProp("category") && existingPrompt.category != "")
        g_PromptEditorCategory.Text := existingPrompt.category
    else
        g_PromptEditorCategory.Text := "General"

    g_PromptEditorGui.Add("Text", "x+16 yp w40", "Char")
    g_PromptEditorChar := g_PromptEditorGui.Add("DropDownList", "yp w80", avail)
    chooseChar := currentChar != "" ? currentChar : avail[1]
    try g_PromptEditorChar.Text := chooseChar
    catch {
        try g_PromptEditorChar.Choose(1)
        catch {
        }
    }

    g_PromptEditorGui.Add("Text", "xm w80", "Prompt file")
    g_PromptEditorFile := g_PromptEditorGui.Add("Edit", "yp w520 ReadOnly", g_PromptEditorFilePath)
    g_PromptEditorGui.Add("Button", "yp w80", "Browse").OnEvent("Click", PromptEditor_OnBrowsePromptFile)

    g_PromptEditorGui.Add("Text", "xm w" . colW . " Section", "Personal context files")
    g_PromptEditorGui.Add("Text", "ys w" . colW, "Work context files")
    g_PromptEditorPersonalLv := g_PromptEditorGui.Add("ListView", "xm w" . colW . " r8", ["Path"])
    g_PromptEditorWorkLv := g_PromptEditorGui.Add("ListView", "x+12 yp w" . colW . " r8", ["Path"])
    g_PromptEditorPersonalLv.ModifyCol(1, colW - 24)
    g_PromptEditorWorkLv.ModifyCol(1, colW - 24)
    PromptEditor_ReloadList("personal")
    PromptEditor_ReloadList("work")

    g_PromptEditorGui.Add("Button", "xm w110", "Add files").OnEvent("Click", (*) => PromptEditor_OnAddFiles("personal"))
    g_PromptEditorGui.Add("Button", "x+8 yp w110", "Paste paths").OnEvent("Click", (*) => PromptEditor_OnPastePaths(
        "personal"))
    g_PromptEditorGui.Add("Button", "x+8 yp w110", "Remove").OnEvent("Click", (*) => PromptEditor_OnRemove("personal"))
    g_PromptEditorGui.Add("Button", "x+24 yp w110", "Add files").OnEvent("Click", (*) => PromptEditor_OnAddFiles("work"
    ))
    g_PromptEditorGui.Add("Button", "x+8 yp w110", "Paste paths").OnEvent("Click", (*) => PromptEditor_OnPastePaths(
        "work"))
    g_PromptEditorGui.Add("Button", "x+8 yp w110", "Remove").OnEvent("Click", (*) => PromptEditor_OnRemove("work"))

    g_PromptEditorGui.Add("Text", "xm w" . (colW * 2 + 12),
    "Paste Explorer Copy as path; quotes are stripped. Empty lists are fine.")
    g_PromptEditorGui.Add("Button", "xm+430 w100 Default", "Save").OnEvent("Click", PromptEditor_OnSave)
    g_PromptEditorGui.Add("Button", "x+8 yp w100", "Cancel").OnEvent("Click", PromptEditor_OnCancel)
}

PromptEditor_CategoryChoices(current) {
    global g_PromptEntries
    seen := Map()
    choices := ["General"]
    seen["general"] := true
    if (current != "" && !seen.Has(StrLower(current))) {
        choices.Push(current)
        seen[StrLower(current)] := true
    }
    for p in g_PromptEntries {
        cat := p.HasProp("category") ? Trim(p.category) : ""
        if (cat = "" || seen.Has(StrLower(cat)))
            continue
        seen[StrLower(cat)] := true
        choices.Push(cat)
    }
    return choices
}

PromptEditor_CenterOnSelector() {
    global g_PromptEditorGui
    mon := UtilitySelector_ActiveMonitorWorkArea()
    g_PromptEditorGui.Show("Hide")
    g_PromptEditorGui.GetPos(, , &gw, &gh)
    guiX := mon.left + (mon.width - gw) // 2
    guiY := mon.top + (mon.height - gh) // 2
    if (guiX < mon.left + 20)
        guiX := mon.left + 20
    if (guiY < mon.top + 20)
        guiY := mon.top + 20
    g_PromptEditorGui.Show("x" . guiX . " y" . guiY)
}

PromptEditor_BindEditorHotkeys(enable) {
    global g_PromptEditorGui
    hwnd := 0
    try {
        if (IsObject(g_PromptEditorGui))
            hwnd := g_PromptEditorGui.Hwnd
    } catch {
        hwnd := 0
    }
    if (!hwnd)
        return
    try HotIfWinActive("ahk_id " hwnd)
    catch {
        return
    }
    try Hotkey("Delete", PromptEditor_OnDeleteKey, enable ? "On" : "Off")
    catch {
    }
    try HotIf()
    catch {
    }
}

PromptEditor_FocusedSide() {
    global g_PromptEditorGui, g_PromptEditorPersonalLv, g_PromptEditorWorkLv
    focused := 0
    try focused := ControlGetFocus("ahk_id " g_PromptEditorGui.Hwnd)
    catch {
        return "personal"
    }
    if (IsObject(g_PromptEditorWorkLv) && focused = g_PromptEditorWorkLv.Hwnd)
        return "work"
    return "personal"
}

PromptEditor_OnDeleteKey(*) {
    PromptEditor_OnRemove(PromptEditor_FocusedSide())
}

PromptEditor_PathsRef(side) {
    global g_PromptEditorPersonalPaths, g_PromptEditorWorkPaths
    return (side = "work") ? g_PromptEditorWorkPaths : g_PromptEditorPersonalPaths
}

PromptEditor_Lv(side) {
    global g_PromptEditorPersonalLv, g_PromptEditorWorkLv
    return (side = "work") ? g_PromptEditorWorkLv : g_PromptEditorPersonalLv
}

PromptEditor_SetPaths(side, arr) {
    global g_PromptEditorPersonalPaths, g_PromptEditorWorkPaths
    if (side = "work")
        g_PromptEditorWorkPaths := arr
    else
        g_PromptEditorPersonalPaths := arr
}

PromptEditor_ReloadList(side) {
    lv := PromptEditor_Lv(side)
    if (!IsObject(lv))
        return
    lv.Delete()
    for p in PromptEditor_PathsRef(side)
        lv.Add("", p)
}

PromptEditor_HasPath(arr, path) {
    needle := StrLower(path)
    for p in arr {
        if (StrLower(p) = needle)
            return true
    }
    return false
}

PromptEditor_AppendPaths(side, incoming) {
    arr := PromptEditor_PathsRef(side)
    for raw in incoming {
        n := PromptData_StripPathQuotes(raw)
        if (n = "" || PromptEditor_HasPath(arr, n))
            continue
        arr.Push(n)
    }
    PromptEditor_SetPaths(side, arr)
    PromptEditor_ReloadList(side)
}

PromptEditor_ParseFileSelect(result) {
    paths := []
    if (result = "" || result = 0)
        return paths
    if (IsObject(result)) {
        if (result.Length = 0)
            return paths
        first := result[1]
        if (result.Length >= 2 && DirExist(first)) {
            second := result[2]
            if (!InStr(second, "\") && !InStr(second, "/")) {
                dir := RTrim(first, "\")
                i := 2
                while (i <= result.Length) {
                    name := Trim(result[i])
                    if (name != "")
                        paths.Push(dir "\" name)
                    i++
                }
                return paths
            }
        }
        for p in result {
            n := Trim(p)
            if (n != "")
                paths.Push(n)
        }
        return paths
    }
    text := StrReplace(StrReplace(result, "`r`n", "`n"), "`r", "`n")
    if (!InStr(text, "`n")) {
        paths.Push(result)
        return paths
    }
    lines := StrSplit(text, "`n")
    dir := RTrim(Trim(lines[1]), "\")
    i := 2
    while (i <= lines.Length) {
        name := Trim(lines[i])
        if (name != "")
            paths.Push(dir "\" name)
        i++
    }
    return paths
}

PromptEditor_OnBrowsePromptFile(*) {
    global g_PromptEditorGui, g_PromptEditorFile, g_PromptEditorFilePath, g_PromptEditorSource
    start := g_PromptEditorFilePath != "" ? PromptData_ResolvePath({ filePath: g_PromptEditorFilePath,
        source: g_PromptEditorSource }) : A_ScriptDir "\assets\prompt\"
    if (IsObject(g_PromptEditorGui)) {
        try g_PromptEditorGui.Opt("-AlwaysOnTop")
        catch {
        }
    }
    selected := FileSelect(1, start, "Select prompt file", "Text (*.txt)")
    if (IsObject(g_PromptEditorGui)) {
        try g_PromptEditorGui.Opt("+AlwaysOnTop")
        try WinActivate("ahk_id " g_PromptEditorGui.Hwnd)
        catch {
        }
    }
    if (selected = "")
        return
    g_PromptEditorFilePath := PromptData_ToStoredPath(selected)
    g_PromptEditorSource := "file"
    try g_PromptEditorFile.Value := g_PromptEditorFilePath
    catch {
    }
}

PromptEditor_OnAddFiles(side) {
    global g_PromptEditorGui
    if (IsObject(g_PromptEditorGui)) {
        try g_PromptEditorGui.Opt("-AlwaysOnTop")
        catch {
        }
    }
    selected := FileSelect("M1", , "Select context files")
    if (IsObject(g_PromptEditorGui)) {
        try g_PromptEditorGui.Opt("+AlwaysOnTop")
        try WinActivate("ahk_id " g_PromptEditorGui.Hwnd)
        catch {
        }
    }
    PromptEditor_AppendPaths(side, PromptEditor_ParseFileSelect(selected))
}

PromptEditor_OnPastePaths(side) {
    raw := ""
    try raw := A_Clipboard
    catch {
        return
    }
    PromptEditor_AppendPaths(side, PromptData_ParsePathList(raw))
}

PromptEditor_OnRemove(side) {
    lv := PromptEditor_Lv(side)
    if (!IsObject(lv))
        return
    selected := []
    row := 0
    loop {
        row := lv.GetNext(row)
        if (!row)
            break
        selected.Push(Integer(row))
    }
    if (selected.Length = 0)
        return
    arr := PromptEditor_PathsRef(side)
    keep := []
    skip := Map()
    for idx in selected
        skip[idx] := true
    loop arr.Length {
        if (!skip.Has(A_Index))
            keep.Push(arr[A_Index])
    }
    PromptEditor_SetPaths(side, keep)
    PromptEditor_ReloadList(side)
}

PromptEditor_OnCancel(*) {
    global g_PromptEditorResult, g_PromptEditorGui
    g_PromptEditorResult := { saved: false }
    PromptEditor_Destroy()
}

PromptEditor_OnSave(*) {
    global g_PromptEditorResult, g_PromptEditorName, g_PromptEditorCategory, g_PromptEditorChar
    global g_PromptEditorFilePath, g_PromptEditorSource, g_PromptEditorAuthor, g_PromptEditorIsEdit
    global g_PromptEditorPersonalPaths, g_PromptEditorWorkPaths

    name := Trim(g_PromptEditorName.Value)
    if (name = "") {
        UtilitySelector_Notify("Name is required.")
        try g_PromptEditorName.Focus()
        catch {
        }
        return
    }
    category := Trim(g_PromptEditorCategory.Text)
    if (category = "")
        category := "General"
    ch := StrLower(Trim(g_PromptEditorChar.Text))
    if (ch = "" || !PromptData_IsValidChar(ch)) {
        UtilitySelector_Notify("Character must be one of the assignment pool keys.")
        return
    }
    if (g_PromptEditorFilePath = "") {
        UtilitySelector_Notify("Prompt file is required.")
        return
    }
    g_PromptEditorResult := {
        saved: true,
        name: name,
        category: category,
        char: ch,
        filePath: g_PromptEditorFilePath,
        source: g_PromptEditorSource,
        author: g_PromptEditorAuthor,
        personal_context_files: PromptData_ParsePathList(g_PromptEditorPersonalPaths),
        work_context_files: PromptData_ParsePathList(g_PromptEditorWorkPaths)
    }
    PromptEditor_Destroy()
}

PromptEditor_Destroy() {
    global g_PromptEditorGui
    PromptEditor_BindEditorHotkeys(false)
    if (IsObject(g_PromptEditorGui)) {
        try g_PromptEditorGui.Destroy()
        catch {
        }
    }
    g_PromptEditorGui := false
}
