; =============================================================================
; Utils module: clip_angel_export_desktop.ahk
; Copy Clip Angel Row 0 (newest clip) preview text to a UTF-8 .txt on Desktop,
; then prompt to rename from a persisted name list (inline CRUD).
; Utility Shortcuts: #!+U → Macros → [c]
; =============================================================================

global g_ClipAngelNameGui := false
global g_ClipAngelNameLv := false
global g_ClipAngelNameRows := []
global g_ClipAngelNameHotkeys := []
global g_ClipAngelNameSourcePath := ""
global g_ClipAngelNameFinalPath := ""
global g_ClipAngelNamePicked := false

ClipAngelExport_NamesCsvPath() {
    return A_ScriptDir "\assets\data\clipangel_desktop_names.csv"
}

ClipAngelExport_CsvEscape(val) {
    s := String(val)
    if (InStr(s, ",") || InStr(s, '"') || InStr(s, "`n") || InStr(s, "`r"))
        return '"' . StrReplace(s, '"', '""') . '"'
    return s
}

ClipAngelExport_SplitCsvLine(line) {
    fields := []
    i := 1
    len := StrLen(line)
    if (len && SubStr(line, len, 1) = "`r") {
        line := SubStr(line, 1, len - 1)
        len := StrLen(line)
    }
    while (i <= len) {
        if (SubStr(line, i, 1) = '"') {
            i += 1
            val := ""
            while (i <= len) {
                c := SubStr(line, i, 1)
                if (c = '"') {
                    if (i < len && SubStr(line, i + 1, 1) = '"') {
                        val .= '"'
                        i += 2
                        continue
                    }
                    i += 1
                    break
                }
                val .= c
                i += 1
            }
            fields.Push(val)
            if (i <= len && SubStr(line, i, 1) = ",")
                i += 1
            continue
        }
        j := i
        while (j <= len && SubStr(line, j, 1) != ",")
            j += 1
        fields.Push(SubStr(line, i, j - i))
        i := j + 1
    }
    return fields
}

ClipAngelExport_LoadNames() {
    path := ClipAngelExport_NamesCsvPath()
    rows := []
    if (!FileExist(path))
        return rows
    try text := ReadUtf8File(path)
    catch {
        return rows
    }
    if (text = "")
        return rows
    if (SubStr(text, 1, 1) = Chr(0xFEFF))
        text := SubStr(text, 2)
    text := StrReplace(text, "`r`n", "`n")
    text := StrReplace(text, "`r", "`n")
    headers := []
    for line in StrSplit(text, "`n") {
        if (Trim(line) = "")
            continue
        fields := ClipAngelExport_SplitCsvLine(line)
        if (headers.Length = 0) {
            for h in fields
                headers.Push(Trim(h))
            continue
        }
        row := Map()
        loop headers.Length {
            key := headers[A_Index]
            row[key] := (A_Index <= fields.Length) ? fields[A_Index] : ""
        }
        if (row.Has("name") && Trim(row["name"]) != "")
            rows.Push(row)
    }
    return rows
}

ClipAngelExport_SaveNames(rows) {
    path := ClipAngelExport_NamesCsvPath()
    dir := ""
    SplitPath(path, , &dir)
    if (dir != "" && !DirExist(dir))
        DirCreate(dir)
    out := "id,name`n"
    for row in rows {
        id := row.Has("id") ? row["id"] : ""
        name := row.Has("name") ? row["name"] : ""
        out .= ClipAngelExport_CsvEscape(id) "," ClipAngelExport_CsvEscape(name) "`n"
    }
    WriteUtf8File(path, out)
}

ClipAngelExport_Unaccent(s) {
    pairs := [["á", "a"], ["à", "a"], ["â", "a"], ["ã", "a"], ["ä", "a"], ["é", "e"], ["ê", "e"], ["è", "e"],
    ["í", "i"], ["ó", "o"], ["ô", "o"], ["õ", "o"], ["ö", "o"], ["ú", "u"], ["ü", "u"], ["ç", "c"],
    ["Á", "A"], ["À", "A"], ["Â", "A"], ["Ã", "A"], ["É", "E"], ["Ê", "E"], ["Í", "I"], ["Ó", "O"],
    ["Ô", "O"], ["Õ", "O"], ["Ú", "U"], ["Ü", "U"], ["Ç", "C"]]
    for p in pairs
        s := StrReplace(s, p[1], p[2])
    return s
}

ClipAngelExport_Slug(name) {
    s := ClipAngelExport_Unaccent(Trim(name))
    s := StrUpper(s)
    out := ""
    loop parse s {
        c := Ord(A_LoopField)
        if ((c >= 65 && c <= 90) || (c >= 48 && c <= 57))
            out .= A_LoopField
    }
    if (StrLen(out) > 8)
        out := SubStr(out, 1, 8)
    if (out = "")
        out := "X"
    return out
}

ClipAngelExport_IdExists(rows, id) {
    for row in rows {
        if (row.Has("id") && row["id"] = id)
            return true
    }
    return false
}

ClipAngelExport_SlugId(name, existing) {
    base := ClipAngelExport_Slug(name)
    id := "NAME_" base
    if (!ClipAngelExport_IdExists(existing, id))
        return id
    n := 2
    loop {
        cand := id . n
        if (!ClipAngelExport_IdExists(existing, cand))
            return cand
        n += 1
    }
}

ClipAngelExport_SanitizeFileName(name) {
    s := Trim(name)
    for ch in ["\", "/", ":", "*", "?", '"', "<", ">", "|"]
        s := StrReplace(s, ch, "")
    s := Trim(s)
    while (InStr(s, "  "))
        s := StrReplace(s, "  ", " ")
    return s
}

ClipAngelExport_UniqueNamedPath(desktopDir, baseName) {
    path := desktopDir "\" baseName ".txt"
    if !FileExist(path)
        return path
    i := 2
    loop {
        path := desktopDir "\" baseName "-" i ".txt"
        if !FileExist(path)
            return path
        i += 1
    }
}

ClipAngelExport_UniqueDesktopPath(desktopDir) {
    stamp := FormatTime(, "yyyyMMdd-HHmmss")
    base := "clipangel-last-" stamp
    path := desktopDir "\" base ".txt"
    if !FileExist(path)
        return path
    i := 2
    loop {
        path := desktopDir "\" base "-" i ".txt"
        if !FileExist(path)
            return path
        i += 1
    }
}

ClipAngelExport_CenterGui(guiObj, w := 420, h := 420) {
    MonitorGetWorkArea(MonitorGetPrimary(), &L, &T, &R, &B)
    x := L + ((R - L) - w) // 2
    y := T + ((B - T) - h) // 2
    guiObj.Show("x" . x . " y" . y . " w" . w . " h" . h)
}

ClipAngelExport_HotIfActive(*) {
    global g_ClipAngelNameGui
    try {
        if (!IsObject(g_ClipAngelNameGui))
            return false
        if !WinActive("ahk_id " g_ClipAngelNameGui.Hwnd)
            return false
        focused := ""
        try focused := g_ClipAngelNameGui.FocusedCtrl
        catch {
            focused := ""
        }
        if (IsObject(focused) && focused.Type = "Edit")
            return false
        return true
    } catch {
        return false
    }
}

ClipAngelExport_UnbindHotkeys() {
    global g_ClipAngelNameHotkeys
    try HotIf(ClipAngelExport_HotIfActive)
    catch {
        try HotIf()
        catch {
        }
        g_ClipAngelNameHotkeys := []
        return
    }
    for key in g_ClipAngelNameHotkeys {
        try Hotkey(key, "Off")
        catch {
        }
    }
    try HotIf()
    catch {
    }
    g_ClipAngelNameHotkeys := []
}

ClipAngelExport_BindHotkeys(pairs) {
    global g_ClipAngelNameGui, g_ClipAngelNameHotkeys
    ClipAngelExport_UnbindHotkeys()
    if (!IsObject(g_ClipAngelNameGui))
        return
    try HotIf(ClipAngelExport_HotIfActive)
    catch {
        return
    }
    for p in pairs {
        try {
            Hotkey(p[1], p[2], "On")
            g_ClipAngelNameHotkeys.Push(p[1])
        } catch {
        }
    }
    try HotIf()
    catch {
    }
}

ClipAngelExport_CloseGui() {
    global g_ClipAngelNameGui, g_ClipAngelNameLv, g_ClipAngelNameRows
    ClipAngelExport_UnbindHotkeys()
    try {
        if (IsObject(g_ClipAngelNameGui))
            g_ClipAngelNameGui.Destroy()
    } catch {
    }
    g_ClipAngelNameGui := false
    g_ClipAngelNameLv := false
    g_ClipAngelNameRows := []
}

ClipAngelExport_Refresh() {
    global g_ClipAngelNameLv, g_ClipAngelNameRows
    if (!IsObject(g_ClipAngelNameLv))
        return
    rows := ClipAngelExport_LoadNames()
    g_ClipAngelNameLv.Delete()
    g_ClipAngelNameRows := []
    for r in rows {
        g_ClipAngelNameRows.Push(r)
        g_ClipAngelNameLv.Add("", r["name"])
    }
    g_ClipAngelNameLv.ModifyCol(1, "AutoHdr")
    if (g_ClipAngelNameRows.Length > 0)
        g_ClipAngelNameLv.Modify(1, "Select Focus Vis")
}

ClipAngelExport_Selected() {
    global g_ClipAngelNameLv, g_ClipAngelNameRows
    if (!IsObject(g_ClipAngelNameLv))
        return false
    row := g_ClipAngelNameLv.GetNext()
    if (!row || row > g_ClipAngelNameRows.Length)
        return false
    return g_ClipAngelNameRows[row]
}

ClipAngelExport_DialogsBegin() {
    global g_ClipAngelNameGui
    try {
        if (IsObject(g_ClipAngelNameGui))
            g_ClipAngelNameGui.Opt("-AlwaysOnTop")
    } catch {
    }
}

ClipAngelExport_DialogsEnd() {
    global g_ClipAngelNameGui
    try {
        if (IsObject(g_ClipAngelNameGui))
            g_ClipAngelNameGui.Opt("+AlwaysOnTop")
    } catch {
    }
}

ClipAngelExport_OwnerOpt() {
    global g_ClipAngelNameGui
    hwnd := 0
    try {
        if (IsObject(g_ClipAngelNameGui))
            hwnd := g_ClipAngelNameGui.Hwnd
    } catch {
        hwnd := 0
    }
    return hwnd ? " Owner" . hwnd : ""
}

ClipAngelExport_Confirm(msg) {
    ClipAngelExport_DialogsBegin()
    result := MsgBox(msg, "ClipAngel names", "YesNo Icon?" . ClipAngelExport_OwnerOpt())
    ClipAngelExport_DialogsEnd()
    return result = "Yes"
}

ClipAngelExport_Alert(msg) {
    ClipAngelExport_DialogsBegin()
    MsgBox(msg, "ClipAngel names", "Icon!" . ClipAngelExport_OwnerOpt())
    ClipAngelExport_DialogsEnd()
}

ClipAngelExport_NameForm(existing) {
    isEdit := IsObject(existing)
    ClipAngelExport_DialogsBegin()
    ; InputBox does not support Owner=hwnd (MsgBox does).
    ib := InputBox("Name for Desktop .txt file", isEdit ? "Edit name" : "Add name",
        "w320", isEdit ? existing["name"] : "")
    ClipAngelExport_DialogsEnd()
    if (ib.Result != "OK")
        return
    name := Trim(ib.Value)
    if (name = "") {
        ClipAngelExport_Alert("Name is required.")
        return
    }
    if (ClipAngelExport_SanitizeFileName(name) = "") {
        ClipAngelExport_Alert("Name has no valid filename characters.")
        return
    }
    rows := ClipAngelExport_LoadNames()
    row := Map(
        "id", isEdit ? existing["id"] : ClipAngelExport_SlugId(name, rows),
    "name", name)
    if (isEdit) {
        out := []
        for r in rows {
            if (r["id"] = existing["id"])
                out.Push(row)
            else
                out.Push(r)
        }
        rows := out
    } else {
        rows.Push(row)
    }
    ClipAngelExport_SaveNames(rows)
    ClipAngelExport_Refresh()
}

ClipAngelExport_Add(*) {
    ClipAngelExport_NameForm(false)
}

ClipAngelExport_Edit(*) {
    sel := ClipAngelExport_Selected()
    if (!sel) {
        try ShowCenteredOverlay_Utils("Select a name", 1200, BANNER_ACCENT_ERROR)
        catch {
        }
        return
    }
    ClipAngelExport_NameForm(sel)
}

ClipAngelExport_Delete(*) {
    sel := ClipAngelExport_Selected()
    if (!sel)
        return
    if (!ClipAngelExport_Confirm("Delete " . sel["name"] . "?"))
        return
    out := []
    for r in ClipAngelExport_LoadNames() {
        if (r["id"] != sel["id"])
            out.Push(r)
    }
    ClipAngelExport_SaveNames(out)
    ClipAngelExport_Refresh()
}

ClipAngelExport_UseSelected(*) {
    global g_ClipAngelNameSourcePath, g_ClipAngelNameFinalPath, g_ClipAngelNamePicked
    sel := ClipAngelExport_Selected()
    if (!sel) {
        try ShowCenteredOverlay_Utils("Select a name", 1200, BANNER_ACCENT_ERROR)
        catch {
        }
        return
    }
    clean := ClipAngelExport_SanitizeFileName(sel["name"])
    if (clean = "") {
        ClipAngelExport_Alert("Selected name is not a valid filename.")
        return
    }
    SplitPath(g_ClipAngelNameSourcePath, , &desktopDir)
    dest := ClipAngelExport_UniqueNamedPath(desktopDir, clean)
    if (dest = g_ClipAngelNameSourcePath) {
        g_ClipAngelNameFinalPath := g_ClipAngelNameSourcePath
        g_ClipAngelNamePicked := true
        ClipAngelExport_CloseGui()
        return
    }
    try {
        FileMove(g_ClipAngelNameSourcePath, dest)
        g_ClipAngelNameFinalPath := dest
        g_ClipAngelNamePicked := true
        ClipAngelExport_CloseGui()
    } catch Error as e {
        ClipAngelExport_Alert("Could not rename: " . e.Message)
    }
}

ClipAngelExport_Cancel(*) {
    ClipAngelExport_CloseGui()
}

; Shows rename picker. Returns final path (renamed or original if Esc/close).
ClipAngelExport_PromptRename(sourcePath) {
    global g_ClipAngelNameGui, g_ClipAngelNameLv, g_ClipAngelNameSourcePath
    global g_ClipAngelNameFinalPath, g_ClipAngelNamePicked
    g_ClipAngelNameSourcePath := sourcePath
    g_ClipAngelNameFinalPath := sourcePath
    g_ClipAngelNamePicked := false

    ClipAngelExport_CloseGui()
    g_ClipAngelNameGui := Gui("+AlwaysOnTop +ToolWindow", "Name Desktop file")
    g_ClipAngelNameGui.SetFont("s10", "Segoe UI")
    g_ClipAngelNameGui.Add("Text", "x12 y10 w390",
        "[Enter] use   [A] add   [E] edit   Delete   Esc keep temp name")
    g_ClipAngelNameLv := g_ClipAngelNameGui.Add("ListView", "x12 y40 w390 h320 Grid -Multi", ["Name"])
    g_ClipAngelNameLv.OnEvent("DoubleClick", (*) => ClipAngelExport_UseSelected())
    g_ClipAngelNameGui.OnEvent("Close", (*) => ClipAngelExport_Cancel())
    g_ClipAngelNameGui.OnEvent("Escape", (*) => ClipAngelExport_Cancel())
    ClipAngelExport_Refresh()
    ClipAngelExport_BindHotkeys([
        ["Enter", (*) => ClipAngelExport_UseSelected()],
        ["a", (*) => ClipAngelExport_Add()],
        ["Insert", (*) => ClipAngelExport_Add()],
        ["e", (*) => ClipAngelExport_Edit()],
        ["Delete", (*) => ClipAngelExport_Delete()],
        ["Escape", (*) => ClipAngelExport_Cancel()]
    ])
    ClipAngelExport_CenterGui(g_ClipAngelNameGui, 420, 400)
    try WinWaitClose("ahk_id " g_ClipAngelNameGui.Hwnd)
    catch {
    }
    ClipAngelExport_UnbindHotkeys()
    return g_ClipAngelNameFinalPath
}

ClipAngel_ExportLastClipToDesktop() {
    if !ClipAngel_TryAcquireAutomationLock()
        return
    savedClip := ClipboardAll()
    try {
        StandardLoadingBar_Show("⏳ Clip Angel: exporting...", BANNER_ACCENT_INTERMEDIATE)
        ClipAngel_ActivateNativeFirstClip()
        if !ClipAngel_WaitForListReady(CLIPANGEL_FAVORITE_OPEN_READY_MS, true) {
            ShowCenteredOverlay_Utils("❌ Clip Angel did not open.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        clipHwnd := ClipAngel_MainHwnd()
        if (clipHwnd)
            ClipAngel_LeaveFavoritesFilter(clipHwnd)
        ClipAngel_SelectClipCopyThenMinimize(0)
        content := A_Clipboard
        if (Trim(content) = "") {
            ShowCenteredOverlay_Utils("❌ Clip empty or not text.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        desktopPath := AiQuickDownload_ResolveDesktopPath()
        if (!desktopPath || !DirExist(desktopPath)) {
            ShowCenteredOverlay_Utils("❌ Desktop folder not found.", 2500, BANNER_ACCENT_ERROR)
            return
        }
        outPath := ClipAngelExport_UniqueDesktopPath(desktopPath)
        WriteUtf8File(outPath, content)
        try StandardLoadingBar_Hide(0)
        catch {
        }
        finalPath := ClipAngelExport_PromptRename(outPath)
        SplitPath(finalPath, &name)
        if (StrLen(name) > 48)
            name := SubStr(name, 1, 45) "..."
        ShowCenteredOverlay_Utils("✅ Saved: " name, 2200, BANNER_ACCENT_SUCCESS)
        try ScriptSoundPlay(A_ScriptDir . "\assets\sounds\copy.wav")
    } catch Error as e {
        ShowCenteredOverlay_Utils("❌ Clip Angel export failed: " . e.Message, 2500, BANNER_ACCENT_ERROR)
    } finally {
        try A_Clipboard := savedClip
        catch {
        }
        try StandardLoadingBar_Hide(0)
        catch {
        }
        ClipAngel_ReleaseAutomationLock()
    }
}

RegisterMacro(ClipAngel_ExportLastClipToDesktop, "📎 ClipAngel last clip → Desktop", "c")