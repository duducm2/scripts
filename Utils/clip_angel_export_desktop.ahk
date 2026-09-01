; =============================================================================
; Utils module: clip_angel_export_desktop.ahk
; Copy clipboard or Clip Angel Row 0 to Desktop, then prompt to rename from a persisted name list.
; Also hosts HotkeyCopy_ShowPostCopyBanner (#!+p 1×/2× post-copy destination menu).
; Utility Shortcuts: #!+U → Macros → [c]
; =============================================================================

global CLIPANGEL_EXPORT_CLIPBOARD_FIRST := true

global g_ClipAngelNameGui := false
global g_ClipAngelNameLv := false
global g_ClipAngelNameHint := false
global g_ClipAngelNameRows := []
global g_ClipAngelNameHotkeys := []
global g_ClipAngelNameSourcePath := ""
global g_ClipAngelNameFinalPath := ""
global g_ClipAngelNamePicked := false
global g_ClipAngelNameExt := "txt"
global g_ClipAngelNameOrigExt := ""
global g_ClipAngelNameOnClose := unset

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

ClipAngelExport_SanitizeExt(ext) {
    s := Trim(ext)
    while (SubStr(s, 1, 1) = ".")
        s := SubStr(s, 2)
    s := ClipAngelExport_SanitizeFileName(s)
    s := StrReplace(s, " ", "")
    if (s = "")
        return ""
    return StrLower(s)
}

; Split "name.ext" into base + ext. Returns true if an extension was found.
ClipAngelExport_SplitNameExt(raw, &baseOut, &extOut) {
    s := Trim(raw)
    baseOut := s
    extOut := ""
    dot := InStr(s, ".", false, -1)
    if (!dot || dot = 1 || dot = StrLen(s))
        return false
    baseCand := SubStr(s, 1, dot - 1)
    extCand := ClipAngelExport_SanitizeExt(SubStr(s, dot + 1))
    if (extCand = "" || ClipAngelExport_SanitizeFileName(baseCand) = "")
        return false
    baseOut := baseCand
    extOut := extCand
    return true
}

; Extension from source path only (no content guessing).
ClipAngelExport_ExtFromPath(path) {
    if (path = "")
        return ""
    srcExt := ""
    try SplitPath(path, , , &srcExt)
    catch {
        srcExt := ""
    }
    return ClipAngelExport_SanitizeExt(srcExt)
}

ClipAngelExport_FormatExtLabel(ext) {
    return ext != "" ? "." . ext : "(no ext)"
}

; Base name for list picks: strip a trailing .ext from registry entries (e.g. PALACE_QUICK_IMAGE.png).
ClipAngelExport_ListBaseName(name) {
    base := name
    ext := ""
    if ClipAngelExport_SplitNameExt(name, &base, &ext)
        return base
    return name
}

ClipAngelExport_UniqueNamedPath(desktopDir, baseName, ext) {
    if (ext = "")
        path := desktopDir "\" baseName
    else
        path := desktopDir "\" baseName "." ext
    if !FileExist(path)
        return path
    i := 2
    loop {
        if (ext = "")
            path := desktopDir "\" baseName "-" i
        else
            path := desktopDir "\" baseName "-" i "." ext
        if !FileExist(path)
            return path
        i += 1
    }
}

ClipAngelExport_UniqueStagingPath(desktopDir, ext) {
    ext := ClipAngelExport_SanitizeExt(ext)
    if (ext = "")
        ext := "txt"
    stamp := FormatTime(, "yyyyMMdd-HHmmss")
    base := "clipangel-last-" stamp
    path := desktopDir "\" base "." ext
    if !FileExist(path)
        return path
    i := 2
    loop {
        path := desktopDir "\" base "-" i "." ext
        if !FileExist(path)
            return path
        i += 1
    }
}

; Write clipboard to Desktop (image / file drop / text). No ClipAngel UI. Returns path or "".
ClipAngelExport_SaveClipboardToDesktop(&errMsg := "") {
    errMsg := ""
    desktopPath := AiQuickDownload_ResolveDesktopPath()
    if (desktopPath = "" || !DirExist(desktopPath)) {
        errMsg := "Desktop folder not found"
        return ""
    }
    savedClip := ClipboardAll()
    try {
        isImage := false
        try isImage := Task_ClipboardHasImage()
        catch {
            isImage := false
        }
        if (isImage) {
            outPath := ClipAngelExport_UniqueStagingPath(desktopPath, "png")
            if (!Task_SaveClipboardImage(outPath)) {
                errMsg := "Could not save image to Desktop"
                return ""
            }
            return outPath
        }
        if Clipboard_HasFileDrop() {
            paths := Clipboard_GetFilePaths()
            if (!paths || paths.Length < 1) {
                errMsg := "No file in clipboard"
                return ""
            }
            src := paths[1]
            if (src = "" || !FileExist(src)) {
                errMsg := "Clipboard file not found"
                return ""
            }
            dest := DesktopCutNewest_UniqueDestPath(desktopPath, src)
            if (dest = "") {
                errMsg := "Could not build Desktop path"
                return ""
            }
            try {
                if DirExist(src)
                    DirCopy(src, dest)
                else
                    FileCopy(src, dest)
            } catch {
                errMsg := "Failed to copy file to Desktop"
                return ""
            }
            if !FileExist(dest) {
                errMsg := "Copy verify failed"
                return ""
            }
            return dest
        }
        content := Trim(A_Clipboard)
        if (content != "") {
            outPath := ClipAngelExport_UniqueStagingPath(desktopPath, "txt")
            WriteUtf8File(outPath, content)
            if !FileExist(outPath) {
                errMsg := "Could not write text to Desktop"
                return ""
            }
            return outPath
        }
        errMsg := "Clipboard empty"
        return ""
    } finally {
        try A_Clipboard := savedClip
        catch {
        }
    }
}

ClipAngelExport_FindDesktopExplorerHwnd() {
    hwnd := WinExist("Área de Trabalho ahk_class CabinetWClass")
    if (hwnd)
        return hwnd
    return WinExist("Desktop ahk_class CabinetWClass")
}

; Activate Explorer showing Desktop (same target as #!+7 hold with Desktop focused).
ClipAngelExport_ActivateDesktopForPaste() {
    hwnd := ClipAngelExport_FindDesktopExplorerHwnd()
    if (hwnd) {
        try {
            if (WinGetMinMax("ahk_id " hwnd) = -1)
                WinRestore("ahk_id " hwnd)
            WinActivate("ahk_id " hwnd)
            if WinWaitActive("ahk_id " hwnd, , 1)
                return hwnd
        } catch {
        }
    }
    desktopPath := AiQuickDownload_ResolveDesktopPath()
    if (desktopPath = "" || !DirExist(desktopPath))
        return 0
    try Run('explorer.exe "' desktopPath '"')
    catch {
        return 0
    }
    deadline := A_TickCount + 3000
    while (A_TickCount < deadline) {
        hwnd := ClipAngelExport_FindDesktopExplorerHwnd()
        if (hwnd) {
            try {
                WinActivate("ahk_id " hwnd)
                if WinWaitActive("ahk_id " hwnd, , 1)
                    return hwnd
            } catch {
            }
        }
        Sleep 50
    }
    return 0
}

; ClipAngel Clip > Paste > Paste file onto Desktop; returns path or "".
ClipAngelExport_PasteFirstClipToDesktop() {
    desktopPath := AiQuickDownload_ResolveDesktopPath()
    if (desktopPath = "" || !DirExist(desktopPath))
        return ""
    beforePath := ""
    beforeStamp := ""
    try {
        beforePath := DesktopCutNewest_ResolveNewestPath(desktopPath)
        if (beforePath != "")
            beforeStamp := AiQuickDownload_ItemStamp(beforePath)
    } catch {
    }
    desktopHwnd := ClipAngelExport_ActivateDesktopForPaste()
    if (!desktopHwnd)
        return ""
    ClipAngel_ActivateNativeFirstClip(desktopHwnd)
    hwnd := ClipAngel_MainHwnd()
    if (!hwnd)
        return ""
    ClipAngel_EnsureWindowActive(hwnd)
    if !ClipAngel_WaitForListReady(CLIPANGEL_FAVORITE_OPEN_READY_MS, true)
        return ""
    if !ClipAngel_InvokePasteEnter(hwnd)
        return ""
    ClipAngel_CloseAndRestoreFocus(desktopHwnd)
    outPath := ""
    try outPath := AiQuickDownload_WaitForNewDesktopFile(desktopPath, beforePath, beforeStamp)
    catch {
        outPath := ""
    }
    return (outPath != "" && FileExist(outPath)) ? outPath : ""
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
    global g_ClipAngelNameGui, g_ClipAngelNameLv, g_ClipAngelNameHint, g_ClipAngelNameRows
    ClipAngelExport_UnbindHotkeys()
    try {
        if (IsObject(g_ClipAngelNameGui))
            g_ClipAngelNameGui.Destroy()
    } catch {
    }
    g_ClipAngelNameGui := false
    g_ClipAngelNameLv := false
    g_ClipAngelNameHint := false
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

ClipAngelExport_ApplyName(name, extOverride := unset) {
    global g_ClipAngelNameSourcePath, g_ClipAngelNameFinalPath, g_ClipAngelNamePicked, g_ClipAngelNameExt
    clean := ClipAngelExport_SanitizeFileName(name)
    if (clean = "") {
        ClipAngelExport_Alert("Name is not a valid filename.")
        return false
    }
    if (IsSet(extOverride)) {
        ext := extOverride != "" ? ClipAngelExport_SanitizeExt(extOverride) : ""
    } else {
        ext := g_ClipAngelNameExt
    }
    SplitPath(g_ClipAngelNameSourcePath, , &desktopDir)
    dest := ClipAngelExport_UniqueNamedPath(desktopDir, clean, ext)
    if (dest = g_ClipAngelNameSourcePath) {
        g_ClipAngelNameFinalPath := g_ClipAngelNameSourcePath
        g_ClipAngelNamePicked := true
        ClipAngelExport_CloseGui()
        return true
    }
    try {
        FileMove(g_ClipAngelNameSourcePath, dest)
        g_ClipAngelNameFinalPath := dest
        g_ClipAngelNamePicked := true
        ClipAngelExport_CloseGui()
        return true
    } catch Error as e {
        ClipAngelExport_Alert("Could not rename: " . e.Message)
        return false
    }
}

ClipAngelExport_UseSelected(*) {
    sel := ClipAngelExport_Selected()
    if (!sel) {
        try ShowCenteredOverlay_Utils("Select a name", 1200, BANNER_ACCENT_ERROR)
        catch {
        }
        return
    }
    ClipAngelExport_ApplyName(ClipAngelExport_ListBaseName(sel["name"]), ClipAngelExport_OrigExt())
}

ClipAngelExport_UseSelectedTxt(*) {
    sel := ClipAngelExport_Selected()
    if (!sel) {
        try ShowCenteredOverlay_Utils("Select a name", 1200, BANNER_ACCENT_ERROR)
        catch {
        }
        return
    }
    ClipAngelExport_ApplyName(ClipAngelExport_ListBaseName(sel["name"]), "txt")
}

ClipAngelExport_OrigExt() {
    global g_ClipAngelNameOrigExt, g_ClipAngelNameSourcePath
    if (g_ClipAngelNameOrigExt != "")
        return g_ClipAngelNameOrigExt
    if (g_ClipAngelNameSourcePath != "")
        return ClipAngelExport_ExtFromPath(g_ClipAngelNameSourcePath)
    return ""
}

; One-shot typed name for this file only (not saved to the CSV list).
; Accepts "name" (uses current ext) or "name.ext" (overrides ext for this file).
ClipAngelExport_UseTyped(*) {
    global g_ClipAngelNameExt
    extLbl := ClipAngelExport_FormatExtLabel(g_ClipAngelNameExt)
    ClipAngelExport_DialogsBegin()
    ib := InputBox("Filename (optional .ext). Not saved to list.`nCurrent ext: " . extLbl,
        "Type file name", "w360", "")
    ClipAngelExport_DialogsEnd()
    if (ib.Result != "OK")
        return
    raw := Trim(ib.Value)
    if (raw = "") {
        ClipAngelExport_Alert("Name is required.")
        return
    }
    base := raw
    ext := ""
    if ClipAngelExport_SplitNameExt(raw, &base, &ext)
        ClipAngelExport_ApplyName(base, ext)
    else
        ClipAngelExport_ApplyName(raw)
}

ClipAngelExport_UpdateHint() {
    global g_ClipAngelNameHint, g_ClipAngelNameExt
    if (!IsObject(g_ClipAngelNameHint))
        return
    origLbl := ClipAngelExport_FormatExtLabel(ClipAngelExport_OrigExt())
    extLbl := ClipAngelExport_FormatExtLabel(g_ClipAngelNameExt)
    g_ClipAngelNameHint.Value := "[Enter] keep ext " . origLbl . "   [1] .txt   [T] type once   [X] ext "
        . extLbl . "   [A] add   [E] edit   Delete   Esc keep temp"
}

; Change extension for this rename session (list picks + bare typed names).
ClipAngelExport_SetExt(*) {
    global g_ClipAngelNameExt
    ClipAngelExport_DialogsBegin()
    ib := InputBox("Extension without dot (e.g. md, csv, json)", "Set file extension",
        "w320", g_ClipAngelNameExt)
    ClipAngelExport_DialogsEnd()
    if (ib.Result != "OK")
        return
    ext := ClipAngelExport_SanitizeExt(ib.Value)
    if (ext = "") {
        ClipAngelExport_Alert("Extension is required.")
        return
    }
    g_ClipAngelNameExt := ext
    ClipAngelExport_UpdateHint()
    try ShowCenteredOverlay_Utils("ℹ Extension: ." . ext, 1200, BANNER_ACCENT_INFO)
    catch {
    }
}

ClipAngelExport_Cancel(*) {
    ClipAngelExport_CloseGui()
}

ClipAngelExport_CopySelected(*) {
    sel := ClipAngelExport_Selected()
    if (!sel) {
        try ShowCenteredOverlay_Utils("Select a name", 1200, BANNER_ACCENT_ERROR)
        catch {
        }
        return
    }
    A_Clipboard := sel["name"]
    try ShowCenteredOverlay_Utils("Copied: " . sel["name"], 1200, BANNER_ACCENT_SUCCESS)
    catch {
    }
}

ClipAngelExport_ManagerClose(*) {
    global g_ClipAngelNameOnClose
    cb := g_ClipAngelNameOnClose
    g_ClipAngelNameOnClose := unset
    ClipAngelExport_CloseGui()
    if (IsSet(cb) && cb) {
        try cb.Call()
        catch {
        }
    }
}

ClipAngelExport_UpdateManagerHint() {
    global g_ClipAngelNameHint
    if (!IsObject(g_ClipAngelNameHint))
        return
    g_ClipAngelNameHint.Value := "[Enter] / [C] copy name   [A] add   [E] edit   Delete   Esc back"
}

; Standalone CRUD + copy for clipangel_desktop_names.csv (Import Management [N]; same list as #!+9 rename picker).
ClipAngelExport_ShowNamesManager(onClose := unset) {
    global g_ClipAngelNameGui, g_ClipAngelNameLv, g_ClipAngelNameHint, g_ClipAngelNameOnClose
    if (IsObject(g_ClipAngelNameGui)) {
        try WinActivate("ahk_id " g_ClipAngelNameGui.Hwnd)
        catch {
        }
        return
    }
    ClipAngelExport_CloseGui()
    g_ClipAngelNameOnClose := onClose
    g_ClipAngelNameGui := Gui("+AlwaysOnTop +ToolWindow", "Quick Download — file names")
    g_ClipAngelNameGui.SetFont("s10", "Segoe UI")
    g_ClipAngelNameHint := g_ClipAngelNameGui.Add("Text", "x12 y10 w390 h36")
    ClipAngelExport_UpdateManagerHint()
    g_ClipAngelNameLv := g_ClipAngelNameGui.Add("ListView", "x12 y50 w390 h310 Grid -Multi", ["Name"])
    g_ClipAngelNameLv.OnEvent("DoubleClick", (*) => ClipAngelExport_CopySelected())
    g_ClipAngelNameGui.OnEvent("Close", (*) => ClipAngelExport_ManagerClose())
    g_ClipAngelNameGui.OnEvent("Escape", (*) => ClipAngelExport_ManagerClose())
    ClipAngelExport_Refresh()
    ClipAngelExport_BindHotkeys([
        ["Enter", (*) => ClipAngelExport_CopySelected()],
        ["c", (*) => ClipAngelExport_CopySelected()],
        ["a", (*) => ClipAngelExport_Add()],
        ["Insert", (*) => ClipAngelExport_Add()],
        ["e", (*) => ClipAngelExport_Edit()],
        ["Delete", (*) => ClipAngelExport_Delete()],
        ["Backspace", (*) => ClipAngelExport_ManagerClose()],
        ["Escape", (*) => ClipAngelExport_ManagerClose()]
    ])
    ClipAngelExport_CenterGui(g_ClipAngelNameGui, 420, 410)
}

; Shows rename picker. Returns final path (renamed or original if Esc/close).
; Original extension: from source path only (Enter keeps it; 1 forces .txt).
ClipAngelExport_PromptRename(sourcePath) {
    global g_ClipAngelNameGui, g_ClipAngelNameLv, g_ClipAngelNameHint, g_ClipAngelNameSourcePath
    global g_ClipAngelNameFinalPath, g_ClipAngelNamePicked, g_ClipAngelNameExt, g_ClipAngelNameOrigExt
    g_ClipAngelNameSourcePath := sourcePath
    g_ClipAngelNameFinalPath := sourcePath
    g_ClipAngelNamePicked := false
    g_ClipAngelNameOrigExt := ClipAngelExport_ExtFromPath(sourcePath)
    g_ClipAngelNameExt := g_ClipAngelNameOrigExt

    ClipAngelExport_CloseGui()
    g_ClipAngelNameGui := Gui("+AlwaysOnTop +ToolWindow", "Name Desktop file")
    g_ClipAngelNameGui.SetFont("s10", "Segoe UI")
    g_ClipAngelNameHint := g_ClipAngelNameGui.Add("Text", "x12 y10 w390 h48")
    ClipAngelExport_UpdateHint()
    g_ClipAngelNameLv := g_ClipAngelNameGui.Add("ListView", "x12 y62 w390 h298 Grid -Multi", ["Name"])
    g_ClipAngelNameLv.OnEvent("DoubleClick", (*) => ClipAngelExport_UseSelected())
    g_ClipAngelNameGui.OnEvent("Close", (*) => ClipAngelExport_Cancel())
    g_ClipAngelNameGui.OnEvent("Escape", (*) => ClipAngelExport_Cancel())
    ClipAngelExport_Refresh()
    ClipAngelExport_BindHotkeys([
        ["Enter", (*) => ClipAngelExport_UseSelected()],
        ["1", (*) => ClipAngelExport_UseSelectedTxt()],
        ["t", (*) => ClipAngelExport_UseTyped()],
        ["x", (*) => ClipAngelExport_SetExt()],
        ["a", (*) => ClipAngelExport_Add()],
        ["Insert", (*) => ClipAngelExport_Add()],
        ["e", (*) => ClipAngelExport_Edit()],
        ["Delete", (*) => ClipAngelExport_Delete()],
        ["Escape", (*) => ClipAngelExport_Cancel()]
    ])
    ClipAngelExport_CenterGui(g_ClipAngelNameGui, 420, 410)
    try WinWaitClose("ahk_id " g_ClipAngelNameGui.Hwnd)
    catch {
    }
    ClipAngelExport_UnbindHotkeys()
    return g_ClipAngelNameFinalPath
}

; Post-#!+p banner context: origin window + companion for Transfer / Read / Paste / Clip Angel.
global g_HotkeyCopy_PostCopyContext := { originHwnd: 0, isCode: false, companion: "" }

HotkeyCopy_ClosePostCopyBanner() {
    try StandardLoadingBar_CloseKeysOverlay()
    catch {
    }
    try StandardLoadingBar_Hide(0)
    catch {
    }
}

ClipAngelExport_OnCancelDesktop(*) {
    HotkeyCopy_ClosePostCopyBanner()
}

ClipAngelExport_OnConfirmDesktop(*) {
    HotkeyCopy_ClosePostCopyBanner()
    Sleep CLIPANGEL_PRE_FAVORITE_INGEST_DELAY_MS
    ClipAngel_ExportLastClipToDesktop()
}

ClipAngelExport_OnFavoriteClip(*) {
    HotkeyCopy_ClosePostCopyBanner()
    clip := Trim(A_Clipboard)
    if (clip = "" || StrLen(clip) < 10) {
        ShowCenteredOverlay_Utils("❌ Nothing to favorite - clipboard empty or too short", 2000, BANNER_ACCENT_ERROR)
        return
    }
    MarkLastClipAsFavorite("first", true)
}

; [C] Transfer clipboard to a Cursor/VS Code window (same as D2C Copy response? C).
HotkeyCopy_OnTransfer(*) {
    global g_HotkeyCopy_PostCopyContext
    HotkeyCopy_ClosePostCopyBanner()
    originHwnd := g_HotkeyCopy_PostCopyContext.originHwnd
    clipRaw := A_Clipboard
    clip := Trim(clipRaw)
    if (clip = "" || StrLen(clip) < 10) {
        ShowCenteredOverlay_Utils("❌ Clipboard empty or too short", 2000, BANNER_ACCENT_ERROR)
        return
    }
    targetHwnd := CursorTransfer_ShowWindowSelector(0)
    if (!targetHwnd) {
        if (originHwnd && WinExist("ahk_id " originHwnd))
            WinActivate("ahk_id " originHwnd)
        try A_Clipboard := clipRaw
        return
    }
    try A_Clipboard := clipRaw
    CursorTransfer_ActivateFocusPaste(targetHwnd, originHwnd)
}

; [R] Read aloud already-copied message (1× only; skip for code / Enterprise).
; Uses Gemini.ahk IPC (same as D2C DoCopyCore) so Utils need not call Gemini-only functions.
HotkeyCopy_OnRead(*) {
    global g_HotkeyCopy_PostCopyContext
    HotkeyCopy_ClosePostCopyBanner()
    companion := g_HotkeyCopy_PostCopyContext.companion
    if (companion = "enterprise") {
        ShowCenteredOverlay_Utils("❌ Read aloud not supported for Gemini Enterprise", 2500, BANNER_ACCENT_ERROR)
        return
    }
    originHwnd := g_HotkeyCopy_PostCopyContext.originHwnd
    WM_TRIGGER_READ_ALOUD := 0x8004
    WM_TRIGGER_COPILOT_READ_ALOUD := 0x8006
    wmRead := (companion = "copilot") ? WM_TRIGGER_COPILOT_READ_ALOUD : WM_TRIGGER_READ_ALOUD
    targetHwnd := GetGeminiScriptMsgTargetHwnd()
    if (!targetHwnd) {
        ShowCenteredOverlay_Utils("❌ Gemini.ahk not running", 2000, BANNER_ACCENT_ERROR)
        return
    }
    prevDH := A_DetectHiddenWindows
    DetectHiddenWindows true
    try {
        ; wParam=1: already copied; skip internal Copy. lParam: origin for focus restore.
        SendMessage(wmRead, 1, originHwnd, , "ahk_id " targetHwnd, , , , 120000)
    } catch {
        ShowCenteredOverlay_Utils("❌ Read aloud failed or timed out", 4000, BANNER_ACCENT_ERROR)
    } finally {
        DetectHiddenWindows prevDH
    }
}

; [W] Paste clipboard to a picked visible window (same as D2C Send dictation? W / #!+L).
HotkeyCopy_OnPasteWindow(*) {
    global g_HotkeyCopy_PostCopyContext
    HotkeyCopy_ClosePostCopyBanner()
    originHwnd := g_HotkeyCopy_PostCopyContext.originHwnd
    D2C_FlowManager.GetInstance().PasteClipboardToVisibleWindow(originHwnd)
}

; [O] Open Clip Angel on newest clip and Edit text (F4) — same path as D2C Send dictation? O.
HotkeyCopy_OnClipAngelEdit(*) {
    global g_HotkeyCopy_PostCopyContext
    HotkeyCopy_ClosePostCopyBanner()
    if !ClipAngel_TryAcquireAutomationLock()
        return

    StandardLoadingBar_Show("⏳ Clip Angel: opening...", BANNER_ACCENT_INTERMEDIATE)
    try {
        originHwnd := g_HotkeyCopy_PostCopyContext.originHwnd
        if (!originHwnd)
            try originHwnd := WinGetID("A")
        originMon := GetAhkMonitorIndexFromHwnd(originHwnd)

        ActivateClipAngelWithFocusCorrection(true, originMon, true)
        clipHwnd := ClipAngel_MainHwnd()
        if (!clipHwnd) {
            StandardLoadingBar_Update("❌ Clip Angel: window not found", BANNER_ACCENT_ERROR)
            return
        }

        if (!WinWaitActive("ahk_id " clipHwnd, , 0.6)) {
            StandardLoadingBar_Update("❌ Clip Angel: failed to activate", BANNER_ACCENT_ERROR)
            return
        }

        StandardLoadingBar_Update("⏳ Clip Angel: opening editor...", BANNER_ACCENT_INTERMEDIATE)
        ClipAngel_LeaveFavoritesFilter(clipHwnd)
        priorSendLevel := A_SendLevel
        SendLevel 0
        SendInput "{Tab}"
        Sleep 40
        SendInput "^a"
        Sleep 40
        SendInput "^c"
        try ClipWait(0.3)
        catch {
        }
        SendInput "{F10}"
        ClipAngel_UiaWaitPreviewFocused(clipHwnd, 150)
        SendInput "{Up}"
        Sleep 40
        SendInput "{F4}"
        SendLevel priorSendLevel

        StandardLoadingBar_Update("⏳ Clip Angel: maximizing...", BANNER_ACCENT_INTERMEDIATE)
        TryMaximizeWindow(clipHwnd)
        StandardLoadingBar_Update("✅ Clip Angel: ready", BANNER_ACCENT_SUCCESS)
    } finally {
        StandardLoadingBar_Hide(350)
        ClipAngel_ReleaseAutomationLock()
    }
}

; After #!+p 1×/2× successful copy: 5s destination banner (Desktop / Favorite / Transfer / …).
; isCode: true after double-tap code copy (omits Read aloud).
HotkeyCopy_ShowPostCopyBanner(isCode := false, originHwnd := 0) {
    global g_HotkeyCopy_PostCopyContext
    companion := ""
    try companion := ResolveGlobalAICompanion()
    catch {
        companion := ""
    }
    g_HotkeyCopy_PostCopyContext := { originHwnd: originHwnd, isCode: isCode, companion: companion }

    keyCallbacks := Map(
        "Y", ClipAngelExport_OnConfirmDesktop,
        "F", ClipAngelExport_OnFavoriteClip,
        "C", HotkeyCopy_OnTransfer,
        "W", HotkeyCopy_OnPasteWindow,
        "O", HotkeyCopy_OnClipAngelEdit,
        "N", ClipAngelExport_OnCancelDesktop,
        "Escape", ClipAngelExport_OnCancelDesktop)

    ; Read aloud only for full message (1×); Enterprise has no read-aloud path yet.
    if (!isCode && companion != "enterprise")
        keyCallbacks["R"] := HotkeyCopy_OnRead

    if (isCode) {
        title := "❓ Copied code — what next? (5s)"
        pk := "[Y] Desktop  [F] Favorite  [C] Transfer  [W] Paste window  [O] Clip Angel  [N] No"
    } else if (companion = "enterprise") {
        title := "❓ Copied message — what next? (5s)"
        pk := "[Y] Desktop  [F] Favorite  [C] Transfer  [W] Paste window  [O] Clip Angel  [N] No"
    } else {
        title := "❓ Copied message — what next? (5s)"
        pk := "[Y] Desktop  [F] Favorite  [C] Transfer  [R] Read  [W] Paste window  [O] Clip Angel  [N] No"
    }

    timeoutMs := 5000
    try timeoutMs := D2C_SUBMIT_MENU_TIMEOUT_MS
    catch {
        timeoutMs := 5000
    }

    StandardLoadingBar_ShowWithKeys(
        title,
        keyCallbacks,
        timeoutMs,
        0,
        ClipAngelExport_OnCancelDesktop,
        BANNER_ACCENT_INTERMEDIATE,
        900,
        17,
        "",
        false,
        pk,
        true,
        true)
}

; Back-compat aliases.
ClipAngelExport_PromptAfterHotkeyCopy() {
    HotkeyCopy_ShowPostCopyBanner(false)
}

ClipAngelExport_PromptAfterCodeCopy() {
    HotkeyCopy_ShowPostCopyBanner(true)
}

ClipAngel_ExportLastClipToDesktop() {
    if !ClipAngel_TryAcquireAutomationLock()
        return
    savedClip := ClipboardAll()
    try {
        StandardLoadingBar_Show("⏳ Clip Angel: exporting...", BANNER_ACCENT_INTERMEDIATE)
        outPath := ""
        errMsg := ""
        clipboardFirst := true
        try clipboardFirst := CLIPANGEL_EXPORT_CLIPBOARD_FIRST
        catch {
            clipboardFirst := true
        }
        if (clipboardFirst)
            outPath := ClipAngelExport_SaveClipboardToDesktop(&errMsg)
        if (outPath = "")
            outPath := ClipAngelExport_PasteFirstClipToDesktop()
        if (outPath = "") {
            msg := errMsg != "" ? errMsg : "Paste file to Desktop failed"
            ShowCenteredOverlay_Utils("❌ " . msg, 2500, BANNER_ACCENT_ERROR)
            return
        }
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