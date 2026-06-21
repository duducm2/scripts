; =============================================================================
; Utils module: context_file_browser.ahk
; Context file browser (Win+Alt+Shift+N)
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; Context file browser — browse context/ and paste full local paths (Win+Alt+Shift+N)
; =============================================================================
global CONTEXT_PREVIEW_MAX_SIZE := 234  ; 180 * 1.3
global g_ContextBrowserLastDir := ""
global g_ContextBrowserPreviewHbm := 0
global g_ContextBrowserFilterQuery := ""
global g_ContextBrowserFileIndex := []
global g_ContextBrowserFilterTyping := false
global g_ContextBrowserSuppressFilterKillFocus := false
global g_ContextBrowserBreadcrumbLink := false
global g_ContextBrowserBreadcrumbSegments := []
global g_ContextBrowserEntryPathLabel := false

ContextBrowser_IsFilterFocused() {
    global g_ContextBrowserFilterCtrl, g_ContextBrowserGui, g_ContextBrowserFilterTyping
    if (g_ContextBrowserFilterTyping)
        return true
    if (!IsObject(g_ContextBrowserFilterCtrl) || !IsObject(g_ContextBrowserGui))
        return false
    try {
        focusedHwnd := ControlGetFocus(, g_ContextBrowserGui)
    } catch {
        return false
    }
    return (focusedHwnd = g_ContextBrowserFilterCtrl.Hwnd)
}

ContextBrowser_FocusFilterField() {
    global g_ContextBrowserFilterCtrl, g_ContextBrowserActive, g_ContextBrowserFilterTyping
    if (!IsObject(g_ContextBrowserFilterCtrl))
        return
    g_ContextBrowserFilterTyping := true
    if (g_ContextBrowserActive)
        ContextBrowser_SetListNavigationHotkeysEnabled(false)
    try g_ContextBrowserFilterCtrl.Focus()
    catch {
    }
    try SendMessage(0x00B1, 0, -1, g_ContextBrowserFilterCtrl)
    catch {
    }
}

ContextBrowser_EnsureGlobals() {
    global CONTEXT_ROOT, g_ContextBrowserActive, g_ContextBrowserGui, g_ContextBrowserCurrentDir
    global g_ContextBrowserEntries, g_ContextBrowserListView, g_ContextBrowserBreadcrumbLink,
        g_ContextBrowserLetterHook
    global g_ContextBrowserPreviewCtrl, g_ContextBrowserPreviewText, g_ContextBrowserFilterCtrl
    global g_ContextBrowserLastDir, g_ContextBrowserPreviewHbm, g_ContextBrowserFilterQuery,
        g_ContextBrowserBreadcrumbSegments
    global g_ContextBrowserFileIndex, g_ContextBrowserFilterTyping, g_ContextBrowserSuppressFilterKillFocus
    global g_ContextBrowserEntryPathLabel
    if !IsSet(CONTEXT_ROOT)
        CONTEXT_ROOT := A_ScriptDir "\context"
    if !IsSet(g_ContextBrowserActive)
        g_ContextBrowserActive := false
    if !IsSet(g_ContextBrowserGui)
        g_ContextBrowserGui := false
    if !IsSet(g_ContextBrowserCurrentDir)
        g_ContextBrowserCurrentDir := ""
    if !IsSet(g_ContextBrowserEntries)
        g_ContextBrowserEntries := []
    if !IsSet(g_ContextBrowserListView)
        g_ContextBrowserListView := false
    if !IsSet(g_ContextBrowserBreadcrumbLink)
        g_ContextBrowserBreadcrumbLink := false
    if !IsSet(g_ContextBrowserBreadcrumbSegments)
        g_ContextBrowserBreadcrumbSegments := []
    if !IsSet(g_ContextBrowserPreviewCtrl)
        g_ContextBrowserPreviewCtrl := false
    if !IsSet(g_ContextBrowserPreviewText)
        g_ContextBrowserPreviewText := false
    if !IsSet(g_ContextBrowserFilterCtrl)
        g_ContextBrowserFilterCtrl := false
    if !IsSet(g_ContextBrowserLastDir)
        g_ContextBrowserLastDir := ""
    if !IsSet(g_ContextBrowserPreviewHbm)
        g_ContextBrowserPreviewHbm := 0
    if !IsSet(g_ContextBrowserFilterQuery)
        g_ContextBrowserFilterQuery := ""
    if !IsSet(g_ContextBrowserFileIndex)
        g_ContextBrowserFileIndex := []
    if !IsSet(g_ContextBrowserFilterTyping)
        g_ContextBrowserFilterTyping := false
    if !IsSet(g_ContextBrowserSuppressFilterKillFocus)
        g_ContextBrowserSuppressFilterKillFocus := false
    if !IsSet(g_ContextBrowserEntryPathLabel)
        g_ContextBrowserEntryPathLabel := false
    if !IsSet(g_ContextBrowserLetterHook)
        g_ContextBrowserLetterHook := ""
}
ContextBrowser_EnsureGlobals()

Context_GetRoot() {
    ContextBrowser_EnsureGlobals()
    global CONTEXT_ROOT
    return CONTEXT_ROOT
}

Context_IsAtRoot(dir) {
    root := Context_GetRoot()
    return (StrLower(RTrim(dir, "\")) = StrLower(RTrim(root, "\")))
}

Context_IsExistingFile(path) {
    if (path = "")
        return false
    attr := FileExist(path)
    return (attr && !InStr(attr, "D"))
}

Context_IsImageExtension(ext) {
    ext := StrLower(ext)
    static imageExts := Map("png", 1, "jpg", 1, "jpeg", 1, "gif", 1, "bmp", 1, "webp", 1, "ico", 1, "tif", 1,
        "tiff", 1)
    return imageExts.Has(ext)
}

Context_IsImagePath(path) {
    if !Context_IsExistingFile(path)
        return false
    SplitPath path, , , &ext
    return Context_IsImageExtension(ext)
}

ContextBrowser_GetPreviewImagePath(entry) {
    if (!IsObject(entry) || entry.type != "file")
        return ""
    if Context_IsImagePath(entry.path)
        return entry.path
    probe := Context_ProbeReference(entry.path)
    if (probe.isRef && Context_IsImagePath(probe.targetPath))
        return probe.targetPath
    return ""
}

ContextBrowser_GetFitImageSize(path, maxW, maxH, &outW, &outH) {
    outW := 0
    outH := 0
    if (path = "" || maxW < 1 || maxH < 1)
        return false
    try hbm := LoadPicture(path)
    catch
        return false
    if !hbm
        return false
    bm := Buffer(A_PtrSize = 8 ? 32 : 24, 0)
    if !DllCall("GetObject", "ptr", hbm, "int", bm.Size, "ptr", bm) {
        DllCall("DeleteObject", "ptr", hbm)
        return false
    }
    imgW := NumGet(bm, 4, "int")
    imgH := NumGet(bm, 8, "int")
    DllCall("DeleteObject", "ptr", hbm)
    if (imgW <= 0 || imgH <= 0)
        return false
    scale := Min(maxW / imgW, maxH / imgH, 1)
    outW := Max(1, Round(imgW * scale))
    outH := Max(1, Round(imgH * scale))
    return true
}

Context_ReadFirstNonemptyLine(path) {
    try {
        content := FileRead(path)
    } catch {
        return ""
    }
    for line in StrSplit(content, "`n", "`r") {
        trimmed := Trim(line)
        if (trimmed != "")
            return trimmed
    }
    return ""
}

Context_IsAbsoluteFilePath(line) {
    return (line != "" && RegExMatch(line, "^[A-Za-z]:\\"))
}

Context_ResolveReferenceLine(refPath, line) {
    line := Trim(line)
    if (line = "")
        return ""
    if Context_IsAbsoluteFilePath(line)
        return Context_IsExistingFile(line) ? line : ""
    SplitPath refPath, , &refDir
    root := Context_GetRoot()
    for candidate in [refDir "\" line, root "\" line] {
        if Context_IsExistingFile(candidate)
            return candidate
    }
    return ""
}

Context_ProbeReference(path) {
    result := { isRef: false, targetPath: "", targetBasename: "" }
    if !Context_IsExistingFile(path)
        return result
    SplitPath path, , , &ext
    ext := StrLower(ext)
    if (ext = "lnk") {
        try {
            target := ComObject("WScript.Shell").CreateShortcut(path).TargetPath
            if (target != "" && Context_IsExistingFile(target)) {
                result.isRef := true
                result.targetPath := target
                SplitPath target, &base
                result.targetBasename := base
            }
        } catch {
        }
        return result
    }
    line := Context_ReadFirstNonemptyLine(path)
    resolved := Context_ResolveReferenceLine(path, line)
    if (resolved != "") {
        result.isRef := true
        result.targetPath := resolved
        SplitPath resolved, &base
        result.targetBasename := base
    }
    return result
}

; Returns path to paste, or "" if a pointer/shortcut reference is broken (caller shows error).
Context_ResolvePastePath(path) {
    if (path = "")
        return ""
    SplitPath path, , , &ext
    ext := StrLower(ext)
    if (ext = "lnk") {
        try {
            target := ComObject("WScript.Shell").CreateShortcut(path).TargetPath
            if (target != "" && Context_IsExistingFile(target))
                return target
        } catch {
        }
        return ""
    }
    line := Context_ReadFirstNonemptyLine(path)
    resolved := Context_ResolveReferenceLine(path, line)
    if (resolved != "")
        return resolved
    return path
}

Context_ResolvePasteTextContent(path) {
    pastePath := Context_ResolvePastePath(path)
    if (pastePath = "" || !Context_IsExistingFile(pastePath))
        return ""
    if Context_IsImagePath(pastePath)
        return ""
    try {
        return FileRead(pastePath, "UTF-8")
    } catch {
        return ""
    }
}

Context_SortNames(names) {
    if (names.Length < 2)
        return names
    list := ""
    for n in names
        list .= n "`n"
    list := Sort(RTrim(list, "`n"), "P`n")
    sorted := []
    if (list != "") {
        for line in StrSplit(list, "`n")
            sorted.Push(line)
    }
    return sorted
}

Context_ListDirEntries(dir) {
    folders := []
    files := []
    try {
        loop files, dir "\*", "D" {
            if (A_LoopFileAttrib ~= "[HS]")
                continue
            folders.Push(A_LoopFileName)
        }
        loop files, dir "\*", "F" {
            if (A_LoopFileAttrib ~= "[HS]")
                continue
            files.Push(A_LoopFileName)
        }
    } catch {
    }
    folders := Context_SortNames(folders)
    files := Context_SortNames(files)
    entries := []
    for name in folders
        entries.Push({ type: "folder", name: name, path: dir "\" name })
    for name in files
        entries.Push({ type: "file", name: name, path: dir "\" name })
    return entries
}

ContextBrowser_BuildViewEntries(dir) {
    entries := []
    if !Context_IsAtRoot(dir)
        entries.Push({ type: "parent", name: "..", path: "" })
    for entry in Context_ListDirEntries(dir)
        entries.Push(entry)
    return entries
}

ContextBrowser_BuildFileIndex() {
    root := RTrim(Context_GetRoot(), "\")
    relPaths := []
    pathByRel := Map()
    try {
        loop files, root "\**", "R" {
            if (A_LoopFileAttrib ~= "[HS]")
                continue
            fullPath := A_LoopFileFullPath
            relPath := SubStr(fullPath, StrLen(root) + 2)
            SplitPath fullPath, &name
            relPaths.Push(relPath)
            pathByRel[relPath] := { type: "file", name: name, path: fullPath, relPath: relPath }
        }
    } catch {
    }
    index := []
    for relPath in Context_SortNames(relPaths)
        index.Push(pathByRel[relPath])
    return index
}

ContextBrowser_GetFileIndex() {
    global g_ContextBrowserFileIndex
    if (!IsObject(g_ContextBrowserFileIndex) || g_ContextBrowserFileIndex.Length = 0)
        g_ContextBrowserFileIndex := ContextBrowser_BuildFileIndex()
    return g_ContextBrowserFileIndex
}

ContextBrowser_SearchFileIndex(index, query) {
    q := Trim(query)
    if (q = "")
        return []
    qLower := StrLower(q)
    qPath := StrReplace(qLower, "/", "\")
    results := []
    for entry in index {
        relLower := StrLower(entry.relPath)
        if InStr(StrLower(entry.name), qLower) || InStr(relLower, qPath)
            results.Push(entry)
    }
    return results
}

ContextBrowser_ApplyNameFilter(entries) {
    global g_ContextBrowserFilterQuery
    q := Trim(g_ContextBrowserFilterQuery)
    if (q = "")
        return entries
    qLower := StrLower(q)
    filtered := []
    for entry in entries {
        if (entry.type = "parent") {
            filtered.Push(entry)
            continue
        }
        if InStr(StrLower(entry.name), qLower)
            filtered.Push(entry)
    }
    return filtered
}

ContextBrowser_FindLinkedResearchJson(imagePath) {
    root := RTrim(Context_GetRoot(), "\")
    pathNorm := RTrim(imagePath, "\")
    if (StrLen(pathNorm) <= StrLen(root) + 1)
        return ""
    if (StrLower(SubStr(pathNorm, 1, StrLen(root))) != StrLower(root))
        return ""
    rel := SubStr(pathNorm, StrLen(root) + 2)
    if !InStr(StrLower(rel), "image-references\")
        return ""
    relJson := StrReplace(rel, "image-references\", "research\", false)
    SplitPath relJson, , &dir, &nameNoExt, &ext
    if (ext != "")
        relJson := dir "\" nameNoExt ".json"
    candidate := root "\" relJson
    return Context_IsExistingFile(candidate) ? candidate : ""
}

ContextBrowser_GetJsonPreviewText(path) {
    SplitPath path, , , &ext
    if (StrLower(ext) != "json")
        return ""
    try content := FileRead(path)
    catch
        return ""
    if RegExMatch(content, '"_meta"\s*:\s*\{[^}]*"topic"\s*:\s*"([^"]+)"', &m)
        return m[1]
    if RegExMatch(content, '"metadata"\s*:\s*\{[^}]*"file_name"\s*:\s*"([^"]+)"', &m)
        return m[1]
    flat := RegExReplace(content, "\s+", " ")
    flat := Trim(flat)
    if (flat = "")
        return ""
    return SubStr(flat, 1, 280) . (StrLen(flat) > 280 ? "…" : "")
}

Context_GetRelativeSubtitle(dir) {
    root := Context_GetRoot()
    rootNorm := RTrim(root, "\")
    dirNorm := RTrim(dir, "\")
    if (StrLower(dirNorm) = StrLower(rootNorm))
        return rootNorm
    rel := SubStr(dirNorm, StrLen(rootNorm) + 2)
    return rootNorm "  »  " StrReplace(rel, "\", " » ")
}

ContextBrowser_GetBreadcrumbSegments(dir) {
    root := RTrim(Context_GetRoot(), "\")
    dirNorm := RTrim(dir, "\")
    segments := [{ label: "context", path: root }]
    if (StrLower(dirNorm) = StrLower(root))
        return segments
    rel := SubStr(dirNorm, StrLen(root) + 2)
    built := root
    for part in StrSplit(rel, "\") {
        if (part = "")
            continue
        built .= "\" part
        segments.Push({ label: part, path: built })
    }
    return segments
}

ContextBrowser_EscapeLinkLabel(label) {
    return StrReplace(label, "&", "&&")
}

ContextBrowser_OnBreadcrumbLinkClick(ctrl, id, href, *) {
    global g_ContextBrowserBreadcrumbSegments
    idx := Integer(id)
    if (idx < 1 || idx > g_ContextBrowserBreadcrumbSegments.Length)
        return
    ContextBrowser_BreadcrumbNavigate(g_ContextBrowserBreadcrumbSegments[idx].path)
}

ContextBrowser_BreadcrumbNavigate(targetDir, *) {
    ContextBrowser_EnsureGlobals()
    global g_ContextBrowserActive, g_ContextBrowserCurrentDir
    if (!g_ContextBrowserActive || targetDir = "" || !DirExist(targetDir))
        return
    g_ContextBrowserCurrentDir := targetDir
    ContextBrowser_RefreshView()
}

ContextBrowser_UpdateBreadcrumbs(dir) {
    global g_ContextBrowserBreadcrumbLink, g_ContextBrowserBreadcrumbSegments
    if (!IsObject(g_ContextBrowserBreadcrumbLink))
        return
    g_ContextBrowserBreadcrumbSegments := ContextBrowser_GetBreadcrumbSegments(dir)
    segments := g_ContextBrowserBreadcrumbSegments
    linkText := ""
    for i, seg in segments {
        if (i > 1)
            linkText .= " » "
        label := ContextBrowser_EscapeLinkLabel(seg.label)
        if (i = segments.Length)
            linkText .= label
        else
            linkText .= Format("<a id=`"{1}`">{2}</a>", i, label)
    }
    g_ContextBrowserBreadcrumbLink.Text := linkText
}

ContextBrowser_UpdateSearchBreadcrumbs(query, count) {
    global g_ContextBrowserBreadcrumbLink, g_ContextBrowserBreadcrumbSegments
    if (!IsObject(g_ContextBrowserBreadcrumbLink))
        return
    root := RTrim(Context_GetRoot(), "\")
    g_ContextBrowserBreadcrumbSegments := [{ label: "context", path: root }]
    qEsc := ContextBrowser_EscapeLinkLabel(query)
    g_ContextBrowserBreadcrumbLink.Text := Format("<a id=`"1`">context</a> » search: `"{1}`" ({2})", qEsc, count)
}

ContextBrowser_IsSearchMode() {
    global g_ContextBrowserFilterQuery
    return (Trim(g_ContextBrowserFilterQuery) != "")
}

ContextBrowser_FormatEntryFolderPath(entry) {
    if (!IsObject(entry) || !entry.HasProp("relPath") || entry.relPath = "")
        return ""
    SplitPath entry.relPath, , &dir
    if (dir = "")
        return ""
    return StrReplace(dir, "\", " » ")
}

ContextBrowser_ClearEntryPathLabel() {
    global g_ContextBrowserEntryPathLabel
    if (!IsObject(g_ContextBrowserEntryPathLabel))
        return
    try g_ContextBrowserEntryPathLabel.Value := ""
    g_ContextBrowserEntryPathLabel.Opt("+Hidden")
}

ContextBrowser_UpdateEntryPathLabel(rowNum) {
    global g_ContextBrowserEntries, g_ContextBrowserEntryPathLabel
    if (!IsObject(g_ContextBrowserEntryPathLabel))
        return
    if (!ContextBrowser_IsSearchMode() || rowNum < 1 || rowNum > g_ContextBrowserEntries.Length) {
        ContextBrowser_ClearEntryPathLabel()
        return
    }
    pathText := ContextBrowser_FormatEntryFolderPath(g_ContextBrowserEntries[rowNum])
    if (pathText = "") {
        ContextBrowser_ClearEntryPathLabel()
        return
    }
    try g_ContextBrowserEntryPathLabel.Value := pathText
    g_ContextBrowserEntryPathLabel.Opt("-Hidden")
}

ContextBrowser_FormatEntryLabel(entry) {
    if (entry.type = "folder")
        return entry.name
    probe := Context_ProbeReference(entry.path)
    if (probe.isRef && probe.targetBasename != "")
        return entry.name "  →  " probe.targetBasename
    if (entry.type = "file" && Context_IsImagePath(entry.path)) {
        linked := ContextBrowser_FindLinkedResearchJson(entry.path)
        if (linked != "") {
            SplitPath linked, &base
            return entry.name "  ⇄  " base
        }
    }
    return entry.name
}

ContextBrowser_EntryKindLabel(entry) {
    if (entry.type = "parent")
        return "Up"
    if (entry.type = "folder")
        return "Folder"
    if (ContextBrowser_GetPreviewImagePath(entry) != "")
        return "Image"
    if (entry.type = "file") {
        SplitPath entry.path, , , &ext
        if (StrLower(ext) = "json")
            return "JSON"
    }
    return "File"
}

ContextBrowser_ReleasePreviewBitmap() {
    global g_ContextBrowserPreviewHbm
    if (g_ContextBrowserPreviewHbm) {
        try DllCall("DeleteObject", "ptr", g_ContextBrowserPreviewHbm)
        g_ContextBrowserPreviewHbm := 0
    }
}

ContextBrowser_ClearTextPreview() {
    global g_ContextBrowserPreviewText
    if (!IsObject(g_ContextBrowserPreviewText))
        return
    try g_ContextBrowserPreviewText.Value := ""
    g_ContextBrowserPreviewText.Opt("+Hidden")
}

ContextBrowser_ClearPreview() {
    global g_ContextBrowserPreviewCtrl
    ContextBrowser_ReleasePreviewBitmap()
    if (!IsObject(g_ContextBrowserPreviewCtrl))
        return
    try g_ContextBrowserPreviewCtrl.Value := ""
    g_ContextBrowserPreviewCtrl.Opt("+Hidden")
}

ContextBrowser_ClearAllPreviews() {
    ContextBrowser_ClearPreview()
    ContextBrowser_ClearTextPreview()
}

ContextBrowser_UpdatePreview(rowNum) {
    ContextBrowser_EnsureGlobals()
    global g_ContextBrowserEntries, g_ContextBrowserPreviewCtrl, g_ContextBrowserPreviewText, CONTEXT_PREVIEW_MAX_SIZE
    global g_ContextBrowserPreviewHbm
    if (!IsObject(g_ContextBrowserPreviewCtrl))
        return
    if (rowNum < 1 || rowNum > g_ContextBrowserEntries.Length) {
        ContextBrowser_ClearAllPreviews()
        ContextBrowser_UpdateEntryPathLabel(0)
        return
    }
    entry := g_ContextBrowserEntries[rowNum]
    imagePath := ContextBrowser_GetPreviewImagePath(entry)
    if (imagePath != "") {
        ContextBrowser_ClearTextPreview()
        fitW := 0
        fitH := 0
        if !ContextBrowser_GetFitImageSize(imagePath, CONTEXT_PREVIEW_MAX_SIZE, CONTEXT_PREVIEW_MAX_SIZE, &fitW, &fitH) {
            ContextBrowser_ClearPreview()
            ContextBrowser_UpdateEntryPathLabel(rowNum)
            return
        }
        try {
            ContextBrowser_ReleasePreviewBitmap()
            scaledHbm := LoadPicture(imagePath, "w" fitW " h" fitH)
            g_ContextBrowserPreviewHbm := scaledHbm
            g_ContextBrowserPreviewCtrl.Move(, , fitW, fitH)
            g_ContextBrowserPreviewCtrl.Value := "HBITMAP:*" scaledHbm
        } catch {
            ContextBrowser_ClearPreview()
            ContextBrowser_UpdateEntryPathLabel(rowNum)
            return
        }
        g_ContextBrowserPreviewCtrl.Opt("-Hidden")
        ContextBrowser_UpdateEntryPathLabel(rowNum)
        return
    }
    ContextBrowser_ClearPreview()
    jsonText := (entry.type = "file") ? ContextBrowser_GetJsonPreviewText(entry.path) : ""
    if (jsonText != "" && IsObject(g_ContextBrowserPreviewText)) {
        g_ContextBrowserPreviewText.Value := jsonText
        g_ContextBrowserPreviewText.Opt("-Hidden")
    } else {
        ContextBrowser_ClearTextPreview()
    }
    ContextBrowser_UpdateEntryPathLabel(rowNum)
}

ContextBrowser_GetActiveMonitorWorkArea(&left, &top, &right, &bottom) {
    MonitorGetWorkArea(1, &left, &top, &right, &bottom)
    activeWin := 0
    try activeWin := WinGetID("A")
    catch {
        return
    }
    if !activeWin
        return
    rect := Buffer(16, 0)
    if !DllCall("GetWindowRect", "ptr", activeWin, "ptr", rect)
        return
    centerX := NumGet(rect, 0, "int") + (NumGet(rect, 8, "int") - NumGet(rect, 0, "int")) // 2
    centerY := NumGet(rect, 4, "int") + (NumGet(rect, 12, "int") - NumGet(rect, 4, "int")) // 2
    loop MonitorGetCount() {
        MonitorGetWorkArea(A_Index, &l, &t, &r, &b)
        if (centerX >= l && centerX <= r && centerY >= t && centerY <= b) {
            left := l
            top := t
            right := r
            bottom := b
            return
        }
    }
}

ContextBrowser_GetEntryNameForLetterJump(entry) {
    if (entry.type = "parent")
        return ""
    return entry.name
}

ContextBrowser_IsActiveForLetterJump() {
    global g_ContextBrowserActive
    return g_ContextBrowserActive
}

ContextBrowser_GetEntriesForLetterJump() {
    global g_ContextBrowserEntries
    return g_ContextBrowserEntries
}

ContextBrowser_HandleLetterJump(rowNum) {
    ContextBrowser_EnsureGlobals()
    global g_ContextBrowserActive, g_ContextBrowserListView
    if (!g_ContextBrowserActive || !IsObject(g_ContextBrowserListView))
        return
    if (rowNum < 1)
        return
    ListView_SelectRowFocused(g_ContextBrowserListView, rowNum)
    ContextBrowser_UpdatePreview(rowNum)
}

ContextBrowser_StartLetterJump() {
    ContextBrowser_EnsureGlobals()
    global g_ContextBrowserLetterHook
    ModalListLetterJump_Start(&g_ContextBrowserLetterHook, ContextBrowser_IsActiveForLetterJump,
        ContextBrowser_GetEntriesForLetterJump, ContextBrowser_GetEntryNameForLetterJump,
        ContextBrowser_HandleLetterJump)
}

ContextBrowser_StopLetterJump() {
    global g_ContextBrowserLetterHook
    ModalListLetterJump_Stop(&g_ContextBrowserLetterHook)
}

ContextBrowser_DisableHotkeys() {
    try Hotkey("Backspace", "Off")
    try Hotkey("Enter", "Off")
    try Hotkey("^Enter", "Off")
    try Hotkey("+Enter", "Off")
    try Hotkey("^c", "Off")
    try Hotkey("^h", "Off")
}

ContextBrowser_EnableHotkeys() {
    try Hotkey("Backspace", (*) => ContextBrowser_HandleBack(), "On")
    try Hotkey("Enter", ContextBrowser_OnEnter, "On")
    try Hotkey("+Enter", ContextBrowser_PasteFocusedContentAsText, "On")
    try Hotkey("^Enter", ContextBrowser_PasteFocusedPathAsText, "On")
    try Hotkey("^c", ContextBrowser_CopyFocusedPath, "On")
    try Hotkey("^h", ContextBrowser_OpenFocusedInExplorer, "On")
}

ContextBrowser_SetListNavigationHotkeysEnabled(enabled) {
    if (enabled) {
        ContextBrowser_EnableHotkeys()
        ContextBrowser_StartLetterJump()
    } else {
        ContextBrowser_StopLetterJump()
        ContextBrowser_DisableHotkeys()
    }
}

ContextBrowser_GetFocusedEntry() {
    global g_ContextBrowserActive, g_ContextBrowserListView, g_ContextBrowserEntries
    if (!g_ContextBrowserActive || !IsObject(g_ContextBrowserListView))
        return false
    rowNum := g_ContextBrowserListView.GetNext(0, "F")
    if (rowNum < 1 || rowNum > g_ContextBrowserEntries.Length)
        return false
    return { rowNum: rowNum, entry: g_ContextBrowserEntries[rowNum] }
}

ContextBrowser_ResolveEntryPath(entry) {
    if (!IsObject(entry))
        return ""
    if (entry.type = "parent") {
        global g_ContextBrowserCurrentDir
        SplitPath g_ContextBrowserCurrentDir, , &parentDir
        return parentDir
    }
    if (entry.type = "folder")
        return entry.path
    pastePath := Context_ResolvePastePath(entry.path)
    return pastePath != "" ? pastePath : entry.path
}

ContextBrowser_CopyFocusedPath(*) {
    ContextBrowser_EnsureGlobals()
    focused := ContextBrowser_GetFocusedEntry()
    if (!IsObject(focused))
        return
    path := ContextBrowser_ResolveEntryPath(focused.entry)
    if (path = "")
        return
    A_Clipboard := path
}

ContextBrowser_OpenFocusedInExplorer(*) {
    ContextBrowser_EnsureGlobals()
    if ContextBrowser_IsFilterFocused()
        return
    focused := ContextBrowser_GetFocusedEntry()
    if (!IsObject(focused))
        return
    entry := focused.entry
    explorerCmd := ""
    if (entry.type = "parent") {
        path := ContextBrowser_ResolveEntryPath(entry)
        if (path != "" && DirExist(path))
            explorerCmd := 'explorer.exe "' path '"'
    } else if (entry.type = "folder") {
        if DirExist(entry.path)
            explorerCmd := 'explorer.exe "' entry.path '"'
    } else {
        path := ContextBrowser_ResolveEntryPath(entry)
        if (path != "" && FileExist(path))
            explorerCmd := 'explorer.exe /select,"' path '"'
    }
    if (explorerCmd = "")
        return
    StandardLoadingBar_Show("⏳ Opening in Explorer...", BANNER_ACCENT_INTERMEDIATE, { passive: true })
    try {
        Run explorerCmd
        CleanupContextBrowser()
    } finally {
        StandardLoadingBar_Hide(350)
    }
}

ContextBrowser_PasteFocusedPathAsText(*) {
    ContextBrowser_EnsureGlobals()
    focused := ContextBrowser_GetFocusedEntry()
    if (!IsObject(focused) || focused.entry.type != "file")
        return
    pastePath := Context_ResolvePastePath(focused.entry.path)
    if (pastePath = "") {
        ShowCenteredOverlay_Utils("❌ Reference target not found for: " focused.entry.name, 2500, BANNER_ACCENT_ERROR)
        return
    }
    CleanupContextBrowser()
    Sleep 50
    InsertText(pastePath)
}

ContextBrowser_PasteFocusedContentAsText(*) {
    ContextBrowser_EnsureGlobals()
    focused := ContextBrowser_GetFocusedEntry()
    if (!IsObject(focused) || focused.entry.type != "file")
        return
    pastePath := Context_ResolvePastePath(focused.entry.path)
    if (pastePath = "") {
        ShowCenteredOverlay_Utils("❌ Reference target not found for: " focused.entry.name, 2500, BANNER_ACCENT_ERROR)
        return
    }
    if Context_IsImagePath(pastePath) {
        ShowCenteredOverlay_Utils("❌ Cannot paste image as text: " focused.entry.name, 2500, BANNER_ACCENT_ERROR)
        return
    }
    content := Context_ResolvePasteTextContent(focused.entry.path)
    if (content = "") {
        ShowCenteredOverlay_Utils("❌ No text content to paste: " focused.entry.name, 2500, BANNER_ACCENT_ERROR)
        return
    }
    CleanupContextBrowser()
    Sleep 50
    InsertText(content)
}

ContextBrowser_OnFilterChange(*) {
    ContextBrowser_EnsureGlobals()
    global g_ContextBrowserFilterCtrl, g_ContextBrowserFilterQuery, g_ContextBrowserActive
    if (!g_ContextBrowserActive || !IsObject(g_ContextBrowserFilterCtrl))
        return
    g_ContextBrowserFilterQuery := g_ContextBrowserFilterCtrl.Value
    ContextBrowser_RefreshView()
}

ContextBrowser_OnFilterFocus(*) {
    global g_ContextBrowserFilterCtrl, g_ContextBrowserActive, g_ContextBrowserFilterTyping
    g_ContextBrowserFilterTyping := true
    if (g_ContextBrowserActive)
        ContextBrowser_SetListNavigationHotkeysEnabled(false)
}

ContextBrowser_OnFilterKillFocus(*) {
    global g_ContextBrowserActive, g_ContextBrowserFilterTyping, g_ContextBrowserSuppressFilterKillFocus
    if (g_ContextBrowserSuppressFilterKillFocus)
        return
    g_ContextBrowserFilterTyping := false
    if (g_ContextBrowserActive)
        ContextBrowser_SetListNavigationHotkeysEnabled(true)
}

CleanupContextBrowser() {
    ContextBrowser_EnsureGlobals()
    global g_ContextBrowserActive, g_ContextBrowserGui, g_ContextBrowserCurrentDir
    global g_ContextBrowserEntries, g_ContextBrowserListView, g_ContextBrowserBreadcrumbLink,
        g_ContextBrowserPreviewCtrl
    global g_ContextBrowserPreviewText, g_ContextBrowserFilterCtrl, g_ContextBrowserLastDir, g_ContextBrowserFileIndex

    if (g_ContextBrowserCurrentDir != "" && DirExist(g_ContextBrowserCurrentDir))
        g_ContextBrowserLastDir := g_ContextBrowserCurrentDir

    g_ContextBrowserActive := false
    g_ContextBrowserFilterTyping := false
    g_ContextBrowserSuppressFilterKillFocus := false
    ContextBrowser_StopLetterJump()
    ContextBrowser_DisableHotkeys()
    ContextBrowser_ReleasePreviewBitmap()
    if (IsObject(g_ContextBrowserGui)) {
        try g_ContextBrowserGui.Destroy()
        g_ContextBrowserGui := false
    }
    g_ContextBrowserEntries := []
    g_ContextBrowserFileIndex := []
    g_ContextBrowserListView := false
    g_ContextBrowserBreadcrumbLink := false
    g_ContextBrowserBreadcrumbSegments := []
    g_ContextBrowserPreviewCtrl := false
    g_ContextBrowserPreviewText := false
    g_ContextBrowserFilterCtrl := false
    g_ContextBrowserEntryPathLabel := false
}

HandleContextBrowserEscape(*) {
    ContextBrowser_EnsureGlobals()
    global g_ContextBrowserActive
    if (g_ContextBrowserActive)
        CleanupContextBrowser()
}

ContextBrowser_HandleBack() {
    ContextBrowser_EnsureGlobals()
    global g_ContextBrowserActive, g_ContextBrowserCurrentDir
    if (!g_ContextBrowserActive)
        return
    if ContextBrowser_IsFilterFocused()
        return
    if Context_IsAtRoot(g_ContextBrowserCurrentDir)
        return
    SplitPath g_ContextBrowserCurrentDir, , &parentDir
    if (parentDir = "" || !DirExist(parentDir))
        return
    g_ContextBrowserCurrentDir := parentDir
    ContextBrowser_RefreshView()
}

ContextBrowser_ActivateEntry(entry) {
    global g_ContextBrowserCurrentDir
    if (!IsObject(entry))
        return
    if (entry.type = "parent") {
        ContextBrowser_HandleBack()
        return
    }
    if (entry.type = "folder") {
        g_ContextBrowserCurrentDir := entry.path
        ContextBrowser_RefreshView()
        return
    }
    pastePath := Context_ResolvePastePath(entry.path)
    if (pastePath = "") {
        ShowCenteredOverlay_Utils("❌ Reference target not found for: " entry.name, 2500, BANNER_ACCENT_ERROR)
        return
    }
    CleanupContextBrowser()
    Sleep 50
    if !InsertFiles([pastePath]) {
        ShowCenteredOverlay_Utils("⚠ Could not attach file — pasting path as text", 2800, BANNER_ACCENT_ERROR)
        InsertText(pastePath)
        return
    }
    if InsertFiles_IsAiChatForeground()
        Sleep 300
}

ContextBrowser_OnSelectRow(rowNum) {
    ContextBrowser_EnsureGlobals()
    global g_ContextBrowserActive, g_ContextBrowserEntries
    if (!g_ContextBrowserActive)
        return
    if (rowNum < 1 || rowNum > g_ContextBrowserEntries.Length)
        return
    ContextBrowser_ActivateEntry(g_ContextBrowserEntries[rowNum])
}

ContextBrowser_OnListDoubleClick(lv, guiEvent, *) {
    rowNum := 0
    try rowNum := guiEvent.EventInfo
    if !rowNum
        rowNum := lv.GetNext(0, "F")
    ContextBrowser_OnSelectRow(rowNum)
}

ContextBrowser_OnItemFocus(lv, guiEvent, *) {
    ContextBrowser_EnsureGlobals()
    global g_ContextBrowserActive
    if (!g_ContextBrowserActive)
        return
    rowNum := 0
    try rowNum := guiEvent.EventInfo
    if !rowNum
        rowNum := lv.GetNext(0, "F")
    ContextBrowser_UpdatePreview(rowNum)
}

ContextBrowser_OnEnter(*) {
    ContextBrowser_EnsureGlobals()
    global g_ContextBrowserListView
    if ContextBrowser_IsFilterFocused()
        return
    if (!IsObject(g_ContextBrowserListView))
        return
    rowNum := g_ContextBrowserListView.GetNext(0, "F")
    ContextBrowser_OnSelectRow(rowNum)
}

ContextBrowser_RefreshView() {
    global g_ContextBrowserActive, g_ContextBrowserCurrentDir, g_ContextBrowserEntries
    global g_ContextBrowserGui, g_ContextBrowserListView, g_ContextBrowserFilterCtrl
    global g_ContextBrowserFilterQuery, g_ContextBrowserFilterTyping, g_ContextBrowserSuppressFilterKillFocus

    dir := g_ContextBrowserCurrentDir
    if (dir = "" || !DirExist(dir)) {
        TrayTip("Context", "Folder not found.", "IconX")
        SetTimer(() => TrayTip(), -5000)
        CleanupContextBrowser()
        return
    }

    wasTypingInFilter := g_ContextBrowserFilterTyping
    g_ContextBrowserSuppressFilterKillFocus := wasTypingInFilter

    if (IsObject(g_ContextBrowserFilterCtrl))
        g_ContextBrowserFilterQuery := g_ContextBrowserFilterCtrl.Value
    q := Trim(g_ContextBrowserFilterQuery)
    if (q != "") {
        g_ContextBrowserEntries := ContextBrowser_SearchFileIndex(ContextBrowser_GetFileIndex(), q)
        ContextBrowser_UpdateSearchBreadcrumbs(q, g_ContextBrowserEntries.Length)
    } else {
        g_ContextBrowserEntries := ContextBrowser_BuildViewEntries(dir)
        ContextBrowser_UpdateBreadcrumbs(dir)
    }

    if (!IsObject(g_ContextBrowserListView)) {
        g_ContextBrowserSuppressFilterKillFocus := false
        return
    }

    lv := g_ContextBrowserListView
    lv.Opt("-Redraw")
    lv.Delete()
    for entry in g_ContextBrowserEntries
        lv.Add("", ContextBrowser_EntryKindLabel(entry), ContextBrowser_FormatEntryLabel(entry))
    lv.Opt("+Redraw")
    if (g_ContextBrowserEntries.Length) {
        typingInFilter := wasTypingInFilter
        if (typingInFilter)
            lv.Modify(1, "Select Vis")
        else
            lv.Modify(1, "Select Focus Vis")
        if (!typingInFilter) {
            try lv.Focus()
            catch {
            }
        }
        ContextBrowser_UpdatePreview(1)
    } else {
        ContextBrowser_UpdatePreview(0)
    }
    g_ContextBrowserActive := true
    g_ContextBrowserSuppressFilterKillFocus := false
    if (wasTypingInFilter)
        g_ContextBrowserFilterTyping := true
    if (!g_ContextBrowserFilterTyping)
        ContextBrowser_StartLetterJump()
}

ContextBrowser_CreateGui() {
    global g_ContextBrowserGui, g_ContextBrowserListView, g_ContextBrowserBreadcrumbLink, g_ContextBrowserPreviewCtrl
    global g_ContextBrowserPreviewText, g_ContextBrowserFilterCtrl, g_ContextBrowserEntryPathLabel

    g_ContextBrowserGui := Gui("+AlwaysOnTop +Resize +MinSize620x340", "Context")
    g_ContextBrowserGui.SetFont("s10", "Segoe UI")
    g_ContextBrowserBreadcrumbLink := g_ContextBrowserGui.Add("Link", "xm w740 h22", "context")
    g_ContextBrowserBreadcrumbLink.OnEvent("Click", ContextBrowser_OnBreadcrumbLinkClick)
    g_ContextBrowserFilterCtrl := g_ContextBrowserGui.Add("Edit", "xm w740", "")
    g_ContextBrowserFilterCtrl.OnEvent("Change", ContextBrowser_OnFilterChange)
    g_ContextBrowserFilterCtrl.OnEvent("Focus", ContextBrowser_OnFilterFocus)
    g_ContextBrowserFilterCtrl.OnEvent("LoseFocus", ContextBrowser_OnFilterKillFocus)
    g_ContextBrowserListView := g_ContextBrowserGui.Add("ListView", "xm w480 h360 Grid -Multi", ["Kind", "Name"])
    g_ContextBrowserPreviewCtrl := g_ContextBrowserGui.Add("Picture", "x+10 yp BackgroundTrans Hidden")
    g_ContextBrowserPreviewText := g_ContextBrowserGui.Add("Edit", "x+10 yp w234 h360 ReadOnly -VScroll Multi Hidden")
    g_ContextBrowserListView.ModifyCol(1, 72)
    g_ContextBrowserListView.ModifyCol(2, "AutoHdr")
    g_ContextBrowserListView.OnEvent("DoubleClick", ContextBrowser_OnListDoubleClick)
    g_ContextBrowserListView.OnEvent("ItemFocus", ContextBrowser_OnItemFocus)
    g_ContextBrowserGui.SetFont("s8", "Segoe UI")
    g_ContextBrowserEntryPathLabel := g_ContextBrowserGui.Add("Text", "xm w740 h28 +Wrap +Hidden", "")
    g_ContextBrowserGui.SetFont("s10", "Segoe UI")
    g_ContextBrowserGui.Add("Text", "xm",
        "click path to jump · filter searches all context files · ↑↓ · letter jump · Enter attach · Shift+Enter text · Ctrl+Enter path · Ctrl+C copy · Ctrl+H explorer · Esc close"
    )
    g_ContextBrowserGui.OnEvent("Escape", HandleContextBrowserEscape)
    g_ContextBrowserGui.OnEvent("Close", (*) => CleanupContextBrowser())
}

ContextBrowser_ShowGui() {
    global g_ContextBrowserGui, g_ContextBrowserListView

    MonitorGetWorkArea(1, &monitorLeft, &monitorTop, &monitorRight, &monitorBottom)
    ContextBrowser_GetActiveMonitorWorkArea(&monitorLeft, &monitorTop, &monitorRight, &monitorBottom)
    monitorWidth := monitorRight - monitorLeft
    monitorHeight := monitorBottom - monitorTop

    g_ContextBrowserGui.Show("w760 h500")
    g_ContextBrowserGui.GetPos(, , &gw, &gh)
    guiX := monitorLeft + (monitorWidth - gw) // 2
    guiY := monitorTop + (monitorHeight - gh) // 2
    margin := 16
    guiX := Max(monitorLeft + margin, Min(guiX, monitorRight - gw - margin))
    guiY := Max(monitorTop + margin, Min(guiY, monitorBottom - gh - margin))
    g_ContextBrowserGui.Show("x" guiX " y" guiY " w760 h500")

    try ContextBrowser_EnableHotkeys()
    catch {
    }
}

ContextBrowser_OpenAtCurrentDir() {
    ContextBrowser_EnsureGlobals()
    global g_ContextBrowserGui, g_ContextBrowserActive, g_ContextBrowserCurrentDir

    dir := g_ContextBrowserCurrentDir
    if (dir = "" || !DirExist(dir)) {
        TrayTip("Context", "Folder not found.", "IconX")
        SetTimer(() => TrayTip(), -5000)
        CleanupContextBrowser()
        return
    }

    if (!IsObject(g_ContextBrowserGui) || !IsObject(g_ContextBrowserBreadcrumbLink)) {
        if IsObject(g_ContextBrowserGui) {
            try g_ContextBrowserGui.Destroy()
            g_ContextBrowserGui := false
        }
        ContextBrowser_CreateGui()
        ContextBrowser_ShowGui()
    }
    ContextBrowser_RefreshView()
    g_ContextBrowserActive := true
    ContextBrowser_FocusFilterField()
}

ShowContextBrowser() {
    ContextBrowser_EnsureGlobals()
    global g_ContextBrowserActive, g_ContextBrowserCurrentDir, g_ContextBrowserLastDir

    if (g_ContextBrowserActive) {
        CleanupContextBrowser()
        return
    }

    try {
        if (IsSet(g_HotstringSelectorActive) && g_HotstringSelectorActive)
            CleanupHotstringSelector()
    } catch {
    }
    try {
        if (IsSet(g_ProjectSelectorActive) && g_ProjectSelectorActive && IsSet(CleanupProjectSelector))
            CleanupProjectSelector()
    } catch {
    }

    root := Context_GetRoot()
    if !DirExist(root) {
        TrayTip("Context", "context folder not found at:`n" root, "IconX")
        SetTimer(() => TrayTip(), -5000)
        return
    }

    if (g_ContextBrowserLastDir != "" && DirExist(g_ContextBrowserLastDir))
        g_ContextBrowserCurrentDir := g_ContextBrowserLastDir
    else
        g_ContextBrowserCurrentDir := root
    ContextBrowser_OpenAtCurrentDir()
}
