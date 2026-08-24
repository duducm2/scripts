; =============================================================================
; Utils module: prompt_context_picker.ahk
; Dynamic context file picker at prompt paste time
; =============================================================================

PromptContextCatalog_List(catalog) {
    catalog := StrLower(Trim(catalog))
    if (catalog = "mnemonic_stories")
        return PromptContextCatalog_MnemonicStories()
    return []
}

PromptContextCatalog_ResolveStoryPath(study) {
    if (!IsObject(study))
        return ""
    slug := Trim(study.Has("notes_rel_path") ? study["notes_rel_path"] : "")
    if (slug = "")
        return ""
    slug := StrReplace(slug, "/", "\")
    practice := Palace_OutputDir() . "\practice\" . slug . ".md"
    if (FileExist(practice))
        return practice
    notesRoot := Palace_NotesStudiesRoot()
    if (notesRoot != "") {
        legacy := RTrim(notesRoot, "\") . "\" . slug . "\mnemonics-" . slug . ".md"
        if (FileExist(legacy))
            return legacy
    }
    return ""
}

PromptContextCatalog_StudyTailSummary(studyId) {
    palaces := Palace_Load("palaces")
    beasts := Palace_Load("beasts")
    lastPalace := ""
    lastNum := -1
    for p in palaces {
        if (p["study_id"] != studyId)
            continue
        num := Number(p.Has("palace_number") ? p["palace_number"] : 0)
        if (num >= lastNum) {
            lastNum := num
            lastPalace := p
        }
    }
    if (!IsObject(lastPalace) || lastNum < 0)
        return "No palaces yet"
    charName := Trim(lastPalace.Has("character_name") ? lastPalace["character_name"] : "")
    palaceId := lastPalace.Has("id") ? lastPalace["id"] : ""
    lastBeast := ""
    lastSort := -1
    for b in beasts {
        if (b["palace_id"] != palaceId)
            continue
        sort := Number(b.Has("sort_order") ? b["sort_order"] : 0)
        if (sort >= lastSort) {
            lastSort := sort
            lastBeast := b
        }
    }
    tail := "Palace " . lastNum
    if (IsObject(lastBeast)) {
        peg := Trim(lastBeast.Has("peg_code") ? lastBeast["peg_code"] : "")
        bname := Trim(lastBeast.Has("beast_name") ? lastBeast["beast_name"] : "")
        if (peg != "")
            tail .= " · [" . peg . "]"
        if (bname != "")
            tail .= " " . bname
    } else {
        tail .= " · (no beasts)"
    }
    if (charName != "")
        tail .= " · Character: " . charName
    return tail
}

PromptContextCatalog_MnemonicStories() {
    Palace_EnsureData()
    studies := Palace_Load("studies")
    items := []
    for s in studies {
        if (s.Has("active") && s["active"] = "0")
            continue
        path := PromptContextCatalog_ResolveStoryPath(s)
        if (path = "" || !FileExist(path))
            continue
        mtime := 0
        try mtime := Number(FileGetTime(path, "M"))
        catch {
            mtime := 0
        }
        studyId := s.Has("id") ? s["id"] : ""
        slug := Trim(s.Has("notes_rel_path") ? s["notes_rel_path"] : "")
        title := s.Has("title") ? s["title"] : slug
        items.Push({
            studyId: studyId,
            studyTitle: title,
            path: path,
            tail: PromptContextCatalog_StudyTailSummary(studyId),
            mtime: mtime,
            fileLabel: PromptData_ToStoredPath(path)
        })
    }
    PromptContextPicker_SortItemsByMtime(items)
    return items
}

PromptContextPicker_SortItemsByMtime(items) {
    loop items.Length - 1 {
        swapped := false
        loop items.Length - A_Index {
            i := A_Index
            if (items[i].mtime < items[i + 1].mtime) {
                tmp := items[i]
                items[i] := items[i + 1]
                items[i + 1] := tmp
                swapped := true
            }
        }
        if (!swapped)
            break
    }
}

PromptContextPicker_ItemFromPath(path) {
    entry := PromptData_NewContextEntry(path)
    abs := PromptData_ContextEntryPath(entry)
    if (abs = "")
        abs := path
    mtime := 0
    try mtime := Number(FileGetTime(abs, "M"))
    catch {
        mtime := 0
    }
    studyTitle := RegExReplace(abs, ".*\\", "")
    tail := ""
    studyId := ""
    practiceNeedle := StrLower(Palace_OutputDir() . "\practice\")
    if (InStr(StrLower(abs), practiceNeedle)) {
        slug := RegExReplace(studyTitle, "\.md$", "", , 1)
        Palace_EnsureData()
        for s in Palace_Load("studies") {
            if (StrLower(Trim(s.Has("notes_rel_path") ? s["notes_rel_path"] : "")) = StrLower(slug)) {
                studyId := s.Has("id") ? s["id"] : ""
                studyTitle := s.Has("title") ? s["title"] : slug
                tail := PromptContextCatalog_StudyTailSummary(studyId)
                break
            }
        }
    }
    if (tail = "")
        tail := FileExist(abs) ? "Custom context file" : "(missing file)"
    return {
        studyId: studyId,
        studyTitle: studyTitle,
        path: abs,
        tail: tail,
        mtime: mtime,
        fileLabel: PromptData_ToStoredPath(abs)
    }
}

PromptContextPicker_BuildPool(prompt) {
    if (!IsObject(prompt))
        return []
    items := []
    seen := Map()
    for e in PromptData_SelectableContextEntriesForCurrentEnv(prompt) {
        abs := PromptData_ContextEntryPath(e)
        if (abs = "" || !FileExist(abs))
            continue
        key := StrLower(abs)
        if (seen.Has(key))
            continue
        seen[key] := true
        items.Push(PromptContextPicker_ItemFromPath(abs))
    }
    PromptContextPicker_SortItemsByMtime(items)
    return items
}

PromptContext_MergeEntries(staticEntries, pickedEntries) {
    out := []
    seen := Map()
    for e in PromptData_ParseContextEntries(staticEntries) {
        p := StrLower(PromptData_ContextEntryPath(e))
        if (p = "" || seen.Has(p))
            continue
        seen[p] := true
        out.Push(e)
    }
    for e in PromptData_ParseContextEntries(pickedEntries) {
        p := StrLower(PromptData_ContextEntryPath(e))
        if (p = "" || seen.Has(p))
            continue
        seen[p] := true
        out.Push(e)
    }
    return out
}

; Returns array of context entries, or false if user cancelled.
PromptContextPicker_ShowPool(items) {
    if (!IsObject(items) || items.Length = 0)
        return []
    lastStudy := Trim(Palace_Setting("General", "LastStudyId", ""))
    preCheck := Map()
    if (lastStudy != "") {
        for it in items {
            if (it.studyId = lastStudy)
                preCheck[it.path] := true
        }
    }
    result := false
    g := Gui("+AlwaysOnTop +ToolWindow", "Attach context files")
    g.SetFont("s10", "Segoe UI")
    g.BackColor := "1E1E1E"
    g.SetFont("s10 cWhite", "Segoe UI")
    g.Add("Text", "x12 y10 w560",
        "Select files to attach (Space toggles check). Static context files attach automatically.")
    g.SetFont("s9 cA0A0A0 Norm", "Segoe UI")
    g.Add("Text", "x12 y32 w560",
        "Esc / Cancel aborts the prompt. Attach with none checked continues with static files only.")
    lv := g.Add("ListView", "x12 y56 w560 h300 Checked Grid", ["Label", "Tail / note", "File"])
    rowPaths := []
    for it in items {
        opts := preCheck.Has(it.path) ? "Check" : ""
        lv.Add(opts, it.studyTitle, it.tail, it.fileLabel)
        rowPaths.Push(it.path)
    }
    lv.ModifyCol(1, 140)
    lv.ModifyCol(2, 240)
    lv.ModifyCol(3, 160)
    ; GetNext(..., "Checked") starts BELOW StartingRow — use row-1 to test this row.
    RowIsChecked(row) {
        return row > 0 && lv.GetNext(row - 1, "Checked") = row
    }
    ToggleRow(*) {
        row := lv.GetNext(0, "Focused")
        if (!row)
            row := lv.GetNext(0, "Selected")
        if (!row)
            return
        if (RowIsChecked(row))
            lv.Modify(row, "-Check")
        else
            lv.Modify(row, "Check")
    }
    PickOk(*) {
        picked := []
        row := 0
        loop {
            row := lv.GetNext(row, "Checked")
            if (!row)
                break
            if (row <= rowPaths.Length)
                picked.Push(PromptData_NewContextEntry(rowPaths[row]))
        }
        if (picked.Length = 0)
            ShowCenteredOverlay_Utils("No files selected — continuing with static attachments only", 2600,
                BANNER_ACCENT_ERROR)
        result := picked
        g.Destroy()
    }
    PickCancel(*) {
        result := false
        g.Destroy()
    }
    ; No Default on Attach — Space must toggle checks, not activate Attach (Enter still works via hotkey).
    g.Add("Button", "x12 y364 w100", "Attach").OnEvent("Click", PickOk)
    g.Add("Button", "x+8 yp w100", "Cancel").OnEvent("Click", PickCancel)
    g.OnEvent("Close", PickCancel)
    g.OnEvent("Escape", PickCancel)
    lv.OnEvent("DoubleClick", ToggleRow)
    HotIfWinActive("ahk_id " g.Hwnd)
    try Hotkey("Space", ToggleRow, "On")
    catch {
    }
    try Hotkey("Enter", PickOk, "On")
    catch {
    }
    g.Show()
    try {
        lv.Focus()
        if (rowPaths.Length > 0)
            lv.Modify(1, "Select Focus Vis")
    } catch {
    }
    try WinWaitClose("ahk_id " g.Hwnd)
    catch {
    }
    try Hotkey("Space", "Off")
    catch {
    }
    try Hotkey("Enter", "Off")
    catch {
    }
    try HotIf()
    catch {
    }
    return result
}

; Legacy catalog-only entry (editor refresh uses catalog builders directly).
PromptContextPicker_Show(catalog) {
    return PromptContextPicker_ShowPool(PromptContextCatalog_List(catalog))
}

PromptContextPicker_AppendMnemonicStoriesToSelectable(side) {
    global IS_WORK_ENVIRONMENT, g_PromptEditorPersonalSelectablePaths, g_PromptEditorWorkSelectablePaths
    arr := (side = "work") ? g_PromptEditorWorkSelectablePaths : g_PromptEditorPersonalSelectablePaths
    if (!IsObject(arr))
        arr := []
    seen := Map()
    for e in arr {
        p := StrLower(PromptData_ContextEntryPath(e))
        if (p != "")
            seen[p] := true
    }
    for it in PromptContextCatalog_MnemonicStories() {
        p := StrLower(it.path)
        if (seen.Has(p))
            continue
        seen[p] := true
        arr.Push(PromptData_NewContextEntry(it.path))
    }
    if (side = "work")
        g_PromptEditorWorkSelectablePaths := arr
    else
        g_PromptEditorPersonalSelectablePaths := arr
}
