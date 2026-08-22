; =============================================================================
; Utils module: mnemonic_palace_import.ahk
; Import AI-generated PALACE_*.csv from Desktop
; =============================================================================

Palace_DesktopNewest(pattern) {
    newest := ""
    newestTime := 0
    loop files A_Desktop . "\" . pattern, "F" {
        ts := Number(A_LoopFileTimeModified)
        if (ts > newestTime) {
            newestTime := ts
            newest := A_LoopFileFullPath
        }
    }
    return newest
}

Palace_ArchiveImported(path) {
    destDir := Palace_DataDir() . "\imported"
    if (!DirExist(destDir))
        DirCreate(destDir)
    SplitPath(path, &name)
    dest := destDir . "\" . FormatTime(, "yyyyMMdd-HHmmss") . "_" . name
    try FileMove(path, dest, 1)
    catch {
        try FileCopy(path, dest, 1)
        catch {
        }
    }
}

; Strip Gemini preambles; find a header that looks like palace CSV.
Palace_ReadAiImportCsv(path, headerHint := "beast_id") {
    text := Palace_ReadUtf8(path)
    if (text = "")
        return []
    text := StrReplace(text, "`r`n", "`n")
    text := StrReplace(text, "`r", "`n")
    lines := StrSplit(text, "`n")
    start := 0
    for idx, line in lines {
        t := Trim(line)
        if (t = "")
            continue
        lower := StrLower(t)
        if (InStr(lower, "file:") = 1 || SubStr(t, 1, 1) = "#")
            continue
        if (InStr(lower, headerHint) || (InStr(lower, "id") && InStr(lower, "kind"))) {
            start := idx
            break
        }
        if ((InStr(lower, "palace_id") || InStr(lower, "street_id")) && InStr(lower, "peg_code")) {
            start := idx
            break
        }
        if (InStr(lower, "study_id") && (InStr(lower, "palace_number") || InStr(lower, "street_number"))) {
            start := idx
            break
        }
    }
    if (!start)
        return Palace_ReadCsv(path)
    cleaned := ""
    loop lines.Length {
        if (A_Index < start)
            continue
        if (cleaned != "")
            cleaned .= "`n"
        cleaned .= lines[A_Index]
    }
    tmp := A_Temp . "\palace_ai_import_norm.csv"
    Palace_WriteUtf8(tmp, cleaned)
    rows := Palace_ReadCsv(tmp)
    try FileDelete(tmp)
    catch {
    }
    return rows
}

; Map legacy street_* column names to palace_* on import rows.
Palace_NormalizePalaceImportRow(r) {
    if (!r.Has("palace_id") && r.Has("street_id"))
        r["palace_id"] := r["street_id"]
    if (!r.Has("palace_number") && r.Has("street_number"))
        r["palace_number"] := r["street_number"]
    if (!r.Has("image_prompt"))
        r["image_prompt"] := ""
    return r
}

Palace_NormalizeBeastImportRow(r) {
    if (!r.Has("palace_id") && r.Has("street_id"))
        r["palace_id"] := r["street_id"]
    ; Accept legacy STREET_ ids by rewriting to PALACE_
    if (r.Has("palace_id")) {
        pid := Trim(r["palace_id"])
        if (SubStr(pid, 1, 7) = "STREET_")
            r["palace_id"] := "PALACE_" . SubStr(pid, 8)
    }
    return r
}

Palace_ShowImportMenu() {
    global g_PalaceGui
    Palace_CloseGui()
    Palace_EnsureData()
    g_PalaceGui := Gui("+AlwaysOnTop +ToolWindow", "Memory Palace — AI import")
    g_PalaceGui.SetFont("s10", "Segoe UI")
    g_PalaceGui.BackColor := "1E1E1E"
    g_PalaceGui.Add("Text", "x16 y12 w640 cWhite", "Import Desktop PALACE_*.csv deliverables")
    g_PalaceGui.SetFont("s9 cA0A0A0", "Segoe UI")
    g_PalaceGui.Add("Text", "x16 y40 w640",
        "Save AI output to Desktop, rename to CSV, then import. Preview before commit.")

    items := [
        ["1", "Import atoms", "PALACE_ATOMS*.csv — knowledge atoms"],
        ["2", "Import beasts", "PALACE_BEASTS*.csv — peg holders"],
        ["3", "Import palaces", "PALACE_PALACES*.csv — Memory Palaces (legacy STREETS OK)"],
        ["4", "Import any", "Newest PALACE_*.csv (auto-detect)"]
    ]
    y := 80
    for it in items {
        g_PalaceGui.SetFont("s14 cF1C40F Bold", "Segoe UI")
        g_PalaceGui.Add("Text", "x28 y" . y . " w40", "[" . it[1] . "]")
        g_PalaceGui.SetFont("s11 cWhite Bold", "Segoe UI")
        g_PalaceGui.Add("Text", "x72 y" . y . " w500", it[2])
        g_PalaceGui.SetFont("s9 cA0A0A0 Norm", "Segoe UI")
        g_PalaceGui.Add("Text", "x72 y" . (y + 22) . " w500", it[3])
        y += 56
    }
    g_PalaceGui.SetFont("s9 c808080", "Segoe UI")
    g_PalaceGui.Add("Text", "x16 y" . y . " w640", "Esc / Backspace — menu")
    g_PalaceGui.OnEvent("Close", (*) => Palace_CloseGui())
    g_PalaceGui.OnEvent("Escape", (*) => Palace_ShowMainMenu())
    Palace_BindHotkeys([
        ["1", (*) => Palace_ImportAtomsFromDesktop()],
        ["2", (*) => Palace_ImportBeastsFromDesktop()],
        ["3", (*) => Palace_ImportPalacesFromDesktop()],
        ["4", (*) => Palace_ImportAutoFromDesktop()],
        ["Backspace", (*) => Palace_ShowMainMenu()],
        ["Escape", (*) => Palace_ShowMainMenu()]
    ])
    Palace_CenterGui(g_PalaceGui, 680, y + 50)
}

Palace_ImportConfirmPreview(title, labels) {
    global g_PalaceGui
    owner := ""
    try {
        if (IsObject(g_PalaceGui))
            owner := " +Owner" . g_PalaceGui.Hwnd
    } catch {
        owner := ""
    }
    Palace_DialogsBegin()
    g := Gui("+AlwaysOnTop +ToolWindow" . owner, title)
    g.SetFont("s10", "Segoe UI")
    lv := g.Add("ListView", "w640 h320 Grid", ["Row"])
    for lab in labels
        lv.Add("", lab)
    lv.ModifyCol(1, "AutoHdr")
    ok := false
    g.Add("Button", "y+8 w120 Default", "Import").OnEvent("Click", (*) => (ok := true, g.Destroy()))
    g.Add("Button", "x+8 w100", "Cancel").OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())
    g.Show()
    try WinWaitClose("ahk_id " g.Hwnd)
    catch {
    }
    Palace_DialogsEnd()
    return ok
}

Palace_ImportAtomsFromDesktop(path := "") {
    if (path = "")
        path := Palace_DesktopNewest("PALACE_ATOMS*.csv")
    if (path = "")
        path := Palace_DesktopNewest("PALACE_ATOMS*.txt")
    if (path = "") {
        Palace_Notify("No PALACE_ATOMS*.csv on Desktop", 2200, BANNER_ACCENT_ERROR)
        return false
    }
    rows := Palace_ReadAiImportCsv(path, "beast_id")
    if (!rows.Length) {
        Palace_Notify("No atom rows found", 2000, BANNER_ACCENT_ERROR)
        return false
    }
    labels := []
    for r in rows {
        lab := (r.Has("kind") ? r["kind"] : "?") . " · "
        . (r.Has("zone") ? r["zone"] : "") . " · "
        . (r.Has("concept") ? SubStr(r["concept"], 1, 50)
            : (r.Has("context") ? SubStr(r["context"], 1, 50) : ""))
        labels.Push(lab)
    }
    if (!Palace_ImportConfirmPreview("Import " . rows.Length . " atom(s)?", labels))
        return false
    atoms := Palace_Load("atoms")
    beasts := Palace_Load("beasts")
    merged := 0
    for r in rows {
        beastId := r.Has("beast_id") ? Trim(r["beast_id"]) : ""
        if (beastId = "" || !Palace_FindById(beasts, beastId)) {
            Palace_Notify("Skip row: unknown beast_id " . beastId, 2000, BANNER_ACCENT_ERROR)
            continue
        }
        kind := r.Has("kind") ? StrLower(Trim(r["kind"])) : "single"
        if (kind = "")
            kind := "single"
        if (kind = "subtopic")
            kind := "zoned"
        id := r.Has("id") && Trim(r["id"]) != "" ? Trim(r["id"]) : Palace_NextId("ATOM_", atoms, 4)
        if (Palace_FindById(atoms, id))
            id := Palace_NextId("ATOM_", atoms, 4)
        concept := r.Has("concept") ? r["concept"] : (r.Has("context") ? r["context"] : "")
        story := r.Has("story") ? r["story"] : (r.Has("narrative") ? r["narrative"] : "")
        sensory := r.Has("sensory") ? r["sensory"] : (r.Has("sensory_channel") ? r["sensory_channel"] : "")
        row := Map(
            "id", id,
            "beast_id", beastId,
            "kind", kind,
            "zone", r.Has("zone") ? r["zone"] : "",
            "zone_label", r.Has("zone_label") ? r["zone_label"] : "",
            "concept", concept,
            "quote", r.Has("quote") ? r["quote"] : "",
            "story", story,
            "ipa", r.Has("ipa") ? r["ipa"] : "",
            "sensory", sensory,
            "sort_order", r.Has("sort_order") ? r["sort_order"] : "1"
        )
        proposed := Palace_FilterBy(atoms, "beast_id", beastId)
        proposed.Push(row)
        err := Palace_ValidateBeastAtoms(beastId, proposed)
        if (err != "") {
            Palace_Notify("Skip: " . err, 2500, BANNER_ACCENT_ERROR)
            continue
        }
        atoms.Push(row)
        merged += 1
    }
    Palace_Save("atoms", atoms)
    Palace_ArchiveImported(path)
    Palace_Notify("Imported " . merged . " atom(s)", 2000, BANNER_ACCENT_SUCCESS)
    Palace_ShowImportMenu()
    return true
}

Palace_ImportBeastsFromDesktop(path := "") {
    if (path = "")
        path := Palace_DesktopNewest("PALACE_BEASTS*.csv")
    if (path = "")
        path := Palace_DesktopNewest("PALACE_BEASTS*.txt")
    if (path = "") {
        Palace_Notify("No PALACE_BEASTS*.csv on Desktop", 2200, BANNER_ACCENT_ERROR)
        return false
    }
    rows := Palace_ReadAiImportCsv(path, "peg_code")
    if (!rows.Length) {
        Palace_Notify("No beast rows found", 2000, BANNER_ACCENT_ERROR)
        return false
    }
    labels := []
    for r in rows
        labels.Push("[" . (r.Has("peg_code") ? r["peg_code"] : "?") . "] "
        . (r.Has("beast_name") ? r["beast_name"] : ""))
    if (!Palace_ImportConfirmPreview("Import " . rows.Length . " beast(s)?", labels))
        return false
    beasts := Palace_Load("beasts")
    palaces := Palace_Load("palaces")
    merged := 0
    for r in rows {
        r := Palace_NormalizeBeastImportRow(r)
        palaceId := r.Has("palace_id") ? Trim(r["palace_id"]) : ""
        if (palaceId = "" || !Palace_FindById(palaces, palaceId)) {
            Palace_Notify("Skip: unknown palace_id", 2000, BANNER_ACCENT_ERROR)
            continue
        }
        peg := r.Has("peg_code") ? Trim(r["peg_code"]) : ""
        name := r.Has("beast_name") ? Trim(r["beast_name"]) : ""
        if (peg = "" || name = "")
            continue
        id := r.Has("id") && Trim(r["id"]) != "" ? Trim(r["id"])
            : "BEAST_" . Palace_Slug(palaceId) . "_" . Palace_Slug(peg)
        if (Palace_FindById(beasts, id))
            id := Palace_SlugId("BEAST_", peg . "_" . name, beasts)
        beasts.Push(Map(
            "id", id,
            "palace_id", palaceId,
            "peg_code", peg,
            "beast_name", name,
            "beast_source", r.Has("beast_source") ? r["beast_source"] : "",
            "sensory_channel", r.Has("sensory_channel") ? r["sensory_channel"] : "",
            "is_smashed", r.Has("is_smashed") ? r["is_smashed"] : "0",
            "sort_order", r.Has("sort_order") ? r["sort_order"] : "1"
        ))
        merged += 1
    }
    Palace_Save("beasts", beasts)
    Palace_ArchiveImported(path)
    Palace_Notify("Imported " . merged . " beast(s)", 2000, BANNER_ACCENT_SUCCESS)
    Palace_ShowImportMenu()
    return true
}

Palace_ImportPalacesFromDesktop(path := "") {
    if (path = "")
        path := Palace_DesktopNewest("PALACE_PALACES*.csv")
    if (path = "")
        path := Palace_DesktopNewest("PALACE_STREETS*.csv")
    if (path = "")
        path := Palace_DesktopNewest("PALACE_PALACES*.txt")
    if (path = "")
        path := Palace_DesktopNewest("PALACE_STREETS*.txt")
    if (path = "") {
        Palace_Notify("No PALACE_PALACES*.csv on Desktop", 2200, BANNER_ACCENT_ERROR)
        return false
    }
    rows := Palace_ReadAiImportCsv(path, "palace_number")
    if (!rows.Length)
        rows := Palace_ReadAiImportCsv(path, "street_number")
    if (!rows.Length) {
        Palace_Notify("No palace rows found", 2000, BANNER_ACCENT_ERROR)
        return false
    }
    labels := []
    for r in rows {
        r := Palace_NormalizePalaceImportRow(r)
        labels.Push("Memory Palace " . (r.Has("palace_number") ? r["palace_number"] : "?")
        . ": " . (r.Has("title") ? r["title"] : ""))
    }
    if (!Palace_ImportConfirmPreview("Import " . rows.Length . " Memory Palace(s)?", labels))
        return false
    palaces := Palace_Load("palaces")
    studies := Palace_Load("studies")
    merged := 0
    for r in rows {
        r := Palace_NormalizePalaceImportRow(r)
        studyId := r.Has("study_id") ? Trim(r["study_id"]) : ""
        if (studyId = "" || !Palace_FindById(studies, studyId)) {
            Palace_Notify("Skip: unknown study_id", 2000, BANNER_ACCENT_ERROR)
            continue
        }
        num := r.Has("palace_number") ? Trim(r["palace_number"]) : ""
        title := r.Has("title") ? Trim(r["title"]) : ""
        if (num = "" || title = "")
            continue
        pad := Format("{:02d}", Integer(num))
        id := r.Has("id") && Trim(r["id"]) != "" ? Trim(r["id"])
            : "PALACE_" . Palace_Slug(studyId) . "_" . pad
        if (SubStr(id, 1, 7) = "STREET_")
            id := "PALACE_" . SubStr(id, 8)
        if (Palace_FindById(palaces, id))
            id := Palace_SlugId("PALACE_", studyId . "_" . num, palaces)
        palaces.Push(Map(
            "id", id,
            "study_id", studyId,
            "palace_number", num,
            "title", title,
            "character_name", r.Has("character_name") ? r["character_name"] : "",
            "image_rel_path", r.Has("image_rel_path") ? r["image_rel_path"] : "",
            "depth_slots_used", r.Has("depth_slots_used") ? r["depth_slots_used"] : "0",
            "image_prompt", r.Has("image_prompt") ? r["image_prompt"] : ""
        ))
        merged += 1
    }
    Palace_Save("palaces", palaces)
    Palace_ArchiveImported(path)
    Palace_Notify("Imported " . merged . " Memory Palace(s)", 2000, BANNER_ACCENT_SUCCESS)
    Palace_ShowImportMenu()
    return true
}

Palace_ImportAutoFromDesktop(*) {
    path := Palace_DesktopNewest("PALACE_ATOMS*.csv")
    if (path != "")
        return Palace_ImportAtomsFromDesktop(path)
    path := Palace_DesktopNewest("PALACE_BEASTS*.csv")
    if (path != "")
        return Palace_ImportBeastsFromDesktop(path)
    path := Palace_DesktopNewest("PALACE_PALACES*.csv")
    if (path != "")
        return Palace_ImportPalacesFromDesktop(path)
    path := Palace_DesktopNewest("PALACE_STREETS*.csv")
    if (path != "")
        return Palace_ImportPalacesFromDesktop(path)
    path := Palace_DesktopNewest("PALACE_*.csv")
    if (path = "")
        path := Palace_DesktopNewest("PALACE_*.txt")
    if (path = "") {
        Palace_Notify("No PALACE_*.csv on Desktop", 2200, BANNER_ACCENT_ERROR)
        return false
    }
    text := StrLower(Palace_ReadUtf8(path))
    if (InStr(text, "beast_id") && InStr(text, "kind"))
        return Palace_ImportAtomsFromDesktop(path)
    if (InStr(text, "peg_code") && (InStr(text, "palace_id") || InStr(text, "street_id")))
        return Palace_ImportBeastsFromDesktop(path)
    if ((InStr(text, "palace_number") || InStr(text, "street_number")) && InStr(text, "study_id"))
        return Palace_ImportPalacesFromDesktop(path)
    Palace_Notify("Could not detect CSV type", 2200, BANNER_ACCENT_ERROR)
    return false
}
