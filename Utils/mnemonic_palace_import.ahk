; =============================================================================
; Utils module: mnemonic_palace_import.ahk
; Import AI-generated PALACE_*.csv pack from Desktop + quick image attach
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

Palace_DesktopNewestCsvOrTxt(baseName) {
    path := Palace_DesktopNewest(baseName . "*.csv")
    if (path = "")
        path := Palace_DesktopNewest(baseName . "*.txt")
    return path
}

Palace_DesktopNewestImage() {
    newest := ""
    newestTime := 0
    for pat in ["*.png", "*.jpg", "*.jpeg"] {
        loop files A_Desktop . "\" . pat, "F" {
            ts := Number(A_LoopFileTimeModified)
            if (ts > newestTime) {
                newestTime := ts
                newest := A_LoopFileFullPath
            }
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
    lv := g.Add("ListView", "w720 h420 Grid", ["Row"])
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

Palace_ResolveDesktopPalacesPath() {
    path := Palace_DesktopNewestCsvOrTxt("PALACE_PALACES")
    if (path = "")
        path := Palace_DesktopNewestCsvOrTxt("PALACE_STREETS")
    return path
}

; Newest PALACE_PACK*.txt/csv, else newest gemini-code*.txt that contains pack FILE markers.
Palace_DesktopNewestPackPath() {
    path := Palace_DesktopNewestCsvOrTxt("PALACE_PACK")
    if (path != "")
        return path
    newest := ""
    newestTime := 0
    loop files A_Desktop . "\gemini-code*.txt", "F" {
        text := Palace_ReadUtf8(A_LoopFileFullPath)
        if (!InStr(text, "===FILE: PALACE_", false) && !InStr(text, "---FILE: PALACE_", false))
            continue
        ts := Number(A_LoopFileTimeModified)
        if (ts > newestTime) {
            newestTime := ts
            newest := A_LoopFileFullPath
        }
    }
    return newest
}

; Extract pure CSV body between ===FILE: name=== / ---FILE: name--- and matching END_FILE.
Palace_ExtractPackFileSection(text, fileName) {
    for style in ["===", "---"] {
        needle := style . "FILE: " . fileName . style
        endNeedle := style . "END_FILE" . style
        pos := InStr(text, needle, false)
        if (!pos)
            continue
        rest := SubStr(text, pos + StrLen(needle))
        endPos := InStr(rest, endNeedle, false)
        if (!endPos)
            continue
        return Trim(SubStr(rest, 1, endPos - 1), "`r`n `t")
    }
    return ""
}

; Extract PREVIEW block (=== or ---). Returns plain text body or "".
Palace_ExtractPackPreview(text) {
    for style in ["===", "---"] {
        needle := style . "PREVIEW" . style
        endNeedle := style . "END_PREVIEW" . style
        pos := InStr(text, needle, false)
        if (!pos)
            continue
        rest := SubStr(text, pos + StrLen(needle))
        endPos := InStr(rest, endNeedle, false)
        if (!endPos)
            continue
        return Trim(SubStr(rest, 1, endPos - 1), "`r`n `t")
    }
    return ""
}

; Split a PALACE_PACK / gemini-code body into row arrays. Requires all three FILE sections.
Palace_SplitPalacePack(path) {
    result := Map("ok", false, "error", "", "preview", "", "palaces", [], "beasts", [], "atoms", [])
    text := Palace_ReadUtf8(path)
    if (text = "") {
        result["error"] := "Empty pack file"
        return result
    }
    result["preview"] := Palace_ExtractPackPreview(text)
    specs := [
        ["palaces", "PALACE_PALACES.csv", "palace_number"],
        ["beasts", "PALACE_BEASTS.csv", "peg_code"],
        ["atoms", "PALACE_ATOMS.csv", "beast_id"]
    ]
    for spec in specs {
        key := spec[1]
        fname := spec[2]
        hint := spec[3]
        body := Palace_ExtractPackFileSection(text, fname)
        if (body = "") {
            result["error"] := "Missing section " . fname
            return result
        }
        tmp := A_Temp . "\palace_pack_" . key . ".csv"
        Palace_WriteUtf8(tmp, body)
        rows := Palace_ReadAiImportCsv(tmp, hint)
        if (key = "palaces" && !rows.Length)
            rows := Palace_ReadAiImportCsv(tmp, "street_number")
        try FileDelete(tmp)
        catch {
        }
        result[key] := rows
    }
    result["ok"] := true
    return result
}

; Single [I]: newest Desktop pack (palaces → beasts → atoms), combined preview, palace-scoped replace.
Palace_ImportMnemonicsFromDesktop(*) {
    Palace_EnsureData()
    pathPalaces := Palace_ResolveDesktopPalacesPath()
    pathBeasts := Palace_DesktopNewestCsvOrTxt("PALACE_BEASTS")
    pathAtoms := Palace_DesktopNewestCsvOrTxt("PALACE_ATOMS")
    pathPack := ""
    packPreview := ""
    palaceRows := []
    beastRows := []
    atomRows := []

    if (pathPalaces = "" && pathBeasts = "" && pathAtoms = "") {
        pathPack := Palace_DesktopNewestPackPath()
        if (pathPack = "") {
            Palace_Notify("No PALACE_PACK / PALACE_*.csv on Desktop", 2800, BANNER_ACCENT_ERROR)
            Palace_ShowMainMenu()
            return false
        }
        split := Palace_SplitPalacePack(pathPack)
        if (!split["ok"]) {
            Palace_Notify(split["error"], 2800, BANNER_ACCENT_ERROR)
            Palace_ShowMainMenu()
            return false
        }
        palaceRows := split["palaces"]
        beastRows := split["beasts"]
        atomRows := split["atoms"]
        packPreview := split["preview"]
    } else {
        if (pathPalaces != "") {
            palaceRows := Palace_ReadAiImportCsv(pathPalaces, "palace_number")
            if (!palaceRows.Length)
                palaceRows := Palace_ReadAiImportCsv(pathPalaces, "street_number")
        }
        if (pathBeasts != "")
            beastRows := Palace_ReadAiImportCsv(pathBeasts, "peg_code")
        if (pathAtoms != "")
            atomRows := Palace_ReadAiImportCsv(pathAtoms, "beast_id")
    }

    if (!palaceRows.Length && !beastRows.Length && !atomRows.Length) {
        Palace_Notify("No rows found in Desktop PALACE pack", 2500, BANNER_ACCENT_ERROR)
        Palace_ShowMainMenu()
        return false
    }

    labels := []
    if (packPreview != "") {
        labels.Push("--- Human PREVIEW ---")
        for line in StrSplit(StrReplace(StrReplace(packPreview, "`r`n", "`n"), "`r", "`n"), "`n")
            labels.Push(line)
        labels.Push("--- Import rows ---")
    }
    if (palaceRows.Length) {
        labels.Push("--- Palaces (" . palaceRows.Length . ") ---")
        for r in palaceRows {
            r := Palace_NormalizePalaceImportRow(r)
            labels.Push("Memory Palace " . (r.Has("palace_number") ? r["palace_number"] : "?")
            . ": " . (r.Has("title") ? r["title"] : ""))
        }
    }
    if (beastRows.Length) {
        labels.Push("--- Beasts (" . beastRows.Length . ") ---")
        for r in beastRows {
            r := Palace_NormalizeBeastImportRow(r)
            labels.Push("[" . (r.Has("peg_code") ? r["peg_code"] : "?") . "] "
            . (r.Has("beast_name") ? r["beast_name"] : "")
            . " @ " . (r.Has("palace_id") ? r["palace_id"] : "?"))
        }
    }
    if (atomRows.Length) {
        labels.Push("--- Atoms (" . atomRows.Length . ") ---")
        for r in atomRows {
            lab := (r.Has("kind") ? r["kind"] : "?") . " · "
            . (r.Has("zone") ? r["zone"] : "") . " · "
            . (r.Has("concept") ? SubStr(r["concept"], 1, 50)
                : (r.Has("context") ? SubStr(r["context"], 1, 50) : ""))
            labels.Push(lab)
        }
    }

    title := "Import pack?  "
        . palaceRows.Length . " palace(s)  ·  "
        . beastRows.Length . " beast(s)  ·  "
        . atomRows.Length . " atom(s)"
    if (!Palace_ImportConfirmPreview(title, labels)) {
        Palace_ShowMainMenu()
        return false
    }

    studies := Palace_Load("studies")
    palaces := Palace_Load("palaces")
    beasts := Palace_Load("beasts")
    atoms := Palace_Load("atoms")
    nPalaces := 0
    nBeasts := 0
    nAtoms := 0

    ; --- Palaces: upsert by id ---
    for r in palaceRows {
        r := Palace_NormalizePalaceImportRow(r)
        studyId := r.Has("study_id") ? Trim(r["study_id"]) : ""
        if (studyId = "" || !Palace_FindById(studies, studyId)) {
            Palace_Notify("Skip palace: unknown study_id " . studyId, 2200, BANNER_ACCENT_ERROR)
            continue
        }
        num := r.Has("palace_number") ? Trim(r["palace_number"]) : ""
        titleP := r.Has("title") ? Trim(r["title"]) : ""
        if (num = "" || titleP = "")
            continue
        pad := Format("{:02d}", Integer(num))
        id := r.Has("id") && Trim(r["id"]) != "" ? Trim(r["id"])
            : "PALACE_" . Palace_Slug(studyId) . "_" . pad
        if (SubStr(id, 1, 7) = "STREET_")
            id := "PALACE_" . SubStr(id, 8)
        img := r.Has("image_rel_path") ? Trim(r["image_rel_path"]) : ""
        existing := Palace_FindById(palaces, id)
        if (IsObject(existing)) {
            existing["study_id"] := studyId
            existing["palace_number"] := num
            existing["title"] := titleP
            existing["character_name"] := r.Has("character_name") ? r["character_name"] : ""
            existing["depth_slots_used"] := r.Has("depth_slots_used") ? r["depth_slots_used"] : "0"
            existing["image_prompt"] := r.Has("image_prompt") ? r["image_prompt"] : ""
            if (img != "")
                existing["image_rel_path"] := img
            nPalaces += 1
        } else {
            palaces.Push(Map(
                "id", id,
                "study_id", studyId,
                "palace_number", num,
                "title", titleP,
                "character_name", r.Has("character_name") ? r["character_name"] : "",
                "image_rel_path", img,
                "depth_slots_used", r.Has("depth_slots_used") ? r["depth_slots_used"] : "0",
                "image_prompt", r.Has("image_prompt") ? r["image_prompt"] : ""
            ))
            nPalaces += 1
        }
    }

    ; --- Beasts: replace all beasts (and cascade atoms) for palace_ids in pack ---
    if (beastRows.Length) {
        replacePalaceIds := Map()
        for r in beastRows {
            r := Palace_NormalizeBeastImportRow(r)
            pid := r.Has("palace_id") ? Trim(r["palace_id"]) : ""
            if (pid != "")
                replacePalaceIds[pid] := true
        }
        removeBeastIds := Map()
        for b in beasts {
            if (replacePalaceIds.Has(b["palace_id"]))
                removeBeastIds[b["id"]] := true
        }
        newAtoms := []
        for a in atoms {
            if (!removeBeastIds.Has(a["beast_id"]))
                newAtoms.Push(a)
        }
        atoms := newAtoms
        newBeasts := []
        for b in beasts {
            if (!replacePalaceIds.Has(b["palace_id"]))
                newBeasts.Push(b)
        }
        beasts := newBeasts

        for r in beastRows {
            r := Palace_NormalizeBeastImportRow(r)
            palaceId := r.Has("palace_id") ? Trim(r["palace_id"]) : ""
            if (palaceId = "" || !Palace_FindById(palaces, palaceId)) {
                Palace_Notify("Skip beast: unknown palace_id " . palaceId, 2200, BANNER_ACCENT_ERROR)
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
            nBeasts += 1
        }
    }

    ; --- Atoms: replace atoms for beast_ids in pack ---
    if (atomRows.Length) {
        replaceBeastIds := Map()
        for r in atomRows {
            bid := r.Has("beast_id") ? Trim(r["beast_id"]) : ""
            if (bid != "")
                replaceBeastIds[bid] := true
        }
        newAtoms2 := []
        for a in atoms {
            if (!replaceBeastIds.Has(a["beast_id"]))
                newAtoms2.Push(a)
        }
        atoms := newAtoms2

        usedAtomIds := Map()
        scratchAtoms := []
        for a in atoms {
            usedAtomIds[a["id"]] := true
            scratchAtoms.Push(a)
        }
        pendingByBeast := Map()
        for r in atomRows {
            beastId := r.Has("beast_id") ? Trim(r["beast_id"]) : ""
            if (beastId = "" || !Palace_FindById(beasts, beastId)) {
                Palace_Notify("Skip atom: unknown beast_id " . beastId, 2200, BANNER_ACCENT_ERROR)
                continue
            }
            kind := r.Has("kind") ? StrLower(Trim(r["kind"])) : "single"
            if (kind = "")
                kind := "single"
            if (kind = "subtopic")
                kind := "zoned"
            id := r.Has("id") && Trim(r["id"]) != "" ? Trim(r["id"]) : ""
            if (id = "" || usedAtomIds.Has(id))
                id := Palace_NextId("ATOM_", scratchAtoms, 4)
            usedAtomIds[id] := true
            scratchAtoms.Push(Map("id", id))
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
            if (!pendingByBeast.Has(beastId))
                pendingByBeast[beastId] := []
            pendingByBeast[beastId].Push(row)
        }
        for beastId, proposed in pendingByBeast {
            err := Palace_ValidateBeastAtoms(beastId, proposed)
            if (err != "") {
                Palace_Notify("Skip atoms for " . beastId . ": " . err, 2800, BANNER_ACCENT_ERROR)
                continue
            }
            for row in proposed {
                atoms.Push(row)
                nAtoms += 1
            }
        }
    }

    if (palaceRows.Length)
        Palace_Save("palaces", palaces)
    if (beastRows.Length) {
        Palace_Save("beasts", beasts)
        Palace_Save("atoms", atoms)
    } else if (atomRows.Length) {
        Palace_Save("atoms", atoms)
    }

    if (pathPack != "")
        Palace_ArchiveImported(pathPack)
    else {
        if (pathPalaces != "")
            Palace_ArchiveImported(pathPalaces)
        if (pathBeasts != "")
            Palace_ArchiveImported(pathBeasts)
        if (pathAtoms != "")
            Palace_ArchiveImported(pathAtoms)
    }

    Palace_Notify("Imported " . nPalaces . " palace(s), " . nBeasts . " beast(s), " . nAtoms . " atom(s)",
        2800, BANNER_ACCENT_SUCCESS)
    Palace_ShowMainMenu()
    return true
}

Palace_LastPalaceForQuickImage() {
    lastStudy := Trim(Palace_Setting("General", "LastStudyId", ""))
    palaces := Palace_Load("palaces")
    best := Palace_PickHighestPalace(palaces, lastStudy)
    if (!IsObject(best) && lastStudy != "")
        best := Palace_PickHighestPalace(palaces, "")
    return best
}

Palace_PickHighestPalace(palaces, studyId) {
    best := false
    bestNum := -1
    bestId := ""
    for p in palaces {
        if (studyId != "" && p["study_id"] != studyId)
            continue
        try
            num := Integer(p["palace_number"])
        catch
            continue
        pid := p.Has("id") ? p["id"] : ""
        if (num > bestNum || (num = bestNum && pid > bestId)) {
            best := p
            bestNum := num
            bestId := pid
        }
    }
    return best
}

; [Q] Copy newest Desktop PNG/JPG onto the last Memory Palace under LastStudyId.
Palace_QuickAttachDesktopImage(*) {
    Palace_EnsureData()
    src := Palace_DesktopNewestImage()
    if (src = "") {
        Palace_Notify("No PNG/JPG on Desktop", 2200, BANNER_ACCENT_ERROR)
        return false
    }
    palace := Palace_LastPalaceForQuickImage()
    if (!IsObject(palace)) {
        Palace_Notify("No Memory Palace to attach to", 2200, BANNER_ACCENT_ERROR)
        return false
    }
    studies := Palace_Load("studies")
    study := Palace_FindById(studies, palace["study_id"])
    if (!IsObject(study)) {
        Palace_Notify("Study missing for palace", 2200, BANNER_ACCENT_ERROR)
        return false
    }
    slug := Trim(study["notes_rel_path"])
    if (slug = "") {
        Palace_Notify("Study has no notes path", 2200, BANNER_ACCENT_ERROR)
        return false
    }
    if (!Palace_EnsureStudyNotesFolder(slug)) {
        Palace_Notify("Could not create study images folder", 2500, BANNER_ACCENT_ERROR)
        return false
    }
    SplitPath(src, , , &ext)
    ext := StrLower(Trim(ext))
    if (ext = "")
        ext := "png"
    rel := StrReplace(slug, "\", "/") . "/images/" . palace["palace_number"] . "." . ext
    root := Palace_NotesStudiesRoot(true)
    if (root = "") {
        Palace_Notify("Notes studies root not found", 2200, BANNER_ACCENT_ERROR)
        return false
    }
    absDest := root . "\" . StrReplace(rel, "/", "\")
    try FileCopy(src, absDest, 1)
    catch as e {
        Palace_Notify("Copy failed: " . e.Message, 2800, BANNER_ACCENT_ERROR)
        return false
    }
    palaces := Palace_Load("palaces")
    existing := Palace_FindById(palaces, palace["id"])
    if (!IsObject(existing)) {
        Palace_Notify("Palace row vanished", 2000, BANNER_ACCENT_ERROR)
        return false
    }
    existing["image_rel_path"] := rel
    Palace_Save("palaces", palaces)
    Palace_Notify("Attached image → " . palace["title"] . " (" . rel . ")", 2800, BANNER_ACCENT_SUCCESS)
    return true
}
