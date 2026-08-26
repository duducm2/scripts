; =============================================================================
; Utils module: mnemonic_palace_import.ahk
; Import AI-generated PALACE_PACK / PALACE_*.txt|.csv from Desktop + quick image attach
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
    newest := ""
    newestTime := 0
    for ext in ["csv", "txt"] {
        loop files A_Desktop . "\" . baseName . "*." . ext, "F" {
            ts := Number(A_LoopFileTimeModified)
            if (ts > newestTime) {
                newestTime := ts
                newest := A_LoopFileFullPath
            }
        }
    }
    return newest
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
; When skipNotes is an Array, malformed rows (field count ≠ header) are skipped and noted.
Palace_ReadAiImportCsv(path, headerHint := "beast_id", skipNotes := 0) {
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
    if (!start) {
        if (IsObject(skipNotes))
            return Palace_ReadCsv(path, true, skipNotes)
        return Palace_ReadCsv(path)
    }
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
    if (IsObject(skipNotes))
        rows := Palace_ReadCsv(tmp, true, skipNotes)
    else
        rows := Palace_ReadCsv(tmp)
    try FileDelete(tmp)
    catch {
    }
    return rows
}

Palace_ImportNoteSkip(skipCounts, reason) {
    r := Trim(reason)
    if (r = "")
        return
    if (!skipCounts.Has(r))
        skipCounts[r] := 0
    skipCounts[r] += 1
}

Palace_ImportSkipSummary(skipCounts, csvNotes := 0) {
    parts := []
    if (IsObject(csvNotes)) {
        for note in csvNotes
            parts.Push(note)
    }
    for reason, n in skipCounts
        parts.Push(n . "× " . reason)
    if (!parts.Length)
        return ""
    out := parts[1]
    loop parts.Length - 1
        out .= "; " . parts[A_Index + 1]
    if (StrLen(out) > 180)
        out := SubStr(out, 1, 177) . "..."
    return out
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

; Resolve import row. On create, pushes a stub onto palaces so later rows get the next number.
; Returns packId (original pack id) so caller can build palaceIdRemap when it differs from canon id.
Palace_ResolvePalaceImportRow(r, studies, &palaces) {
    r := Palace_NormalizePalaceImportRow(r)
    rawStudy := r.Has("study_id") ? Trim(r["study_id"]) : ""
    studyId := Palace_ResolveStudyId(rawStudy, studies)
    if (studyId = "")
        return Map("ok", false, "skip", "unknown study_id " . rawStudy)
    r["study_id"] := studyId
    titleP := r.Has("title") ? Trim(r["title"]) : ""
    if (titleP = "")
        return Map("ok", false, "skip", "missing title")

    packId := r.Has("id") ? Trim(r["id"]) : ""
    if (SubStr(packId, 1, 7) = "STREET_")
        packId := "PALACE_" . SubStr(packId, 8)

    existing := false
    if (packId != "") {
        existing := Palace_FindById(palaces, packId)
        if (!IsObject(existing)) {
            rewritten := Palace_RewritePalaceIdForStudy(packId, studyId)
            if (rewritten != "" && rewritten != packId)
                existing := Palace_FindById(palaces, rewritten)
        }
        if (!IsObject(existing)) {
            for p in palaces {
                if (p.Has("study_id") && p["study_id"] = studyId && Palace_PalaceIdsSoftEqual(packId, p["id"])) {
                    existing := p
                    break
                }
            }
        }
    }
    if (!IsObject(existing))
        existing := Palace_FindPalaceByTitle(palaces, studyId, titleP)

    num := ""
    id := ""
    if (IsObject(existing)) {
        num := String(existing["palace_number"])
        id := existing["id"]
    } else {
        packNum := r.Has("palace_number") ? Trim(r["palace_number"]) : ""
        usePackNum := false
        if (packNum != "") {
            try {
                Integer(packNum)
                exceptId := packId
                rewritten := Palace_RewritePalaceIdForStudy(packId, studyId)
                if (rewritten != "")
                    exceptId := rewritten
                if (!Palace_PalaceNumberInUse(palaces, studyId, packNum, exceptId))
                    usePackNum := true
            } catch {
            }
        }
        if (usePackNum)
            num := String(Integer(packNum))
        else
            num := String(Palace_NextPalaceNumber(palaces, studyId))
        id := Palace_CanonPalaceId(studyId, num)
        if (Palace_IdExists(palaces, id))
            id := Palace_SlugId("PALACE_", Palace_StudySlugFromId(studyId) . "_" . num, palaces)
        palaces.Push(Map(
            "id", id,
            "study_id", studyId,
            "palace_number", num,
            "title", titleP,
            "character_name", "",
            "image_rel_path", "",
            "depth_slots_used", "0",
            "image_prompt", "",
            "_import_stub", "1"
        ))
        existing := false
    }
    return Map("ok", true, "skip", "", "studyId", studyId, "title", titleP, "num", num, "id", id,
        "packId", packId, "existing", existing, "row", r)
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
; If END_FILE is missing, salvages until the next FILE marker or EOF (&salvaged set true).
; Also matches fileName without a trailing .csv (e.g. PALACE_ATOMS).
Palace_ExtractPackFileSection(text, fileName, &salvaged := false) {
    salvaged := false
    names := [fileName]
    if (RegExMatch(fileName, "i)\.csv$"))
        names.Push(RegExReplace(fileName, "i)\.csv$", ""))
    else
        names.Push(fileName . ".csv")
    for name in names {
        for style in ["===", "---"] {
            needle := style . "FILE: " . name . style
            endNeedle := style . "END_FILE" . style
            pos := InStr(text, needle, false)
            if (!pos)
                continue
            rest := SubStr(text, pos + StrLen(needle))
            endPos := InStr(rest, endNeedle, false)
            if (endPos)
                return Trim(SubStr(rest, 1, endPos - 1), "`r`n `t")
            ; Truncated pack: cut at next FILE section (either style) or take EOF.
            nextPos := 0
            for style2 in ["===", "---"] {
                n2 := InStr(rest, "`n" . style2 . "FILE:", false)
                if (n2 && (!nextPos || n2 < nextPos))
                    nextPos := n2
            }
            salvaged := true
            if (nextPos)
                return Trim(SubStr(rest, 1, nextPos - 1), "`r`n `t")
            return Trim(rest, "`r`n `t")
        }
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
; result["csvNotes"] lists malformed CSV rows skipped under strict field-count checks.
Palace_SplitPalacePack(path) {
    result := Map("ok", false, "error", "", "preview", "", "palaces", [], "beasts", [], "atoms", [], "csvNotes", [])
    text := Palace_ReadUtf8(path)
    if (text = "") {
        result["error"] := "Empty pack file"
        return result
    }
    result["preview"] := Palace_ExtractPackPreview(text)
    csvNotes := []
    specs := [
        ["palaces", "PALACE_PALACES.csv", "palace_number"],
        ["beasts", "PALACE_BEASTS.csv", "peg_code"],
        ["atoms", "PALACE_ATOMS.csv", "beast_id"]
    ]
    for spec in specs {
        key := spec[1]
        fname := spec[2]
        hint := spec[3]
        sectionSalvaged := false
        body := Palace_ExtractPackFileSection(text, fname, &sectionSalvaged)
        if (body = "") {
            result["error"] := "Missing section " . fname
            return result
        }
        if (sectionSalvaged)
            csvNotes.Push(fname . ": missing END_FILE; salvaged until next section/EOF")
        tmp := A_Temp . "\palace_pack_" . key . ".csv"
        Palace_WriteUtf8(tmp, body)
        sectionNotes := []
        rows := Palace_ReadAiImportCsv(tmp, hint, sectionNotes)
        if (key = "palaces" && !rows.Length)
            rows := Palace_ReadAiImportCsv(tmp, "street_number", sectionNotes)
        for note in sectionNotes
            csvNotes.Push(fname . ": " . note)
        try FileDelete(tmp)
        catch {
        }
        if (!rows.Length) {
            result["csvNotes"] := csvNotes
            result["error"] := fname . " has no valid rows (pack may be truncated)"
            return result
        }
        result[key] := rows
    }
    packCheck := Palace_ValidatePackBeastPacking(result["palaces"], result["beasts"])
    if (!packCheck["ok"]) {
        result["csvNotes"] := csvNotes
        result["error"] := packCheck["error"]
        return result
    }
    result["csvNotes"] := csvNotes
    result["ok"] := true
    return result
}

; Reject packs where any non-final palace (by palace_number) has 1–4 beasts.
; Only the highest-numbered palace in the pack (per study) may be under-filled.
Palace_ValidatePackBeastPacking(palaceRows, beastRows) {
    out := Map("ok", true, "error", "")
    if (!IsObject(palaceRows) || !IsObject(beastRows) || !palaceRows.Length || !beastRows.Length)
        return out
    metaById := Map()
    for p in palaceRows {
        pid := Trim(p.Has("id") ? p["id"] : "")
        if (pid = "")
            continue
        num := 0
        try num := Integer(p.Has("palace_number") ? p["palace_number"] : 0)
        catch {
            num := 0
        }
        studyId := Trim(p.Has("study_id") ? p["study_id"] : "")
        metaById[pid] := Map("num", num, "study", studyId, "title", Trim(p.Has("title") ? p["title"] : pid))
    }
    counts := Map()
    for b in beastRows {
        pid := Trim(b.Has("palace_id") ? b["palace_id"] : "")
        if (pid = "")
            continue
        if (!counts.Has(pid))
            counts[pid] := 0
        counts[pid] += 1
    }
    ; studyId -> { maxNum, pattern parts sorted later }
    byStudy := Map()
    for pid, meta in metaById {
        nBeasts := counts.Has(pid) ? counts[pid] : 0
        if (nBeasts <= 0)
            continue
        if (nBeasts > 5) {
            out["ok"] := false
            out["error"] := "Packing error: " . meta["title"] . " has " . nBeasts
                . " beasts (max 5)"
            return out
        }
        sid := meta["study"]
        if (sid = "")
            sid := "_"
        if (!byStudy.Has(sid))
            byStudy[sid] := Map("maxNum", -1, "rows", [])
        if (meta["num"] > byStudy[sid]["maxNum"])
            byStudy[sid]["maxNum"] := meta["num"]
        byStudy[sid]["rows"].Push(Map("pid", pid, "num", meta["num"], "n", nBeasts, "title", meta["title"]))
    }
    for sid, info in byStudy {
        maxNum := info["maxNum"]
        pattern := []
        ; Sort rows by palace_number ascending for error pattern
        rows := info["rows"]
        loop {
            swapped := false
            i := 1
            while (i < rows.Length) {
                if (rows[i]["num"] > rows[i + 1]["num"]) {
                    tmp := rows[i]
                    rows[i] := rows[i + 1]
                    rows[i + 1] := tmp
                    swapped := true
                }
                i += 1
            }
            if (!swapped)
                break
        }
        for row in rows {
            pattern.Push(row["n"])
            if (row["n"] < 5 && row["num"] < maxNum) {
                patStr := ""
                for n in pattern
                    patStr .= (patStr = "" ? "" : "+") . n
                ; finish pattern for message
                j := pattern.Length + 1
                while (j <= rows.Length) {
                    patStr .= "+" . rows[j]["n"]
                    j += 1
                }
                out["ok"] := false
                out["error"] := "Packing error: palace " . row["num"] . " (" . row["title"]
                . ") has " . row["n"] . " beast(s) but later palaces exist — only the last"
                . " palace may have <5 (got " . patStr . "; need 5+5+…+remainder on last)"
                return out
            }
        }
    }
    return out
}

; Resolve newest Desktop PLAN_PACK (or gemini-code with PLANS.csv). Empty if none.
Palace_ResolveDesktopPlanPackPath() {
    pathPlanPack := Palace_DesktopNewestCsvOrTxt("PLAN_PACK")
    if (pathPlanPack != "")
        return pathPlanPack
    newest := ""
    newestTime := 0
    loop files A_Desktop . "\gemini-code*.txt", "F" {
        text := Palace_ReadUtf8(A_LoopFileFullPath)
        if (!InStr(text, "===FILE: PLANS.csv", false) && !InStr(text, "---FILE: PLANS.csv", false))
            continue
        ts := Number(A_LoopFileTimeModified)
        if (ts > newestTime) {
            newestTime := ts
            newest := A_LoopFileFullPath
        }
    }
    return newest
}

; Main menu [J]: PLAN_PACK only.
Palace_ImportPlanPackFromDesktop(*) {
    Palace_EnsureData()
    pathPlanPack := Palace_ResolveDesktopPlanPackPath()
    if (pathPlanPack = "") {
        Palace_Notify("No PLAN_PACK / gemini-code (PLANS.csv) on Desktop", 2800, BANNER_ACCENT_ERROR)
        Palace_ShowMainMenu()
        return false
    }
    return Palace_ImportPlansFromDesktop(pathPlanPack)
}

; Main menu [I]: mnemonic PALACE packs only (never PLAN_PACK).
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
    csvNotes := []

    if (pathPalaces = "" && pathBeasts = "" && pathAtoms = "") {
        pathPack := Palace_DesktopNewestPackPath()
        if (pathPack = "") {
            Palace_Notify("No PALACE_PACK / PALACE_*.txt|.csv on Desktop", 2800, BANNER_ACCENT_ERROR)
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
        if (split.Has("csvNotes"))
            csvNotes := split["csvNotes"]
    } else {
        if (pathPalaces != "") {
            notesP := []
            palaceRows := Palace_ReadAiImportCsv(pathPalaces, "palace_number", notesP)
            if (!palaceRows.Length)
                palaceRows := Palace_ReadAiImportCsv(pathPalaces, "street_number", notesP)
            for note in notesP
                csvNotes.Push("palaces: " . note)
        }
        if (pathBeasts != "") {
            notesB := []
            beastRows := Palace_ReadAiImportCsv(pathBeasts, "peg_code", notesB)
            for note in notesB
                csvNotes.Push("beasts: " . note)
        }
        if (pathAtoms != "") {
            notesA := []
            atomRows := Palace_ReadAiImportCsv(pathAtoms, "beast_id", notesA)
            for note in notesA
                csvNotes.Push("atoms: " . note)
        }
    }

    if (!palaceRows.Length && !beastRows.Length && !atomRows.Length) {
        summary := Palace_ImportSkipSummary(Map(), csvNotes)
        msg := "No rows found in Desktop PALACE pack"
        if (summary != "")
            msg .= " — " . summary
        Palace_Notify(msg, 3200, BANNER_ACCENT_ERROR)
        Palace_ShowMainMenu()
        return false
    }

    packCheck := Palace_ValidatePackBeastPacking(palaceRows, beastRows)
    if (!packCheck["ok"]) {
        Palace_Notify(packCheck["error"], 4500, BANNER_ACCENT_ERROR)
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
    if (csvNotes.Length) {
        labels.Push("--- CSV warnings ---")
        for note in csvNotes
            labels.Push(note)
    }
    studiesPreview := Palace_Load("studies")
    palacesPreview := Palace_Load("palaces")
    if (palaceRows.Length) {
        labels.Push("--- Palaces (" . palaceRows.Length . ") ---")
        for r in palaceRows {
            resolved := Palace_ResolvePalaceImportRow(r, studiesPreview, &palacesPreview)
            if (!resolved["ok"]) {
                labels.Push("Skip: " . resolved["skip"])
                continue
            }
            labels.Push("Memory Palace " . resolved["num"] . ": " . resolved["title"]
                . " [" . resolved["id"] . "]")
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
    palaceIdRemap := Map()
    beastIdRemap := Map()
    syncStudyIds := Map()
    skipCounts := Map()

    ; --- Palaces: upsert by id / title; auto-assign palace_number when needed ---
    for r in palaceRows {
        resolved := Palace_ResolvePalaceImportRow(r, studies, &palaces)
        if (!resolved["ok"]) {
            Palace_ImportNoteSkip(skipCounts, resolved["skip"])
            Palace_Notify("Skip palace: " . resolved["skip"], 2200, BANNER_ACCENT_ERROR)
            continue
        }
        studyId := resolved["studyId"]
        titleP := resolved["title"]
        num := resolved["num"]
        id := resolved["id"]
        packPalaceId := resolved.Has("packId") ? Trim(resolved["packId"]) : ""
        if (packPalaceId != "" && packPalaceId != id)
            palaceIdRemap[packPalaceId] := id
        palaceIdRemap[id] := id
        rowIn := resolved["row"]
        img := rowIn.Has("image_rel_path") ? Trim(rowIn["image_rel_path"]) : ""
        if (img != "") {
            absImg := Palace_ResolveImagePath(img)
            if (absImg = "" || !FileExist(absImg))
                img := ""
        }
        promptIn := rowIn.Has("image_prompt") ? Trim(rowIn["image_prompt"]) : ""
        existing := resolved["existing"]
        if (IsObject(existing)) {
            existing["study_id"] := studyId
            existing["palace_number"] := num
            existing["title"] := titleP
            if (rowIn.Has("character_name") && Trim(rowIn["character_name"]) != "")
                existing["character_name"] := rowIn["character_name"]
            if (rowIn.Has("depth_slots_used") && Trim(rowIn["depth_slots_used"]) != "")
                existing["depth_slots_used"] := rowIn["depth_slots_used"]
            if (promptIn != "")
                existing["image_prompt"] := promptIn
            if (img != "")
                existing["image_rel_path"] := img
            syncStudyIds[studyId] := true
            nPalaces += 1
        } else {
            ; Stub already pushed by Resolve — fill real fields (or replace stub).
            stub := Palace_FindById(palaces, id)
            if (IsObject(stub)) {
                stub["study_id"] := studyId
                stub["palace_number"] := num
                stub["title"] := titleP
                stub["character_name"] := rowIn.Has("character_name") ? rowIn["character_name"] : ""
                stub["image_rel_path"] := img
                stub["depth_slots_used"] := rowIn.Has("depth_slots_used") ? rowIn["depth_slots_used"] : "0"
                stub["image_prompt"] := promptIn
                if (stub.Has("_import_stub"))
                    stub.Delete("_import_stub")
            } else {
                palaces.Push(Map(
                    "id", id,
                    "study_id", studyId,
                    "palace_number", num,
                    "title", titleP,
                    "character_name", rowIn.Has("character_name") ? rowIn["character_name"] : "",
                    "image_rel_path", img,
                    "depth_slots_used", rowIn.Has("depth_slots_used") ? rowIn["depth_slots_used"] : "0",
                    "image_prompt", promptIn
                ))
            }
            syncStudyIds[studyId] := true
            nPalaces += 1
        }
    }

    ; Strip any leftover preview stubs (should not remain after create fill)
    cleanedPalaces := []
    for p in palaces {
        if (p.Has("_import_stub"))
            continue
        cleanedPalaces.Push(p)
    }
    palaces := cleanedPalaces

    ; --- Beasts: wipe only palaces with at least one insertable row; remap colliding ids ---
    if (beastRows.Length) {
        insertable := []
        replacePalaceIds := Map()
        for r in beastRows {
            r := Palace_NormalizeBeastImportRow(r)
            rawPalaceId := r.Has("palace_id") ? Trim(r["palace_id"]) : ""
            peg := r.Has("peg_code") ? Trim(r["peg_code"]) : ""
            name := r.Has("beast_name") ? Trim(r["beast_name"]) : ""
            palaceId := Palace_ResolvePalaceIdRef(rawPalaceId, palaces, palaceIdRemap)
            if (palaceId = "") {
                Palace_ImportNoteSkip(skipCounts, "unknown palace_id " . rawPalaceId)
                Palace_Notify("Skip beast: unknown palace_id " . rawPalaceId, 2200, BANNER_ACCENT_ERROR)
                continue
            }
            r["palace_id"] := palaceId
            if (peg = "" || name = "") {
                Palace_ImportNoteSkip(skipCounts, "beast missing peg/name")
                continue
            }
            insertable.Push(r)
            replacePalaceIds[palaceId] := true
        }
        if (replacePalaceIds.Count) {
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
        }
        for r in insertable {
            palaceId := Trim(r["palace_id"])
            peg := Trim(r["peg_code"])
            name := Trim(r["beast_name"])
            packId := r.Has("id") && Trim(r["id"]) != "" ? Trim(r["id"]) : ""
            id := packId != "" ? packId
                : "BEAST_" . Palace_Slug(palaceId) . "_" . Palace_Slug(peg)
            if (Palace_FindById(beasts, id))
                id := Palace_SlugId("BEAST_", peg . "_" . name, beasts)
            if (packId != "")
                beastIdRemap[packId] := id
            beastIdRemap[id] := id
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
            sid := Palace_StudyIdForPalace(palaceId)
            if (sid != "")
                syncStudyIds[sid] := true
            nBeasts += 1
        }
    }

    ; --- Atoms: validate groups first; wipe+insert only for groups that pass ---
    if (atomRows.Length) {
        usedAtomIds := Map()
        scratchAtoms := []
        for a in atoms {
            usedAtomIds[a["id"]] := true
            scratchAtoms.Push(a)
        }
        pendingByBeast := Map()
        for r in atomRows {
            beastId := r.Has("beast_id") ? Trim(r["beast_id"]) : ""
            if (beastId != "" && beastIdRemap.Has(beastId))
                beastId := beastIdRemap[beastId]
            if (beastId = "" || !Palace_FindById(beasts, beastId)) {
                Palace_ImportNoteSkip(skipCounts, "unknown beast_id " . beastId)
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
        okBeastIds := Map()
        for beastId, proposed in pendingByBeast {
            err := Palace_ValidateBeastAtoms(beastId, proposed)
            if (err != "") {
                Palace_ImportNoteSkip(skipCounts, "atoms " . beastId . ": " . err)
                Palace_Notify("Skip atoms for " . beastId . ": " . err, 2800, BANNER_ACCENT_ERROR)
                continue
            }
            okBeastIds[beastId] := true
        }
        if (okBeastIds.Count) {
            kept := []
            for a in atoms {
                if (!okBeastIds.Has(a["beast_id"]))
                    kept.Push(a)
            }
            atoms := kept
            for beastId, proposed in pendingByBeast {
                if (!okBeastIds.Has(beastId))
                    continue
                for row in proposed {
                    atoms.Push(row)
                    nAtoms += 1
                    sid := Palace_StudyIdForBeast(beastId)
                    if (sid != "")
                        syncStudyIds[sid] := true
                }
            }
        }
    }

    if (nPalaces)
        Palace_Save("palaces", palaces)
    if (nBeasts) {
        Palace_Save("beasts", beasts)
        Palace_Save("atoms", atoms)
    } else if (nAtoms) {
        Palace_Save("atoms", atoms)
    }

    if (nPalaces + nBeasts + nAtoms = 0) {
        summary := Palace_ImportSkipSummary(skipCounts, csvNotes)
        msg := "0 imported — Desktop files kept"
        if (summary != "")
            msg .= " — " . summary
        Palace_Notify(msg, 3500, BANNER_ACCENT_ERROR)
        Palace_ShowMainMenu()
        return false
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

    syncIds := []
    for sid, _ in syncStudyIds
        syncIds.Push(sid)
    if (syncIds.Length)
        Palace_SyncPracticeMd(syncIds)
    Palace_Notify("Imported " . nPalaces . " palace(s), " . nBeasts . " beast(s), " . nAtoms . " atom(s)",
        2800, BANNER_ACCENT_SUCCESS)
    Palace_ShowMainMenu()
    return true
}

Palace_PalacesMissingImage() {
    palaces := Palace_Load("palaces")
    missing := []
    for p in palaces {
        img := Trim(p.Has("image_rel_path") ? p["image_rel_path"] : "")
        abs := (img = "") ? "" : Palace_ResolveImagePath(img)
        if (abs = "" || !FileExist(abs))
            missing.Push(p)
    }
    ; Sort by study title then palace_number
    loop {
        swapped := false
        i := 1
        while (i < missing.Length) {
            a := missing[i]
            b := missing[i + 1]
            sa := StrLower(Palace_StudyTitle(a.Has("study_id") ? a["study_id"] : ""))
            sb := StrLower(Palace_StudyTitle(b.Has("study_id") ? b["study_id"] : ""))
            na := 0
            nb := 0
            try na := Integer(a.Has("palace_number") ? a["palace_number"] : 0)
            catch {
            }
            try nb := Integer(b.Has("palace_number") ? b["palace_number"] : 0)
            catch {
            }
            if (StrCompare(sa, sb) > 0 || (sa = sb && na > nb)) {
                missing[i] := b
                missing[i + 1] := a
                swapped := true
            }
            i += 1
        }
        if (!swapped)
            break
    }
    return missing
}

; Returns selected palace Map, or false if cancelled / empty.
Palace_PickPalaceMissingImage() {
    global g_PalaceGui
    missing := Palace_PalacesMissingImage()
    if (!missing.Length)
        return false
    owner := ""
    try {
        if (IsObject(g_PalaceGui))
            owner := " +Owner" . g_PalaceGui.Hwnd
    } catch {
        owner := ""
    }
    Palace_DialogsBegin()
    g := Gui("+AlwaysOnTop +ToolWindow" . owner, "Attach image — palaces without image")
    g.SetFont("s10", "Segoe UI")
    g.Add("Text", "w560", "Select a Memory Palace (Study · # · Name). Enter / double-click confirms.")
    lv := g.Add("ListView", "w560 h320 Grid", ["Study", "#", "Name"])
    try Palace_StyleDarkListView(lv)
    catch {
    }
    for p in missing {
        studyLab := Palace_StudyTitle(p.Has("study_id") ? p["study_id"] : "")
        num := p.Has("palace_number") ? p["palace_number"] : ""
        title := p.Has("title") ? p["title"] : ""
        lv.Add("", studyLab, num, title)
    }
    lv.ModifyCol(1, 180)
    lv.ModifyCol(2, 50)
    lv.ModifyCol(3, 300)
    if (missing.Length > 0)
        lv.Modify(1, "Select Focus Vis")
    chosen := false
    PickOk(*) {
        row := lv.GetNext()
        if (!row || row > missing.Length)
            return
        chosen := missing[row]
        g.Destroy()
    }
    g.Add("Button", "y+8 w100 Default", "OK").OnEvent("Click", PickOk)
    g.Add("Button", "x+8 w100", "Cancel").OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())
    lv.OnEvent("DoubleClick", PickOk)
    g.Show()
    try lv.Focus()
    catch {
    }
    try WinWaitClose("ahk_id " g.Hwnd)
    catch {
    }
    Palace_DialogsEnd()
    return chosen
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

; [Q] Pick a palace without image, then copy newest Desktop PNG/JPG onto it.
Palace_QuickAttachDesktopImage(*) {
    Palace_EnsureData()
    if (!Palace_PalacesMissingImage().Length) {
        Palace_Notify("All Memory Palaces already have an image", 2500, BANNER_ACCENT_ERROR)
        return false
    }
    palace := Palace_PickPalaceMissingImage()
    if (!IsObject(palace))
        return false
    src := Palace_DesktopNewestImage()
    if (src = "") {
        Palace_Notify("No PNG/JPG on Desktop", 2200, BANNER_ACCENT_ERROR)
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
    SplitPath(src, , , &ext)
    ext := StrLower(Trim(ext))
    if (ext = "")
        ext := "png"
    slugNorm := StrReplace(slug, "\", "/")
    rel := "practice/images/" . slugNorm . "/" . palace["palace_number"] . "." . ext
    practiceDestDir := Palace_PracticeDir() . "\images\" . StrReplace(slugNorm, "/", "\")
    if (!DirExist(practiceDestDir)) {
        try DirCreate(practiceDestDir)
        catch {
            Palace_Notify("Could not create practice images folder", 2500, BANNER_ACCENT_ERROR)
            return false
        }
    }
    absDest := practiceDestDir . "\" . palace["palace_number"] . "." . ext
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
    Palace_SyncPracticeMd([palace["study_id"]])
    Palace_Notify("Attached image → " . palace["title"] . " (" . rel . ")", 2800, BANNER_ACCENT_SUCCESS)
    return true
}

; --- Study plan pack import (PLAN_PACK.txt) ---

Palace_WriteTempCsvFromSection(body) {
    tmp := A_Temp . "\palace_plan_import_" . A_TickCount . ".csv"
    Palace_WriteUtf8(tmp, body)
    return tmp
}

Palace_ImportPlansFromDesktop(pathPack) {
    text := Palace_ReadUtf8(pathPack)
    if (text = "") {
        Palace_Notify("Empty PLAN_PACK", 2200, BANNER_ACCENT_ERROR)
        Palace_ShowMainMenu()
        return false
    }
    plansBody := Palace_ExtractPackFileSection(text, "PLANS.csv")
    itemsBody := Palace_ExtractPackFileSection(text, "PLAN_ITEMS.csv")
    resBody := Palace_ExtractPackFileSection(text, "PLAN_RESOURCES.csv")
    if (plansBody = "" || itemsBody = "") {
        Palace_Notify("PLAN_PACK needs PLANS.csv + PLAN_ITEMS.csv sections", 3000, BANNER_ACCENT_ERROR)
        Palace_ShowMainMenu()
        return false
    }
    tmpP := Palace_WriteTempCsvFromSection(plansBody)
    tmpI := Palace_WriteTempCsvFromSection(itemsBody)
    planRows := Palace_ReadAiImportCsv(tmpP, "study_id")
    itemRows := Palace_ReadAiImportCsv(tmpI, "plan_id")
    resRows := []
    if (resBody != "") {
        tmpR := Palace_WriteTempCsvFromSection(resBody)
        resRows := Palace_ReadAiImportCsv(tmpR, "plan_id")
        try FileDelete(tmpR)
        catch {
        }
    }
    try FileDelete(tmpP)
    catch {
    }
    try FileDelete(tmpI)
    catch {
    }

    if (!planRows.Length) {
        Palace_Notify("No plan rows in PLAN_PACK", 2200, BANNER_ACCENT_ERROR)
        Palace_ShowMainMenu()
        return false
    }

    preview := Palace_ExtractPackPreview(text)
    labels := []
    if (preview != "") {
        labels.Push("--- Human PREVIEW ---")
        for line in StrSplit(StrReplace(StrReplace(preview, "`r`n", "`n"), "`r", "`n"), "`n")
            labels.Push(line)
        labels.Push("--- Import rows ---")
    }
    labels.Push("--- Plans (" . planRows.Length . ") ---")
    for r in planRows
        labels.Push((r.Has("id") ? r["id"] : "?") . " · " . (r.Has("title") ? r["title"] : "")
        . " @ " . (r.Has("study_id") ? r["study_id"] : "?"))
    labels.Push("--- Items (" . itemRows.Length . ") ---")
    nShow := 0
    for r in itemRows {
        if (nShow >= 40) {
            labels.Push("… +" . (itemRows.Length - nShow) . " more")
            break
        }
        labels.Push((r.Has("section_path") ? r["section_path"] : "?") . " · "
        . SubStr(r.Has("text") ? r["text"] : "", 1, 60))
        nShow += 1
    }
    if (resRows.Length)
        labels.Push("--- Resources (" . resRows.Length . ") ---")
    if (!Palace_ImportConfirmPreview("Import PLAN_PACK", labels)) {
        Palace_ShowMainMenu()
        return false
    }

    studies := Palace_Load("studies")
    plans := Palace_Load("plans")
    items := Palace_Load("plan_items")
    resources := Palace_Load("plan_resources")
    syncIds := []
    nPlans := 0
    nItems := 0

    for pr in planRows {
        studyId := pr.Has("study_id") ? Trim(pr["study_id"]) : ""
        if (studyId = "" || !Palace_FindById(studies, studyId)) {
            Palace_Notify("Skip plan: unknown study_id " . studyId, 2500, BANNER_ACCENT_ERROR)
            continue
        }
        packPlanId := pr.Has("id") ? Trim(pr["id"]) : ""
        ; Replace existing active plan for this study
        oldIds := Map()
        keptPlans := []
        for p in plans {
            if (p["study_id"] = studyId) {
                oldIds[p["id"]] := true
            } else {
                keptPlans.Push(p)
            }
        }
        plans := keptPlans
        keptItems := []
        for it in items {
            if (!oldIds.Has(it["plan_id"]))
                keptItems.Push(it)
        }
        items := keptItems
        keptRes := []
        for r in resources {
            if (!oldIds.Has(r["plan_id"]))
                keptRes.Push(r)
        }
        resources := keptRes

        planId := packPlanId != "" ? packPlanId : Palace_NextId("PLAN_", plans, 4)
        if (Palace_FindById(plans, planId))
            planId := Palace_NextId("PLAN_", plans, 4)
        plans.Push(Map(
            "id", planId,
            "study_id", studyId,
            "title", pr.Has("title") ? Trim(pr["title"]) : "Study Plan",
            "sort_order", pr.Has("sort_order") ? pr["sort_order"] : "1",
            "active", pr.Has("active") ? pr["active"] : "1"
        ))
        nPlans += 1
        syncIds.Push(studyId)

        packIdToReal := Map()
        if (packPlanId != "")
            packIdToReal[packPlanId] := planId
        packIdToReal[planId] := planId

        for ir in itemRows {
            pid := ir.Has("plan_id") ? Trim(ir["plan_id"]) : ""
            if (pid = "")
                continue
            if (packIdToReal.Has(pid))
                pid := packIdToReal[pid]
            else if (pid != planId && packPlanId != "" && pid = packPlanId)
                pid := planId
            else if (pid != planId)
                continue
            iid := ir.Has("id") && Trim(ir["id"]) != "" ? Trim(ir["id"]) : Palace_NextId("PITEM_", items, 4)
            if (Palace_FindById(items, iid))
                iid := Palace_NextId("PITEM_", items, 4)
            checked := ir.Has("checked") ? Trim(ir["checked"]) : "0"
            if (checked = "true" || checked = "yes" || checked = "x" || checked = "✅")
                checked := "1"
            if (checked != "1")
                checked := "0"
            items.Push(Map(
                "id", iid,
                "plan_id", planId,
                "section_path", ir.Has("section_path") ? Trim(ir["section_path"]) : "Backlog",
                "text", ir.Has("text") ? Trim(ir["text"]) : "",
                "checked", checked,
                "sort_order", ir.Has("sort_order") ? ir["sort_order"] : String(nItems + 1)
            ))
            nItems += 1
        }
        for rr in resRows {
            pid := rr.Has("plan_id") ? Trim(rr["plan_id"]) : ""
            if (packIdToReal.Has(pid))
                pid := packIdToReal[pid]
            else if (pid != planId)
                continue
            rid := rr.Has("id") && Trim(rr["id"]) != "" ? Trim(rr["id"]) : Palace_NextId("PRES_", resources, 4)
            if (Palace_FindById(resources, rid))
                rid := Palace_NextId("PRES_", resources, 4)
            resources.Push(Map(
                "id", rid,
                "plan_id", planId,
                "section_path", rr.Has("section_path") ? Trim(rr["section_path"]) : "",
                "line", rr.Has("line") ? rr["line"] : "",
                "sort_order", rr.Has("sort_order") ? rr["sort_order"] : "1"
            ))
        }
    }

    if (!nPlans) {
        Palace_Notify("Import produced no plans", 2500, BANNER_ACCENT_ERROR)
        Palace_ShowMainMenu()
        return false
    }

    Palace_Save("plans", plans)
    Palace_Save("plan_items", items)
    Palace_Save("plan_resources", resources)
    Palace_ArchiveImported(pathPack)
    Palace_SyncPlansMd(syncIds)
    Palace_Notify("Imported " . nPlans . " plan(s), " . nItems . " item(s)", 2800, BANNER_ACCENT_SUCCESS)
    Palace_ShowMainMenu()
    return true
}
