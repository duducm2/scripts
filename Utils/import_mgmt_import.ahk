; =============================================================================
; Utils module: import_mgmt_import.ahk
; Import AI-generated JOB_SEARCH_UPDATE packs from Desktop into opportunities.csv
; =============================================================================

ImportMgmt_DesktopNewest(pattern) {
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

ImportMgmt_MdFence() {
    bt := Chr(96)
    return bt . bt . bt
}

ImportMgmt_ExtractPackFileSection(text, fileName) {
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

ImportMgmt_StripMarkdownFences(body) {
    fence := ImportMgmt_MdFence()
    body := Trim(body, "`r`n `t")
    if (SubStr(body, 1, 3) = fence) {
        nl := InStr(body, "`n")
        body := nl ? SubStr(body, nl + 1) : ""
    }
    body := Trim(body, "`r`n `t")
    if (SubStr(body, -3) = fence)
        body := Trim(SubStr(body, 1, StrLen(body) - 3), "`r`n `t")
    return body
}

ImportMgmt_StripPackPreview(text) {
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
        text := SubStr(text, 1, pos - 1) . SubStr(rest, endPos + StrLen(endNeedle))
        break
    }
    return text
}

ImportMgmt_ExtractCsvFromHeader(text) {
    fence := ImportMgmt_MdFence()
    text := StrReplace(text, "`r`n", "`n")
    text := StrReplace(text, "`r", "`n")
    lines := StrSplit(text, "`n")
    start := 0
    for idx, line in lines {
        t := Trim(line)
        if (t = "")
            continue
        lower := StrLower(t)
        if (InStr(lower, fence) = 1)
            continue
        if (InStr(lower, "file:") = 1 || SubStr(t, 1, 1) = "#")
            continue
        if (InStr(lower, "company") && InStr(lower, "status")) {
            start := idx
            break
        }
    }
    if (!start)
        return ""
    cleaned := ""
    loop lines.Length {
        if (A_Index < start)
            continue
        line := lines[A_Index]
        if (Trim(line) = fence)
            break
        if (cleaned != "")
            cleaned .= "`n"
        cleaned .= line
    }
    return Trim(cleaned, "`r`n `t")
}

ImportMgmt_MaterializeAiCsv(path, expectedFileName := "JOB_SEARCH_UPDATE.csv") {
    text := ImportMgmt_ReadUtf8(path)
    if (text = "")
        return path
    stem := expectedFileName
    stem := RegExReplace(stem, "i)\.(csv|txt|ini)$", "")
    body := ImportMgmt_ExtractPackFileSection(text, stem . ".csv")
    if (body = "")
        body := ImportMgmt_ExtractPackFileSection(text, stem . ".txt")
    if (body = "")
        body := ImportMgmt_ExtractCsvFromHeader(ImportMgmt_StripPackPreview(text))
    if (body = "")
        return path
    body := ImportMgmt_StripMarkdownFences(body)
    if (Trim(body) = "")
        return path
    tmp := A_Temp . "\job_search_pack_" . StrReplace(expectedFileName, ".", "_") . ".csv"
    ImportMgmt_WriteUtf8(tmp, body)
    return tmp
}

ImportMgmt_ReadAiImportCsv(path) {
    text := ImportMgmt_ReadUtf8(path)
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
        if (InStr(lower, "company") && InStr(lower, "status")) {
            start := idx
            break
        }
    }
    if (!start)
        return ImportMgmt_ReadCsv(path)
    cleaned := ""
    loop lines.Length {
        if (A_Index < start)
            continue
        if (cleaned != "")
            cleaned .= "`n"
        cleaned .= lines[A_Index]
    }
    tmp := A_Temp . "\job_search_ai_import_norm.csv"
    ImportMgmt_WriteUtf8(tmp, cleaned)
    rows := ImportMgmt_ReadCsv(tmp)
    try FileDelete(tmp)
    catch {
    }
    cleaned_rows := []
    for r in rows {
        c := r.Has("company") ? StrLower(Trim(r["company"])) : ""
        if (c = "company")
            continue
        cleaned_rows.Push(r)
    }
    return cleaned_rows
}

ImportMgmt_DesktopNewestCodeDump() {
    newest := ""
    newestTime := 0
    loop files A_Desktop . "\gemini-code*.txt", "F" {
        body := ImportMgmt_ReadUtf8(A_LoopFileFullPath)
        if (body = "")
            continue
        lower := StrLower(body)
        hasJob := InStr(lower, "company,role_title,status") || InStr(lower, "id,company,role_title")
        || InStr(lower, "file: job_search_update")
        if (!hasJob)
            continue
        ts := Number(A_LoopFileTimeModified)
        if (ts > newestTime) {
            newestTime := ts
            newest := A_LoopFileFullPath
        }
    }
    return newest
}

ImportMgmt_ArchiveImported(path) {
    destDir := ImportMgmt_DataDir() . "\imported"
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

ImportMgmt_WriteAiCompanionImportError(errorMsg, extraNotes := "") {
    errorMsg := Trim(errorMsg)
    if (errorMsg = "")
        return ""
    body := "The Desktop Job Search importer rejected my last JOB_SEARCH_UPDATE.txt. Fix and re-deliver.`r`n`r`n"
        . "IMPORT ERROR`r`n"
        . errorMsg . "`r`n`r`n"
    if (Trim(extraNotes) != "")
        body .= "EXTRA NOTES`r`n" . Trim(extraNotes) . "`r`n`r`n"
    body .= "WHAT YOU MUST DO`r`n"
        . "- Re-deliver one complete JOB_SEARCH_UPDATE.txt (download chip preferred; else one marked code fence).`r`n"
        . "- Never claim you saved to Desktop / disk. I save the file myself.`r`n"
        . "- Pack must include ===PREVIEW=== … ===END_PREVIEW=== and:`r`n"
        . "  ===FILE: JOB_SEARCH_UPDATE.csv=== … ===END_FILE===`r`n"
        . "- FILE body is pure CSV with header: id,company,role_title,status,status_date,applied_date,source,notes`r`n"
        . "- status must be one of: applied | screening | interviewing | offer | rejected | withdrawn | on_hold`r`n"
        . "- Match existing rows by id or company from attached opportunities.csv.`r`n`r`n"
        . "After you fix it, I will save JOB_SEARCH_UPDATE.txt to Desktop and re-import.`r`n"
    path := A_Desktop . "\JOB_SEARCH_AI_FIX.txt"
    try {
        ImportMgmt_WriteUtf8(path, body)
        return path
    } catch {
        return ""
    }
}

ImportMgmt_FailAiImport(errorMsg, notifyMs := 2200, extraNotes := "") {
    path := ImportMgmt_WriteAiCompanionImportError(errorMsg, extraNotes)
    notify := errorMsg
    if (path != "")
        notify .= " — AI fix → Desktop JOB_SEARCH_AI_FIX.txt"
    ImportMgmt_Notify(notify, notifyMs, BANNER_ACCENT_ERROR)
    return path != ""
}

ImportMgmt_MergeImportRow(existing, incoming) {
    headers := ImportMgmt_Headers()
    oldStatus := existing.Has("status") ? Trim(existing["status"]) : ""
    for key in headers {
        if (key = "id")
            continue
        if (!incoming.Has(key))
            continue
        val := Trim(incoming[key])
        if (val = "")
            continue
        if (key = "status") {
            valid := ImportMgmt_ValidStatus(val)
            if (valid = "")
                throw Error("Invalid status: " . val)
            existing[key] := valid
        } else {
            existing[key] := val
        }
    }
    newStatus := existing.Has("status") ? Trim(existing["status"]) : ""
    if (newStatus != "" && newStatus != oldStatus) {
        if (!existing.Has("status_date") || Trim(existing["status_date"]) = "")
            existing["status_date"] := ImportMgmt_Today()
    }
    return existing
}

ImportMgmt_NewRowFromImport(incoming, rows) {
    row := Map()
    for key in ImportMgmt_Headers()
        row[key] := incoming.Has(key) ? Trim(incoming[key]) : ""
    company := row["company"]
    if (company = "")
        throw Error("New opportunity requires company")
    if (row["id"] = "")
        row["id"] := ImportMgmt_SlugId("JOB", company, rows)
    status := ImportMgmt_ValidStatus(row["status"])
    if (status = "")
        status := "applied"
    row["status"] := status
    if (Trim(row["status_date"]) = "")
        row["status_date"] := ImportMgmt_Today()
    return row
}

ImportMgmt_ImportFromDesktop(*) {
    path := ImportMgmt_DesktopNewest("JOB_SEARCH_UPDATE*.txt")
    if (path = "")
        path := ImportMgmt_DesktopNewest("JOB_SEARCH_UPDATE*.csv")
    if (path = "")
        path := ImportMgmt_DesktopNewestCodeDump()
    if (path = "" || !FileExist(path)) {
        ImportMgmt_FailAiImport("No JOB_SEARCH_UPDATE file on Desktop", 2000)
        return false
    }
    sourcePath := path
    csvPath := ImportMgmt_MaterializeAiCsv(path, "JOB_SEARCH_UPDATE.csv")
    rows := ImportMgmt_ReadAiImportCsv(csvPath)
    if (csvPath != sourcePath) {
        try FileDelete(csvPath)
        catch {
        }
    }
    if (!rows.Length) {
        ImportMgmt_FailAiImport("File has no data rows", 2200)
        return false
    }
    opportunities := ImportMgmt_Load()
    nUpdated := 0
    nAdded := 0
    errors := []
    for incoming in rows {
        try {
            id := incoming.Has("id") ? Trim(incoming["id"]) : ""
            company := incoming.Has("company") ? Trim(incoming["company"]) : ""
            existing := false
            if (id != "")
                existing := ImportMgmt_FindById(opportunities, id)
            if (!IsObject(existing) && company != "")
                existing := ImportMgmt_FindByCompany(opportunities, company)
            if (IsObject(existing)) {
                ImportMgmt_MergeImportRow(existing, incoming)
                nUpdated += 1
            } else {
                newRow := ImportMgmt_NewRowFromImport(incoming, opportunities)
                opportunities.Push(newRow)
                nAdded += 1
            }
        } catch as e {
            label := company != "" ? company : id
            if (label = "")
                label := "row"
            errors.Push(label . ": " . e.Message)
        }
    }
    if (!nUpdated && !nAdded) {
        extra := errors.Length ? "`r`n" . errors[1] : ""
        ImportMgmt_FailAiImport("No rows applied" . extra, 2200)
        return false
    }
    ImportMgmt_Save(opportunities)
    ImportMgmt_ArchiveImported(sourcePath)
    msg := "Imported "
    parts := []
    if (nUpdated)
        parts.Push(nUpdated . " update(s)")
    if (nAdded)
        parts.Push(nAdded . " new")
    msg .= parts.Length ? parts[1] . (parts.Length > 1 ? ", " . parts[2] : "") : "0 rows"
    if (errors.Length)
        msg .= " (" . errors.Length . " skipped)"
    ImportMgmt_Notify(msg, 1800, BANNER_ACCENT_SUCCESS)
    return true
}
