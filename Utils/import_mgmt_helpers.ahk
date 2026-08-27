; =============================================================================
; Utils module: import_mgmt_helpers.ahk
; CSV storage, job opportunity helpers for Import Management
; =============================================================================

global g_ImportMgmtGui := false
global g_ImportMgmtHotkeys := []

ImportMgmt_DataDir() {
    dir := A_ScriptDir . "\job_search\data"
    if (!DirExist(dir))
        DirCreate(dir)
    imported := dir . "\imported"
    if (!DirExist(imported))
        DirCreate(imported)
    return dir
}

ImportMgmt_EnsureData() {
    ImportMgmt_DataDir()
    path := ImportMgmt_DataDir() . "\opportunities.csv"
    if (!FileExist(path))
        ImportMgmt_WriteCsv("opportunities.csv", [], ImportMgmt_Headers())
}

ImportMgmt_WriteUtf8(path, content) {
    f := FileOpen(path, "w", "UTF-8")
    if (!f)
        throw Error("Could not write " . path)
    f.Write(content)
    f.Close()
}

ImportMgmt_ReadUtf8(path) {
    if (!FileExist(path))
        return ""
    f := FileOpen(path, "r", "UTF-8")
    if (!f)
        return ""
    text := f.Read()
    f.Close()
    if (SubStr(text, 1, 1) = Chr(0xFEFF))
        text := SubStr(text, 2)
    return text
}

ImportMgmt_SplitCsvLine(line) {
    fields := []
    len := StrLen(line)
    i := 1
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
        } else {
            next := InStr(line, ",", false, i)
            if (!next) {
                fields.Push(SubStr(line, i))
                break
            }
            fields.Push(SubStr(line, i, next - i))
            i := next + 1
            if (i > len)
                fields.Push("")
        }
    }
    return fields
}

ImportMgmt_CsvEscape(val) {
    s := String(val)
    if (InStr(s, ",") || InStr(s, '"') || InStr(s, "`n") || InStr(s, "`r"))
        return '"' . StrReplace(s, '"', '""') . '"'
    return s
}

ImportMgmt_ReadCsv(fileName) {
    path := InStr(fileName, "\") || InStr(fileName, "/") ? fileName : ImportMgmt_DataDir() . "\" . fileName
    rows := []
    text := ImportMgmt_ReadUtf8(path)
    if (text = "")
        return rows
    text := StrReplace(text, "`r`n", "`n")
    text := StrReplace(text, "`r", "`n")
    lines := StrSplit(text, "`n")
    headers := []
    for idx, line in lines {
        if (Trim(line) = "")
            continue
        fields := ImportMgmt_SplitCsvLine(line)
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
        rows.Push(row)
    }
    return rows
}

ImportMgmt_WriteCsv(fileName, rows, headers) {
    path := InStr(fileName, "\") || InStr(fileName, "/") ? fileName : ImportMgmt_DataDir() . "\" . fileName
    out := ""
    loop headers.Length {
        if (A_Index > 1)
            out .= ","
        out .= ImportMgmt_CsvEscape(headers[A_Index])
    }
    out .= "`n"
    for row in rows {
        loop headers.Length {
            if (A_Index > 1)
                out .= ","
            key := headers[A_Index]
            val := row.Has(key) ? row[key] : ""
            out .= ImportMgmt_CsvEscape(val)
        }
        out .= "`n"
    }
    ImportMgmt_WriteUtf8(path, out)
}

ImportMgmt_Headers() {
    return ["id", "company", "role_title", "status", "status_date", "applied_date", "source", "notes"]
}

ImportMgmt_Save(rows) {
    ImportMgmt_WriteCsv("opportunities.csv", rows, ImportMgmt_Headers())
}

ImportMgmt_Load() {
    ImportMgmt_EnsureData()
    return ImportMgmt_ReadCsv("opportunities.csv")
}

ImportMgmt_FindById(rows, id) {
    id := Trim(id)
    if (id = "")
        return false
    for row in rows {
        if (row.Has("id") && row["id"] = id)
            return row
    }
    return false
}

ImportMgmt_FindByCompany(rows, company) {
    norm := ImportMgmt_NormalizeCompany(company)
    if (norm = "")
        return false
    for row in rows {
        if (ImportMgmt_NormalizeCompany(row.Has("company") ? row["company"] : "") = norm)
            return row
    }
    return false
}

ImportMgmt_IdExists(rows, id) {
    for row in rows {
        if (row.Has("id") && row["id"] = id)
            return true
    }
    return false
}

ImportMgmt_NextId(prefix, rows, pad := 3) {
    maxN := 0
    for row in rows {
        id := row.Has("id") ? row["id"] : ""
        if (SubStr(id, 1, StrLen(prefix)) != prefix)
            continue
        rest := SubStr(id, StrLen(prefix) + 1)
        if (rest != "" && IsDigit(rest)) {
            n := Integer(rest)
            if (n > maxN)
                maxN := n
        }
    }
    return prefix . Format("{:0" . pad . "d}", maxN + 1)
}

ImportMgmt_Unaccent(s) {
    pairs := [["á", "a"], ["à", "a"], ["â", "a"], ["ã", "a"], ["ä", "a"], ["é", "e"], ["ê", "e"], ["è", "e"],
    ["í", "i"], ["ó", "o"], ["ô", "o"], ["õ", "o"], ["ö", "o"], ["ú", "u"], ["ü", "u"], ["ç", "c"],
    ["Á", "A"], ["À", "A"], ["Â", "A"], ["Ã", "A"], ["É", "E"], ["Ê", "E"], ["Í", "I"], ["Ó", "O"],
    ["Ô", "O"], ["Õ", "O"], ["Ú", "U"], ["Ç", "C"]]
    for p in pairs
        s := StrReplace(s, p[1], p[2])
    return s
}

ImportMgmt_NormalizeCompany(name) {
    s := ImportMgmt_Unaccent(Trim(name))
    s := StrLower(s)
    s := RegExReplace(s, "[^a-z0-9]+", "")
    return s
}

ImportMgmt_Slug(name) {
    s := ImportMgmt_Unaccent(Trim(name))
    s := StrUpper(s)
    out := ""
    loop parse s {
        c := Ord(A_LoopField)
        if ((c >= 65 && c <= 90) || (c >= 48 && c <= 57))
            out .= A_LoopField
    }
    if (StrLen(out) > 24)
        out := SubStr(out, 1, 24)
    if (out = "")
        out := "X"
    return out
}

ImportMgmt_SlugId(prefix, name, existing) {
    base := ImportMgmt_Slug(name)
    id := prefix . "_" . base
    if (!ImportMgmt_IdExists(existing, id))
        return id
    n := 2
    loop {
        cand := id . n
        if (!ImportMgmt_IdExists(existing, cand))
            return cand
        n += 1
    }
}

ImportMgmt_ValidStatuses() {
    return ["applied", "screening", "interviewing", "offer", "rejected", "withdrawn", "on_hold"]
}

ImportMgmt_ValidStatus(s) {
    s := StrLower(Trim(s))
    if (s = "")
        return ""
    for st in ImportMgmt_ValidStatuses() {
        if (s = st)
            return st
    }
    return ""
}

ImportMgmt_Today() {
    return FormatTime(, "yyyy-MM-dd")
}

ImportMgmt_CloseGui() {
    global g_ImportMgmtGui
    ImportMgmt_UnbindHotkeys()
    try {
        if (IsObject(g_ImportMgmtGui))
            g_ImportMgmtGui.Destroy()
    } catch {
    }
    g_ImportMgmtGui := false
}

ImportMgmt_UnbindHotkeys() {
    global g_ImportMgmtHotkeys
    try HotIf(ImportMgmt_HotIfKeys)
    catch {
    }
    for key in g_ImportMgmtHotkeys {
        try Hotkey(key, "Off")
        catch {
        }
    }
    g_ImportMgmtHotkeys := []
    try HotIf()
    catch {
    }
}

ImportMgmt_HotIfKeys(*) {
    global g_ImportMgmtGui
    if (!IsObject(g_ImportMgmtGui))
        return false
    try {
        return WinActive("ahk_id " g_ImportMgmtGui.Hwnd)
    } catch {
        return false
    }
}

ImportMgmt_BindHotkeys(pairs) {
    global g_ImportMgmtGui, g_ImportMgmtHotkeys
    ImportMgmt_UnbindHotkeys()
    if (!IsObject(g_ImportMgmtGui))
        return
    try HotIf(ImportMgmt_HotIfKeys)
    catch {
        return
    }
    for p in pairs {
        try {
            Hotkey(p[1], p[2], "On")
            g_ImportMgmtHotkeys.Push(p[1])
        } catch {
        }
    }
    try HotIf()
    catch {
    }
}

ImportMgmt_CenterGui(guiObj, w := 560, h := 220) {
    MonitorGetWorkArea(MonitorGetPrimary(), &L, &T, &R, &B)
    x := L + ((R - L) - w) // 2
    y := T + ((B - T) - h) // 2
    guiObj.Show("x" . x . " y" . y . " w" . w . " h" . h)
}

ImportMgmt_Notify(msg, ms := 1800, accent := "") {
    if (accent = "")
        accent := BANNER_ACCENT_INFO
    try ShowCenteredOverlay_Utils(msg, ms, accent)
    catch {
        TrayTip("Import Management", msg)
    }
}
