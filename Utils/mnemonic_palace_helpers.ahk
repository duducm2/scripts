; =============================================================================
; Utils module: mnemonic_palace_helpers.ahk
; Memory Palace — CSV I/O, paths, GUI shared helpers, vocabulary
; =============================================================================

global g_PalaceGui := false
global g_PalaceHotkeys := []
global g_PalaceFilterStudyId := ""
global g_PalaceFilterPalaceId := ""
global g_PalaceFilterBeastId := ""

Palace_Terms() {
    return [
        ["Study", "A broad subject domain (e.g. English, German, science, piano). Contains one or more Memory Palaces."],
        ["Memory Palace", "A location with exactly one generated image. Numbered within its Study."],
        ["Character", "Sourced from the canon characters.json. Exactly one character anchors each Memory Palace."],
        ["Beast", "Sourced from the canon bestiary.json. Peg animal/creature that carries a Knowledge Atom."],
        ["Knowledge Atom",
            "A discrete piece of information on a Beast, made of Concept, Quote, Story, and Sensory."],
        ["Concept", "Rehearsal definition of the fact — what you recall to know what the atom means."],
        ["Quote", "Verbatim source payload (e.g. from a transcript). Not the Concept."],
        ["Story", "Bizarre mnemonic narrative / action that encodes the Concept."],
        ["Sensory",
            "Which sensory modality the Story emphasizes (visual, auditory, tactile, olfactory, gustatory, thermal)."],
        ["Mapping",
            "A Knowledge Atom attaches to a Beast. A Beast carries one Knowledge Atom, or up to four zoned Knowledge Atoms (Z1–Z4)."]
    ]
}

Palace_DataDir() {
    dir := A_ScriptDir . "\mnemonics\data"
    if (!DirExist(dir))
        DirCreate(dir)
    imported := dir . "\imported"
    if (!DirExist(imported))
        DirCreate(imported)
    return dir
}

Palace_OutputDir() {
    dir := A_ScriptDir . "\mnemonics\output"
    if (!DirExist(dir))
        DirCreate(dir)
    return dir
}

Palace_PythonDir() {
    return A_ScriptDir . "\mnemonics\python"
}

Palace_FindPythonCmd() {
    candidates := ["py -3", "py", "python3", "python"]
    for c in candidates {
        try {
            ec := RunWait(A_ComSpec . ' /c ' . c . ' -c "print(1)" >nul 2>&1', , "Hide")
            if (ec = 0)
                return c
        } catch {
        }
    }
    localApps := EnvGet("LOCALAPPDATA")
    pathGlobs := [
        localApps . "\Programs\Python\Python3*\python.exe",
        EnvGet("ProgramFiles") . "\Python3*\python.exe",
        "C:\Python3*\python.exe"
    ]
    for g in pathGlobs {
        loop files g, "F" {
            try {
                ec := RunWait('"' . A_LoopFileFullPath . '" -c "print(1)"', , "Hide")
                if (ec = 0)
                    return '"' . A_LoopFileFullPath . '"'
            } catch {
            }
        }
    }
    return ""
}

Palace_SettingsPath() {
    return Palace_DataDir() . "\settings.ini"
}

Palace_EnsureSettings() {
    path := Palace_SettingsPath()
    if (FileExist(path))
        return
    content := "[Dashboard]`n"
        . "ShowPalaceGrid=1`n"
        . "ShowMissingImages=1`n"
        . "ShowBeastCounts=1`n"
        . "`n[General]`n"
        . "LastStudyId=`n"
        . "NotesStudiesRoot=`n"
    Palace_WriteUtf8(path, content)
}

Palace_EnsureData() {
    Palace_DataDir()
    Palace_EnsureSettings()
    for kind in ["studies", "palaces", "beasts", "atoms"] {
        path := Palace_DataDir() . "\" . kind . ".csv"
        if (!FileExist(path))
            Palace_Save(kind, [])
    }
}

Palace_Setting(section, key, default := "") {
    val := IniRead(Palace_SettingsPath(), section, key, default)
    if (val = "ERROR")
        return default
    return val
}

Palace_SetSetting(section, key, value) {
    IniWrite(value, Palace_SettingsPath(), section, key)
}

Palace_NotesStudiesRoot() {
    override := Trim(Palace_Setting("General", "NotesStudiesRoot", ""))
    if (override != "" && DirExist(override))
        return RTrim(override, "\")
    try {
        root := GetNotesRepoPath()
        if (root != "") {
            studies := RTrim(root, "\") . "\studies"
            if (DirExist(studies))
                return studies
        }
    } catch {
    }
    return ""
}

Palace_ResolveImagePath(imageRelPath) {
    if (Trim(imageRelPath) = "")
        return ""
    if (InStr(imageRelPath, ":") || SubStr(imageRelPath, 1, 2) = "\\")
        return imageRelPath
    root := Palace_NotesStudiesRoot()
    if (root = "")
        return ""
    return root . "\" . StrReplace(imageRelPath, "/", "\")
}

Palace_WriteUtf8(path, content) {
    f := FileOpen(path, "w", "UTF-8")
    if (!f)
        throw Error("Could not write " . path)
    f.Write(content)
    f.Close()
}

Palace_ReadUtf8(path) {
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

Palace_SplitCsvLine(line) {
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

Palace_CsvEscape(val) {
    s := String(val)
    if (InStr(s, ",") || InStr(s, '"') || InStr(s, "`n") || InStr(s, "`r"))
        return '"' . StrReplace(s, '"', '""') . '"'
    return s
}

Palace_ReadCsv(fileName) {
    path := InStr(fileName, "\") || InStr(fileName, "/") ? fileName : Palace_DataDir() . "\" . fileName
    rows := []
    text := Palace_ReadUtf8(path)
    if (text = "")
        return rows
    text := StrReplace(text, "`r`n", "`n")
    text := StrReplace(text, "`r", "`n")
    lines := StrSplit(text, "`n")
    headers := []
    for idx, line in lines {
        if (Trim(line) = "")
            continue
        fields := Palace_SplitCsvLine(line)
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

Palace_WriteCsv(fileName, rows, headers) {
    path := InStr(fileName, "\") || InStr(fileName, "/") ? fileName : Palace_DataDir() . "\" . fileName
    out := ""
    loop headers.Length {
        if (A_Index > 1)
            out .= ","
        out .= Palace_CsvEscape(headers[A_Index])
    }
    out .= "`n"
    for row in rows {
        loop headers.Length {
            if (A_Index > 1)
                out .= ","
            key := headers[A_Index]
            val := row.Has(key) ? row[key] : ""
            out .= Palace_CsvEscape(val)
        }
        out .= "`n"
    }
    Palace_WriteUtf8(path, out)
}

Palace_Headers(kind) {
    switch kind {
        case "studies":
            return ["id", "title", "notes_rel_path", "sort_order", "active"]
        case "palaces":
            return ["id", "study_id", "palace_number", "title", "character_name", "image_rel_path", "depth_slots_used",
                "image_prompt"]
        case "beasts":
            return ["id", "palace_id", "peg_code", "beast_name", "beast_source", "sensory_channel", "is_smashed",
                "sort_order"]
        case "atoms":
            return ["id", "beast_id", "kind", "zone", "zone_label", "concept", "quote", "story", "sensory", "ipa",
                "sort_order"]
        default:
            return []
    }
}

Palace_Save(kind, rows) {
    Palace_WriteCsv(kind . ".csv", rows, Palace_Headers(kind))
}

Palace_Load(kind) {
    return Palace_ReadCsv(kind . ".csv")
}

Palace_FindById(rows, id) {
    for row in rows {
        if (row.Has("id") && row["id"] = id)
            return row
    }
    return false
}

Palace_IdExists(rows, id) {
    for row in rows {
        if (row.Has("id") && row["id"] = id)
            return true
    }
    return false
}

Palace_Unaccent(s) {
    pairs := [["á", "a"], ["à", "a"], ["â", "a"], ["ã", "a"], ["ä", "a"], ["é", "e"], ["ê", "e"], ["è", "e"],
    ["í", "i"], ["ó", "o"], ["ô", "o"], ["õ", "o"], ["ö", "o"], ["ú", "u"], ["ü", "u"], ["ç", "c"],
    ["Á", "A"], ["À", "A"], ["Â", "A"], ["Ã", "A"], ["É", "E"], ["Ê", "E"], ["Í", "I"], ["Ó", "O"],
    ["Ô", "O"], ["Õ", "O"], ["Ú", "U"], ["Ç", "C"]]
    for p in pairs
        s := StrReplace(s, p[1], p[2])
    return s
}

Palace_Slug(name) {
    s := Palace_Unaccent(Trim(name))
    s := StrUpper(s)
    out := ""
    loop parse s {
        c := Ord(A_LoopField)
        if ((c >= 65 && c <= 90) || (c >= 48 && c <= 57))
            out .= A_LoopField
    }
    if (StrLen(out) > 12)
        out := SubStr(out, 1, 12)
    if (out = "")
        out := "X"
    return out
}

Palace_SlugId(prefix, name, existing) {
    base := Palace_Slug(name)
    id := prefix . base
    if (!Palace_IdExists(existing, id))
        return id
    n := 2
    loop {
        cand := id . n
        if (!Palace_IdExists(existing, cand))
            return cand
        n += 1
    }
}

Palace_NextId(prefix, rows, pad := 3) {
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

Palace_FilterBy(rows, key, value) {
    out := []
    for r in rows {
        if (r.Has(key) && r[key] = value)
            out.Push(r)
    }
    return out
}

Palace_StudyTitle(studyId) {
    s := Palace_FindById(Palace_Load("studies"), studyId)
    return s ? s["title"] : studyId
}

Palace_PalaceLabel(palaceId) {
    st := Palace_FindById(Palace_Load("palaces"), palaceId)
    if (!st)
        return palaceId
    return "Memory Palace " . st["palace_number"] . ": " . st["title"]
}

Palace_BeastLabel(beastId) {
    b := Palace_FindById(Palace_Load("beasts"), beastId)
    if (!b)
        return beastId
    return "[" . b["peg_code"] . "] " . b["beast_name"]
}

Palace_ValidateBeastAtoms(beastId, proposedRows := false) {
    atoms := IsObject(proposedRows) ? proposedRows : Palace_FilterBy(Palace_Load("atoms"), "beast_id", beastId)
    singles := 0
    zoned := 0
    for a in atoms {
        kind := a.Has("kind") ? StrLower(a["kind"]) : "single"
        if (kind = "zoned" || kind = "subtopic")
            zoned += 1
        else
            singles += 1
    }
    if (singles > 0 && zoned > 0)
        return "Beast cannot mix a single Knowledge Atom with zoned Knowledge Atoms."
    if (singles > 1)
        return "Beast may carry only one comprehensive Knowledge Atom."
    if (zoned > 4)
        return "Beast may carry at most four zoned Knowledge Atoms (Z1–Z4)."
    return ""
}

Palace_CloseGui() {
    global g_PalaceGui, g_PalaceHotkeys
    Palace_UnbindHotkeys()
    if (IsObject(g_PalaceGui)) {
        try g_PalaceGui.Destroy()
        catch {
        }
    }
    g_PalaceGui := false
}

Palace_UnbindHotkeys() {
    global g_PalaceHotkeys
    try HotIf(Palace_HotIfPalaceKeys)
    catch {
    }
    for item in g_PalaceHotkeys {
        try Hotkey(item, "Off")
        catch {
        }
    }
    g_PalaceHotkeys := []
    try HotIf()
    catch {
    }
}

Palace_GuiFocusIsEdit() {
    global g_PalaceGui
    if (!IsObject(g_PalaceGui))
        return false
    try {
        focused := ControlGetFocus("ahk_id " g_PalaceGui.Hwnd)
        return (focused != "" && InStr(focused, "Edit") = 1)
    } catch {
        return false
    }
}

Palace_HotIfPalaceKeys(*) {
    global g_PalaceGui
    if (!IsObject(g_PalaceGui))
        return false
    try {
        if (!WinActive("ahk_id " g_PalaceGui.Hwnd))
            return false
        return !Palace_GuiFocusIsEdit()
    } catch {
        return false
    }
}

Palace_BindHotkeys(pairs) {
    global g_PalaceGui, g_PalaceHotkeys
    Palace_UnbindHotkeys()
    if (!IsObject(g_PalaceGui))
        return
    try HotIf(Palace_HotIfPalaceKeys)
    catch {
        return
    }
    for p in pairs {
        try {
            Hotkey(p[1], p[2], "On")
            g_PalaceHotkeys.Push(p[1])
        } catch {
        }
    }
    try HotIf()
    catch {
    }
}

Palace_CenterGui(guiObj, w := 920, h := 620) {
    MonitorGetWorkArea(MonitorGetPrimary(), &L, &T, &R, &B)
    x := L + ((R - L) - w) // 2
    y := T + ((B - T) - h) // 2
    guiObj.Show("x" . x . " y" . y . " w" . w . " h" . h)
}

Palace_Notify(msg, ms := 1800, accent := "") {
    if (accent = "")
        accent := BANNER_ACCENT_INFO
    try ShowCenteredOverlay_Utils("Memory Palace: " . msg, ms, accent)
    catch {
        TrayTip("Memory Palace", msg)
    }
}

Palace_DialogsBegin() {
    global g_PalaceGui
    try {
        if (IsObject(g_PalaceGui))
            g_PalaceGui.Opt("-AlwaysOnTop")
    } catch {
    }
}

Palace_DialogsEnd() {
    global g_PalaceGui
    try {
        if (IsObject(g_PalaceGui))
            g_PalaceGui.Opt("+AlwaysOnTop")
    } catch {
    }
}

Palace_OwnerOpt() {
    global g_PalaceGui
    hwnd := 0
    try {
        if (IsObject(g_PalaceGui))
            hwnd := g_PalaceGui.Hwnd
    } catch {
        hwnd := 0
    }
    return hwnd ? " Owner" . hwnd : ""
}

Palace_Confirm(msg, title := "Memory Palace") {
    Palace_DialogsBegin()
    result := MsgBox(msg, title, "YesNo Icon?" . Palace_OwnerOpt())
    Palace_DialogsEnd()
    return result = "Yes"
}

Palace_Alert(msg, title := "Memory Palace") {
    Palace_DialogsBegin()
    MsgBox(msg, title, "Icon!" . Palace_OwnerOpt())
    Palace_DialogsEnd()
}

Palace_PickList(title, labels, values) {
    if (!labels.Length)
        return ""
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
    lv := g.Add("ListView", "w420 h280 Grid", ["Choice"])
    loop labels.Length
        lv.Add("", labels[A_Index])
    lv.ModifyCol(1, "AutoHdr")
    chosen := ""
    g.Add("Button", "y+8 w100 Default", "OK").OnEvent("Click", PickOk)
    g.Add("Button", "x+8 w100", "Cancel").OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())
    lv.OnEvent("DoubleClick", PickOk)
    g.Show()
    try WinWaitClose("ahk_id " g.Hwnd)
    catch {
    }
    Palace_DialogsEnd()
    return chosen

    PickOk(*) {
        row := lv.GetNext()
        if (!row || row > values.Length)
            return
        chosen := values[row]
        g.Destroy()
    }
}

Palace_PickStudy() {
    studies := Palace_Load("studies")
    labels := []
    values := []
    for s in studies {
        if (s.Has("active") && s["active"] = "0")
            continue
        labels.Push(s["title"] . " (" . s["notes_rel_path"] . ")")
        values.Push(s["id"])
    }
    return Palace_PickList("Pick study", labels, values)
}

Palace_PickPalace(studyId := "") {
    palaces := Palace_Load("palaces")
    labels := []
    values := []
    for st in palaces {
        if (studyId != "" && st["study_id"] != studyId)
            continue
        labels.Push(Palace_StudyTitle(st["study_id"]) . " · Memory Palace " . st["palace_number"] . ": " . st["title"])
        values.Push(st["id"])
    }
    return Palace_PickList("Pick Memory Palace", labels, values)
}

Palace_PickBeast(palaceId := "") {
    beasts := Palace_Load("beasts")
    labels := []
    values := []
    for b in beasts {
        if (palaceId != "" && b["palace_id"] != palaceId)
            continue
        labels.Push("[" . b["peg_code"] . "] " . b["beast_name"])
        values.Push(b["id"])
    }
    return Palace_PickList("Pick beast", labels, values)
}
