; =============================================================================
; Utils module: prompt_context_presets.ahk
; Named context-file bundles for the prompt editor
; =============================================================================

PromptContextPresets_IniPath() {
    return A_ScriptDir "\assets\data\prompt_context_presets.ini"
}

PromptContextPresets_EnsureDefault() {
    path := PromptContextPresets_IniPath()
    if (FileExist(path))
        return
    try DirCreate(A_ScriptDir "\assets\data")
    catch {
    }
    content := ""
        . "[Preset_1]`n"
        . "Name=Finance daily CSV pack`n"
        . "PersonalContextFiles=`n"
        . "PersonalContextCompact=`n"
        . "PersonalContextCsvKeep=`n"
        . "WorkContextFiles=`n"
        . "WorkContextCompact=`n"
        . "WorkContextCsvKeep=`n"
        . "`n"
        . "[Preset_2]`n"
        . "Name=Example (empty)`n"
        . "PersonalContextFiles=`n"
        . "PersonalContextCompact=`n"
        . "PersonalContextCsvKeep=`n"
        . "WorkContextFiles=`n"
        . "WorkContextCompact=`n"
        . "WorkContextCsvKeep=`n"
    try FileAppend(content, path, "UTF-8")
    catch {
    }
}

PromptContextPresets_ScanSections(path) {
    sections := []
    if (!FileExist(path))
        return sections
    try content := FileRead(path, "UTF-8")
    catch {
        return sections
    }
    pos := 1
    loop 200 {
        if !RegExMatch(content, "\[(Preset_[^\]]+)\]", &m, pos)
            break
        sections.Push(m[1])
        pos := m.Pos(0) + m.Len(0)
    }
    return sections
}

PromptContextPresets_HasNumberedSections(path) {
    if (!FileExist(path))
        return false
    try name := IniRead(path, "Preset_1", "Name", "")
    catch {
        return false
    }
    name := PromptData_NormalizeIniValue(name)
    return (name != "" && name != "ERROR")
}

PromptContextPresets_ReadLegacyPaths(section, key) {
    raw := ""
    try raw := IniRead(PromptContextPresets_IniPath(), section, key, "")
    catch {
        raw := ""
    }
    return PromptData_ParsePathList(PromptData_NormalizeIniValue(raw))
}

PromptContextPresets_ReadSideEntries(section, prefix) {
    iniPath := PromptContextPresets_IniPath()
    entries := PromptData_ReadContextEntries(iniPath, section, prefix)
    if (entries.Length > 0)
        return entries
    legacyKey := prefix
    paths := PromptContextPresets_ReadLegacyPaths(section, legacyKey)
    out := []
    for p in paths
        out.Push(PromptData_NewContextEntry(p))
    return out
}

PromptContextPresets_MigrateIfNeeded() {
    path := PromptContextPresets_IniPath()
    PromptContextPresets_EnsureDefault()
    if (!FileExist(path))
        return
    if (PromptContextPresets_HasNumberedSections(path))
        return
    sections := PromptContextPresets_ScanSections(path)
    if (sections.Length = 0)
        return
    list := []
    for section in sections {
        name := ""
        try name := IniRead(path, section, "Name", "")
        catch {
            continue
        }
        name := PromptData_NormalizeIniValue(name)
        if (name = "" || name = "ERROR")
            continue
        list.Push({
            name: name,
            personal: PromptContextPresets_ReadSideEntries(section, "Personal"),
            work: PromptContextPresets_ReadSideEntries(section, "Work")
        })
    }
    if (list.Length > 0)
        PromptContextPresets_SaveAll(list)
}

PromptContextPresets_LoadAll() {
    PromptContextPresets_EnsureDefault()
    PromptContextPresets_MigrateIfNeeded()
    path := PromptContextPresets_IniPath()
    list := []
    loop 200 {
        section := "Preset_" . A_Index
        name := ""
        try name := IniRead(path, section, "Name", "ERROR")
        catch {
            break
        }
        if (name = "ERROR")
            break
        name := PromptData_NormalizeIniValue(name)
        list.Push({
            id: String(A_Index),
            name: name,
            personal: PromptContextPresets_ReadSideEntries(section, "Personal"),
            work: PromptContextPresets_ReadSideEntries(section, "Work")
        })
    }
    return list
}

PromptContextPresets_SaveAll(list) {
    path := PromptContextPresets_IniPath()
    try DirCreate(A_ScriptDir "\assets\data")
    catch {
    }
    try FileDelete(path)
    catch {
    }
    if (!IsObject(list))
        return true
    try {
        idx := 1
        for preset in list {
            section := "Preset_" . idx
            name := preset.HasProp("name") ? preset.name : ""
            IniWrite(name, path, section, "Name")
            personal := preset.HasProp("personal") ? preset.personal : []
            work := preset.HasProp("work") ? preset.work : []
            PromptData_WriteContextEntries(path, section, "Personal", personal)
            PromptData_WriteContextEntries(path, section, "Work", work)
            idx += 1
        }
        return true
    } catch {
        return false
    }
}

PromptContextPresets_List() {
    list := []
    for p in PromptContextPresets_LoadAll()
        list.Push({ id: p.id, section: "Preset_" . p.id, name: p.name })
    return list
}

PromptContextPresets_GetEntries(presetId, side) {
    section := "Preset_" . presetId
    prefix := (side = "work") ? "Work" : "Personal"
    return PromptContextPresets_ReadSideEntries(section, prefix)
}

PromptContextPresets_GetPreset(presetId) {
    for p in PromptContextPresets_LoadAll() {
        if (p.id = String(presetId))
            return p
    }
    return false
}

PromptContextPresets_GetPaths(presetId, side) {
    entries := PromptContextPresets_GetEntries(presetId, side)
    paths := []
    for e in entries {
        p := PromptData_ContextEntryPath(e)
        if (p != "")
            paths.Push(p)
    }
    preset := PromptContextPresets_GetPreset(presetId)
    name := IsObject(preset) ? preset.name : ""
    return { ok: paths.Length > 0, paths: paths, name: name, entries: entries }
}

PromptContextPresets_ChoiceLabels() {
    labels := ["(none)"]
    for p in PromptContextPresets_List()
        labels.Push(p.name)
    return labels
}

PromptContextPresets_IdByLabel(label) {
    for p in PromptContextPresets_List() {
        if (p.name = label)
            return p.id
    }
    return ""
}

PromptContextPresets_RefreshDropdown(ctrl) {
    if (!IsObject(ctrl))
        return
    current := ""
    try current := ctrl.Text
    catch {
    }
    labels := PromptContextPresets_ChoiceLabels()
    try ctrl.Delete()
    catch {
    }
    for label in labels {
        try ctrl.Add(label)
        catch {
        }
    }
    restored := false
    if (current != "") {
        for label in labels {
            if (label = current) {
                try {
                    ctrl.Text := label
                    restored := true
                } catch {
                }
                break
            }
        }
    }
    if (!restored) {
        try ctrl.Text := "(none)"
        catch {
        }
    }
}

PromptContextPresets_NormalizeName(name) {
    return Trim(name)
}

PromptContextPresets_NameTaken(name, exceptIndex := 0) {
    n := StrLower(PromptContextPresets_NormalizeName(name))
    if (n = "")
        return false
    idx := 0
    for p in PromptContextPresets_LoadAll() {
        idx += 1
        if (idx = exceptIndex)
            continue
        if (StrLower(p.name) = n)
            return true
    }
    return false
}

PromptContextPresets_SaveFromEditor(name, personalEntries, workEntries) {
    name := PromptContextPresets_NormalizeName(name)
    if (name = "")
        return { ok: false, msg: "Name is required." }
    if (PromptContextPresets_NameTaken(name))
        return { ok: false, msg: "A preset with that name already exists." }
    personal := PromptData_ParseContextEntries(personalEntries)
    work := PromptData_ParseContextEntries(workEntries)
    if (personal.Length = 0 && work.Length = 0)
        return { ok: false, msg: "Add at least one context file before saving a preset." }
    list := PromptContextPresets_LoadAll()
    list.Push({ name: name, personal: personal, work: work })
    if (!PromptContextPresets_SaveAll(list))
        return { ok: false, msg: "Could not write preset file." }
    return { ok: true, msg: "Preset saved.", name: name }
}
