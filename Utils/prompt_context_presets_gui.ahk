; =============================================================================
; Utils module: prompt_context_presets_gui.ahk
; Preset Manager GUI, Save-as-preset dialog, contextual help
; =============================================================================

global g_PresetMgrGui := false
global g_PresetMgrHelpGui := false
global g_PresetMgrOwnerHwnd := 0
global g_PresetMgrOnClose := ""
global g_PresetMgrPresets := []
global g_PresetMgrEditIndex := 0
global g_PresetMgrNameCtrl := false
global g_PresetMgrListLv := false
global g_PresetMgrPersonalPaths := []
global g_PresetMgrWorkPaths := []
global g_PresetMgrPersonalLv := false
global g_PresetMgrWorkLv := false
global g_PresetMgrFlagCtrls := Map()
global g_PresetMgrSuppressFlagSync := false

PromptContextPresets_HelpText() {
    return "
(
Context presets

What is a preset?
A preset is a reusable bundle of context attachment files (personal and/or work lists). Use presets when several prompts share the same file pack — for example finance CSVs or a brand guide JSON.

Personal vs work
Each preset can store two lists. At paste time, only the list for your current environment (personal or work) is attached.

Apply (in the prompt editor)
Pick a preset and click Apply on the personal or work side. Files are added to the prompt you are editing; if a path already exists, Compact and CSV keep flags are updated from the preset. Click Save on the prompt editor to keep changes on that prompt.

Save as preset
Captures the current personal and work lists from the open prompt editor, including Compact and CSV keep flags. You need at least one file in either list.

Preset Manager
Create, rename, edit file lists, or delete presets. Changes are written to assets\data\prompt_context_presets.ini.

When to use presets vs per-prompt lists
Use per-prompt lists for one-off attachments. Use presets for packs you reuse across many prompts.

Compact and CSV keep
Presets store the same flags as prompts. Set them on each row in the manager (or in the editor before Save as preset). F1 in the prompt editor still documents compact/CSV rules.
)"
}

PromptContextPresets_ShowHelp(*) {
    global g_PresetMgrHelpGui, g_PresetMgrGui, g_PromptEditorGui
    if (IsObject(g_PresetMgrHelpGui)) {
        try WinActivate("ahk_id " g_PresetMgrHelpGui.Hwnd)
        catch {
        }
        return
    }
    ownerOpt := "+AlwaysOnTop +ToolWindow"
    ownerHwnd := 0
    if (IsObject(g_PresetMgrGui))
        ownerHwnd := g_PresetMgrGui.Hwnd
    else if (IsObject(g_PromptEditorGui))
        ownerHwnd := g_PromptEditorGui.Hwnd
    if (ownerHwnd)
        ownerOpt .= " +Owner" . ownerHwnd
    g_PresetMgrHelpGui := Gui(ownerOpt, "Context presets help")
    g_PresetMgrHelpGui.SetFont("s10", "Segoe UI")
    g_PresetMgrHelpGui.Add("Edit", "xm w520 r20 ReadOnly -WantReturn Wrap", PromptContextPresets_HelpText())
    g_PresetMgrHelpGui.Add("Button", "xm+220 w80 Default", "Close").OnEvent("Click", PromptContextPresets_CloseHelp)
    g_PresetMgrHelpGui.OnEvent("Close", PromptContextPresets_CloseHelp)
    g_PresetMgrHelpGui.OnEvent("Escape", PromptContextPresets_CloseHelp)
    mon := UtilitySelector_ActiveMonitorWorkArea()
    g_PresetMgrHelpGui.Show("Hide")
    g_PresetMgrHelpGui.GetPos(, , &gw, &gh)
    g_PresetMgrHelpGui.Show("x" (mon.left + (mon.width - gw) // 2) " y" (mon.top + (mon.height - gh) // 2))
}

PromptContextPresets_CloseHelp(*) {
    global g_PresetMgrHelpGui
    if (IsObject(g_PresetMgrHelpGui)) {
        try g_PresetMgrHelpGui.Destroy()
        catch {
        }
    }
    g_PresetMgrHelpGui := false
}

PromptContextPresets_ShowSaveDialog(personalEntries, workEntries, ownerHwnd := 0) {
    ib := InputBox("Preset name:", "Save as preset", "w420", "")
    if (ib.Result != "OK")
        return { ok: false, cancelled: true }
    r := PromptContextPresets_SaveFromEditor(ib.Value, personalEntries, workEntries)
    if (r.ok) {
        UtilitySelector_Notify("Preset saved: " . r.name)
        PromptEditor_RefreshPresetDropdowns()
    } else {
        UtilitySelector_Notify(r.msg)
    }
    return r
}

PromptContextPresets_ShowManager(ownerHwnd := 0, onClose := "") {
    global g_PresetMgrGui, g_PresetMgrOwnerHwnd, g_PresetMgrOnClose
    if (IsObject(g_PresetMgrGui)) {
        try WinActivate("ahk_id " g_PresetMgrGui.Hwnd)
        catch {
        }
        return
    }
    g_PresetMgrOwnerHwnd := ownerHwnd
    g_PresetMgrOnClose := onClose
    g_PresetMgrPresets := PromptContextPresets_LoadAll()
    g_PresetMgrEditIndex := 0
    g_PresetMgrPersonalPaths := []
    g_PresetMgrWorkPaths := []
    ownerOpt := "+AlwaysOnTop +ToolWindow"
    if (ownerHwnd)
        ownerOpt .= " +Owner" . ownerHwnd
    g_PresetMgrGui := Gui(ownerOpt, "Context preset manager")
    g_PresetMgrGui.SetFont("s10", "Segoe UI")
    PromptContextPresets_BuildManager()
    g_PresetMgrGui.OnEvent("Close", PromptContextPresets_CloseManager)
    g_PresetMgrGui.OnEvent("Escape", PromptContextPresets_CloseManager)
    UtilitySelector_DialogsBegin()
    mon := UtilitySelector_ActiveMonitorWorkArea()
    g_PresetMgrGui.Show("Hide")
    g_PresetMgrGui.GetPos(, , &gw, &gh)
    g_PresetMgrGui.Show("x" (mon.left + (mon.width - gw) // 2) " y" (mon.top + (mon.height - gh) // 2))
    try WinWaitClose("ahk_id " g_PresetMgrGui.Hwnd)
    catch {
    }
    UtilitySelector_DialogsEnd()
}

PromptContextPresets_BuildManager() {
    global g_PresetMgrGui, g_PresetMgrListLv, g_PresetMgrNameCtrl
    global g_PresetMgrPersonalLv, g_PresetMgrWorkLv
    colW := 360
    pathCol := colW - 154
    g_PresetMgrGui.Add("Text", "xm w80", "Presets")
    g_PresetMgrListLv := g_PresetMgrGui.Add("ListView", "yp w240 r8", ["Name", "P", "W"])
    g_PresetMgrListLv.ModifyCol(1, 150)
    g_PresetMgrListLv.ModifyCol(2, 35)
    g_PresetMgrListLv.ModifyCol(3, 35)
    g_PresetMgrListLv.OnEvent("ItemFocus", PromptContextPresets_OnListFocus)
    g_PresetMgrListLv.OnEvent("Click", PromptContextPresets_OnListFocus)
    g_PresetMgrGui.Add("Button", "xm y+4 w70", "New").OnEvent("Click", PromptContextPresets_OnNew)
    g_PresetMgrGui.Add("Button", "x+6 yp w70", "Delete").OnEvent("Click", PromptContextPresets_OnDelete)
    g_PresetMgrGui.Add("Button", "x+6 yp w70", "Help").OnEvent("Click", PromptContextPresets_ShowHelp)

    g_PresetMgrGui.Add("Text", "x+16 yp-4 w80", "Name")
    g_PresetMgrNameCtrl := g_PresetMgrGui.Add("Edit", "yp w" . (colW * 2 - 80 - 256))

    g_PresetMgrGui.Add("Text", "xm section w" . colW, "Personal files")
    g_PresetMgrGui.Add("Text", "ys w" . colW, "Work files")
    g_PresetMgrPersonalLv := g_PresetMgrGui.Add("ListView", "xm w" . colW . " r7", ["Path", "Compact", "CSV keep"])
    g_PresetMgrWorkLv := g_PresetMgrGui.Add("ListView", "x+12 yp w" . colW . " r7", ["Path", "Compact", "CSV keep"])
    for lv in [g_PresetMgrPersonalLv, g_PresetMgrWorkLv] {
        lv.ModifyCol(1, pathCol)
        lv.ModifyCol(2, 50)
        lv.ModifyCol(3, 80)
    }
    g_PresetMgrPersonalLv.OnEvent("ItemFocus", (ctrl, info) => PromptContextPresets_OnSideFocus("personal", info))
    g_PresetMgrWorkLv.OnEvent("ItemFocus", (ctrl, info) => PromptContextPresets_OnSideFocus("work", info))
    g_PresetMgrPersonalLv.OnEvent("Click", (ctrl, info) => PromptContextPresets_OnSideFocus("personal", info))
    g_PresetMgrWorkLv.OnEvent("Click", (ctrl, info) => PromptContextPresets_OnSideFocus("work", info))

    g_PresetMgrGui.Add("Button", "xm w100", "Add files").OnEvent("Click", (*) => PromptContextPresets_OnAddFiles(
        "personal"))
    g_PresetMgrGui.Add("Button", "x+8 yp w100", "Paste paths").OnEvent("Click", (*) =>
        PromptContextPresets_OnPastePaths(
            "personal"))
    g_PresetMgrGui.Add("Button", "x+8 yp w80", "Remove").OnEvent("Click", (*) => PromptContextPresets_OnRemove(
        "personal"))
    g_PresetMgrGui.Add("Button", "x+24 yp w100", "Add files").OnEvent("Click", (*) => PromptContextPresets_OnAddFiles(
        "work"))
    g_PresetMgrGui.Add("Button", "x+8 yp w100", "Paste paths").OnEvent("Click", (*) =>
        PromptContextPresets_OnPastePaths(
            "work"))
    g_PresetMgrGui.Add("Button", "x+8 yp w80", "Remove").OnEvent("Click", (*) => PromptContextPresets_OnRemove("work"))

    personalCompact := g_PresetMgrGui.Add("CheckBox", "xm w" . colW, "Compact")
    workCompact := g_PresetMgrGui.Add("CheckBox", "x+12 yp w" . colW, "Compact")
    personalCompact.OnEvent("Click", (*) => PromptContextPresets_OnCompactClick("personal"))
    workCompact.OnEvent("Click", (*) => PromptContextPresets_OnCompactClick("work"))
    personalCsvFromLabel := g_PresetMgrGui.Add("Text", "xm w90 Hidden", "CSV keep from")
    personalCsvFrom := g_PresetMgrGui.Add("Edit", "yp w50 Number Hidden")
    personalCsvToLabel := g_PresetMgrGui.Add("Text", "yp w20 Hidden", "to")
    personalCsvTo := g_PresetMgrGui.Add("Edit", "yp w50 Number Hidden")
    workCsvFromLabel := g_PresetMgrGui.Add("Text", "x+162 yp w90 Hidden", "CSV keep from")
    workCsvFrom := g_PresetMgrGui.Add("Edit", "yp w50 Number Hidden")
    workCsvToLabel := g_PresetMgrGui.Add("Text", "yp w20 Hidden", "to")
    workCsvTo := g_PresetMgrGui.Add("Edit", "yp w50 Number Hidden")
    personalCsvFrom.OnEvent("Change", (*) => PromptContextPresets_OnCsvKeepChange("personal"))
    personalCsvTo.OnEvent("Change", (*) => PromptContextPresets_OnCsvKeepChange("personal"))
    workCsvFrom.OnEvent("Change", (*) => PromptContextPresets_OnCsvKeepChange("work"))
    workCsvTo.OnEvent("Change", (*) => PromptContextPresets_OnCsvKeepChange("work"))
    g_PresetMgrFlagCtrls := Map()
    g_PresetMgrFlagCtrls["personal"] := {
        compact: personalCompact,
        csvFromLabel: personalCsvFromLabel,
        csvFrom: personalCsvFrom,
        csvToLabel: personalCsvToLabel,
        csvTo: personalCsvTo,
        selected: 0
    }
    g_PresetMgrFlagCtrls["work"] := {
        compact: workCompact,
        csvFromLabel: workCsvFromLabel,
        csvFrom: workCsvFrom,
        csvToLabel: workCsvToLabel,
        csvTo: workCsvTo,
        selected: 0
    }

    g_PresetMgrGui.Add("Button", "xm w100 Default", "Save preset").OnEvent("Click", PromptContextPresets_OnSavePreset)
    g_PresetMgrGui.Add("Button", "x+8 yp w100", "Close").OnEvent("Click", PromptContextPresets_CloseManager)
    PromptContextPresets_ReloadManagerList()
    PromptContextPresets_OnNew()
}

PromptContextPresets_MgrPathsRef(side) {
    global g_PresetMgrPersonalPaths, g_PresetMgrWorkPaths
    return (side = "work") ? g_PresetMgrWorkPaths : g_PresetMgrPersonalPaths
}

PromptContextPresets_MgrSetPaths(side, arr) {
    global g_PresetMgrPersonalPaths, g_PresetMgrWorkPaths
    if (side = "work")
        g_PresetMgrWorkPaths := arr
    else
        g_PresetMgrPersonalPaths := arr
}

PromptContextPresets_MgrLv(side) {
    global g_PresetMgrPersonalLv, g_PresetMgrWorkLv
    return (side = "work") ? g_PresetMgrWorkLv : g_PresetMgrPersonalLv
}

PromptContextPresets_MgrSelectedRow(side) {
    lv := PromptContextPresets_MgrLv(side)
    if (!IsObject(lv))
        return 0
    row := lv.GetNext(0, "F")
    if (!row)
        row := lv.GetNext(0)
    return row
}

PromptContextPresets_MgrEventRow(info) {
    row := 0
    try {
        if (IsObject(info) && info.HasProp("EventInfo"))
            row := Integer(info.EventInfo)
    } catch {
        row := 0
    }
    return row
}

PromptContextPresets_MgrReloadSide(side) {
    lv := PromptContextPresets_MgrLv(side)
    if (!IsObject(lv))
        return
    selected := PromptContextPresets_MgrSelectedRow(side)
    lv.Delete()
    for e in PromptContextPresets_MgrPathsRef(side)
        lv.Add("", PromptData_ContextEntryPath(e), PromptData_ContextCompactLabel(e), PromptData_ContextCsvKeepLabel(e))
    if (selected >= 1 && selected <= PromptContextPresets_MgrPathsRef(side).Length)
        lv.Modify(selected, "Select Focus Vis")
    else
        PromptContextPresets_MgrLoadFlags(side, 0)
}

PromptContextPresets_MgrFlagCtrls(side) {
    global g_PresetMgrFlagCtrls
    if (!IsObject(g_PresetMgrFlagCtrls) || !g_PresetMgrFlagCtrls.Has(side))
        return false
    return g_PresetMgrFlagCtrls[side]
}

PromptContextPresets_MgrSetCsvVisible(side, visible) {
    ctrls := PromptContextPresets_MgrFlagCtrls(side)
    if (!IsObject(ctrls))
        return
    opt := visible ? "-Hidden" : "+Hidden"
    try ctrls.csvFromLabel.Opt(opt)
    try ctrls.csvFrom.Opt(opt)
    try ctrls.csvToLabel.Opt(opt)
    try ctrls.csvTo.Opt(opt)
    catch {
    }
}

PromptContextPresets_MgrLoadFlags(side, row) {
    global g_PresetMgrSuppressFlagSync, g_PresetMgrFlagCtrls
    ctrls := PromptContextPresets_MgrFlagCtrls(side)
    if (!IsObject(ctrls))
        return
    g_PresetMgrSuppressFlagSync := true
    arr := PromptContextPresets_MgrPathsRef(side)
    ctrls.selected := row
    if (row < 1 || row > arr.Length) {
        try ctrls.compact.Value := 0
        try ctrls.compact.Enabled := false
        try ctrls.csvFrom.Value := ""
        try ctrls.csvTo.Value := ""
        PromptContextPresets_MgrSetCsvVisible(side, false)
        g_PresetMgrSuppressFlagSync := false
        return
    }
    e := arr[row]
    try ctrls.compact.Enabled := true
    try ctrls.compact.Value := (e.HasProp("compact") && e.compact) ? 1 : 0
    isCsv := PromptData_IsCsvPath(PromptData_ContextEntryPath(e))
    PromptContextPresets_MgrSetCsvVisible(side, isCsv)
    if (isCsv) {
        from := (e.HasProp("csvKeepFrom") && e.csvKeepFrom >= 1) ? e.csvKeepFrom : ""
        to := (e.HasProp("csvKeepTo") && e.csvKeepTo >= 1) ? e.csvKeepTo : ""
        try ctrls.csvFrom.Value := from
        try ctrls.csvTo.Value := to
    } else {
        try ctrls.csvFrom.Value := ""
        try ctrls.csvTo.Value := ""
    }
    g_PresetMgrSuppressFlagSync := false
    g_PresetMgrFlagCtrls[side] := ctrls
}

PromptContextPresets_OnSideFocus(side, info) {
    row := PromptContextPresets_MgrEventRow(info)
    if (row < 1)
        row := PromptContextPresets_MgrSelectedRow(side)
    PromptContextPresets_MgrLoadFlags(side, row)
}

PromptContextPresets_MgrRefreshRow(side, row) {
    lv := PromptContextPresets_MgrLv(side)
    arr := PromptContextPresets_MgrPathsRef(side)
    if (!IsObject(lv) || row < 1 || row > arr.Length)
        return
    e := arr[row]
    lv.Modify(row, , PromptData_ContextEntryPath(e), PromptData_ContextCompactLabel(e), PromptData_ContextCsvKeepLabel(
        e))
}

PromptContextPresets_OnCompactClick(side) {
    global g_PresetMgrSuppressFlagSync
    if (g_PresetMgrSuppressFlagSync)
        return
    ctrls := PromptContextPresets_MgrFlagCtrls(side)
    row := PromptContextPresets_MgrSelectedRow(side)
    arr := PromptContextPresets_MgrPathsRef(side)
    if (row < 1 || row > arr.Length)
        return
    arr[row].compact := ctrls.compact.Value ? 1 : 0
    PromptContextPresets_MgrRefreshRow(side, row)
}

PromptContextPresets_OnCsvKeepChange(side) {
    global g_PresetMgrSuppressFlagSync
    if (g_PresetMgrSuppressFlagSync)
        return
    ctrls := PromptContextPresets_MgrFlagCtrls(side)
    row := PromptContextPresets_MgrSelectedRow(side)
    arr := PromptContextPresets_MgrPathsRef(side)
    if (row < 1 || row > arr.Length)
        return
    from := 0
    to := 0
    try from := Integer(Trim(ctrls.csvFrom.Value))
    catch {
    }
    try to := Integer(Trim(ctrls.csvTo.Value))
    catch {
    }
    if (from < 1 || to < 1 || from > to) {
        from := 0
        to := 0
    }
    arr[row].csvKeepFrom := from
    arr[row].csvKeepTo := to
    PromptContextPresets_MgrRefreshRow(side, row)
}

PromptContextPresets_MgrHasPath(arr, path) {
    norm := StrLower(PromptData_StripPathQuotes(path))
    for e in arr {
        if (StrLower(PromptData_ContextEntryPath(e)) = norm)
            return true
    }
    return false
}

PromptContextPresets_MgrAppendPaths(side, incoming) {
    arr := PromptContextPresets_MgrPathsRef(side)
    for raw in incoming {
        n := PromptData_StripPathQuotes(raw)
        if (n = "" || PromptContextPresets_MgrHasPath(arr, n))
            continue
        arr.Push(PromptData_NewContextEntry(n))
    }
    PromptContextPresets_MgrSetPaths(side, arr)
    PromptContextPresets_MgrReloadSide(side)
}

PromptContextPresets_OnAddFiles(side) {
    global g_PresetMgrGui
    if (IsObject(g_PresetMgrGui)) {
        try g_PresetMgrGui.Opt("-AlwaysOnTop")
        catch {
        }
    }
    selected := FileSelect("M1", , "Select context files")
    if (IsObject(g_PresetMgrGui)) {
        try g_PresetMgrGui.Opt("+AlwaysOnTop")
        try WinActivate("ahk_id " g_PresetMgrGui.Hwnd)
        catch {
        }
    }
    PromptContextPresets_MgrAppendPaths(side, PromptEditor_ParseFileSelect(selected))
}

PromptContextPresets_OnPastePaths(side) {
    raw := ""
    try raw := A_Clipboard
    catch {
        return
    }
    PromptContextPresets_MgrAppendPaths(side, PromptData_ParsePathList(raw))
}

PromptContextPresets_OnRemove(side) {
    lv := PromptContextPresets_MgrLv(side)
    if (!IsObject(lv))
        return
    selected := []
    row := 0
    loop {
        row := lv.GetNext(row)
        if (!row)
            break
        selected.Push(Integer(row))
    }
    if (selected.Length = 0)
        return
    arr := PromptContextPresets_MgrPathsRef(side)
    keep := []
    skip := Map()
    for idx in selected
        skip[idx] := true
    loop arr.Length {
        if (!skip.Has(A_Index))
            keep.Push(arr[A_Index])
    }
    PromptContextPresets_MgrSetPaths(side, keep)
    PromptContextPresets_MgrReloadSide(side)
}

PromptContextPresets_ReloadManagerList() {
    global g_PresetMgrListLv, g_PresetMgrPresets
    if (!IsObject(g_PresetMgrListLv))
        return
    g_PresetMgrPresets := PromptContextPresets_LoadAll()
    g_PresetMgrListLv.Delete()
    for p in g_PresetMgrPresets
        g_PresetMgrListLv.Add("", p.name, p.personal.Length, p.work.Length)
}

PromptContextPresets_LoadFormFromPreset(preset) {
    global g_PresetMgrNameCtrl, g_PresetMgrPersonalPaths, g_PresetMgrWorkPaths
    if (!IsObject(preset))
        return
    try g_PresetMgrNameCtrl.Value := preset.name
    catch {
    }
    g_PresetMgrPersonalPaths := PromptData_ParseContextEntries(preset.personal)
    g_PresetMgrWorkPaths := PromptData_ParseContextEntries(preset.work)
    PromptContextPresets_MgrReloadSide("personal")
    PromptContextPresets_MgrReloadSide("work")
}

PromptContextPresets_OnListFocus(*) {
    global g_PresetMgrListLv, g_PresetMgrPresets, g_PresetMgrEditIndex
    row := g_PresetMgrListLv.GetNext(0, "F")
    if (!row)
        row := g_PresetMgrListLv.GetNext(0)
    if (!row || row > g_PresetMgrPresets.Length)
        return
    g_PresetMgrEditIndex := row
    PromptContextPresets_LoadFormFromPreset(g_PresetMgrPresets[row])
}

PromptContextPresets_OnNew(*) {
    global g_PresetMgrNameCtrl, g_PresetMgrEditIndex, g_PresetMgrPersonalPaths, g_PresetMgrWorkPaths
    g_PresetMgrEditIndex := 0
    try g_PresetMgrNameCtrl.Value := ""
    catch {
    }
    g_PresetMgrPersonalPaths := []
    g_PresetMgrWorkPaths := []
    PromptContextPresets_MgrReloadSide("personal")
    PromptContextPresets_MgrReloadSide("work")
}

PromptContextPresets_OnDelete(*) {
    global g_PresetMgrListLv, g_PresetMgrPresets, g_PresetMgrEditIndex
    row := g_PresetMgrListLv.GetNext(0, "F")
    if (!row) {
        UtilitySelector_Notify("Select a preset to delete.")
        return
    }
    name := g_PresetMgrPresets[row].name
    if (MsgBox("Delete preset '" name "'?", "Delete preset", "YesNo Icon!") != "Yes")
        return
    keep := []
    idx := 0
    for p in g_PresetMgrPresets {
        idx += 1
        if (idx != row)
            keep.Push(p)
    }
    if (!PromptContextPresets_SaveAll(keep)) {
        UtilitySelector_Notify("Could not save preset file.")
        return
    }
    PromptContextPresets_ReloadManagerList()
    PromptContextPresets_OnNew()
    PromptEditor_RefreshPresetDropdowns()
}

PromptContextPresets_OnSavePreset(*) {
    global g_PresetMgrNameCtrl, g_PresetMgrPresets, g_PresetMgrEditIndex
    global g_PresetMgrPersonalPaths, g_PresetMgrWorkPaths
    name := PromptContextPresets_NormalizeName(g_PresetMgrNameCtrl.Value)
    if (name = "") {
        UtilitySelector_Notify("Preset name is required.")
        return
    }
    personal := PromptData_ParseContextEntries(g_PresetMgrPersonalPaths)
    work := PromptData_ParseContextEntries(g_PresetMgrWorkPaths)
    if (personal.Length = 0 && work.Length = 0) {
        UtilitySelector_Notify("Add at least one context file to personal or work list.")
        return
    }
    if (PromptContextPresets_NameTaken(name, g_PresetMgrEditIndex)) {
        UtilitySelector_Notify("A preset with that name already exists.")
        return
    }
    entry := { name: name, personal: personal, work: work }
    list := []
    if (g_PresetMgrEditIndex >= 1 && g_PresetMgrEditIndex <= g_PresetMgrPresets.Length) {
        idx := 0
        for p in g_PresetMgrPresets {
            idx += 1
            list.Push(idx = g_PresetMgrEditIndex ? entry : p)
        }
    } else {
        for p in g_PresetMgrPresets
            list.Push(p)
        list.Push(entry)
    }
    if (!PromptContextPresets_SaveAll(list)) {
        UtilitySelector_Notify("Could not save preset file.")
        return
    }
    PromptContextPresets_ReloadManagerList()
    g_PresetMgrPresets := PromptContextPresets_LoadAll()
    loop g_PresetMgrPresets.Length {
        if (g_PresetMgrPresets[A_Index].name = name) {
            g_PresetMgrEditIndex := A_Index
            try g_PresetMgrListLv.Modify(A_Index, "Select Focus Vis")
            catch {
            }
            break
        }
    }
    PromptEditor_RefreshPresetDropdowns()
    UtilitySelector_Notify("Preset saved.")
}

PromptContextPresets_CloseManager(*) {
    global g_PresetMgrGui, g_PresetMgrOnClose
    PromptContextPresets_CloseHelp()
    if (IsObject(g_PresetMgrGui)) {
        try g_PresetMgrGui.Destroy()
        catch {
        }
    }
    g_PresetMgrGui := false
    PromptEditor_RefreshPresetDropdowns()
    if (g_PresetMgrOnClose != "") {
        try g_PresetMgrOnClose.Call()
        catch {
        }
    }
    g_PresetMgrOnClose := ""
}
