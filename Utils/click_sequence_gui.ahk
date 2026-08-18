; =============================================================================
; Utils module: click_sequence_gui.ahk
; Keyboard-navigable CRUD: Macros → Slots → Sibling Sequences → Clicks → Aliases.
; =============================================================================

global g_ClickSeqGui := false
global g_ClickSeqLv := false
global g_ClickSeqHint := false
global g_ClickSeqHeader := false
global g_ClickSeqActive := false
global g_ClickSeqLevel := "macros"
global g_ClickSeqMacroId := ""
global g_ClickSeqSlotIndex := 0
global g_ClickSeqGroupId := ""
global g_ClickSeqSeqIndex := 0
global g_ClickSeqClickIndex := 0
global g_ClickSeqRows := []
global g_ClickSeqHotkeys := []
global g_ClickSeqFormKind := "name"
global g_ClickSeqFormMatch := "exact"
global g_ClickSeqFormHwnd := 0

ClickSeqGui_Launch() {
    global g_ClickSeqLevel, g_ClickSeqMacroId, g_ClickSeqSlotIndex, g_ClickSeqGroupId, g_ClickSeqSeqIndex, g_ClickSeqClickIndex
    ClickSeqData_Load(true)
    g_ClickSeqLevel := "macros"
    g_ClickSeqMacroId := ""
    g_ClickSeqSlotIndex := 0
    g_ClickSeqGroupId := ""
    g_ClickSeqSeqIndex := 0
    g_ClickSeqClickIndex := 0
    ClickSeqGui_Rebuild()
}

ClickSeqGui_Close() {
    global g_ClickSeqGui, g_ClickSeqActive, g_ClickSeqLv, g_ClickSeqHint, g_ClickSeqHeader
    ClickSeqGui_UnbindHotkeys()
    if (IsObject(g_ClickSeqGui)) {
        try g_ClickSeqGui.Destroy()
        catch {
        }
    }
    g_ClickSeqGui := false
    g_ClickSeqLv := false
    g_ClickSeqHint := false
    g_ClickSeqHeader := false
    g_ClickSeqActive := false
}

ClickSeqGui_UnbindHotkeys() {
    global g_ClickSeqHotkeys
    try HotIf(ClickSeqGui_HotIfKeys)
    catch {
    }
    for item in g_ClickSeqHotkeys {
        try Hotkey(item, "Off")
        catch {
        }
    }
    g_ClickSeqHotkeys := []
    try HotIf()
    catch {
    }
}

ClickSeqGui_GuiFocusIsEdit() {
    global g_ClickSeqGui
    if (!IsObject(g_ClickSeqGui))
        return false
    try {
        focused := ControlGetFocus("ahk_id " g_ClickSeqGui.Hwnd)
        return (focused != "" && InStr(focused, "Edit") = 1)
    } catch {
        return false
    }
}

ClickSeqGui_FormHotIfKeys(*) {
    global g_ClickSeqFormHwnd
    if (!g_ClickSeqFormHwnd)
        return false
    try {
        if (!WinActive("ahk_id " g_ClickSeqFormHwnd))
            return false
        focused := ControlGetFocus("ahk_id " g_ClickSeqFormHwnd)
        return !(focused != "" && InStr(focused, "Edit") = 1)
    } catch {
        return false
    }
}

ClickSeqGui_HotIfKeys(*) {
    global g_ClickSeqGui, g_ClickSeqActive
    if (!g_ClickSeqActive || !IsObject(g_ClickSeqGui))
        return false
    try {
        if (!WinActive("ahk_id " g_ClickSeqGui.Hwnd))
            return false
        return !ClickSeqGui_GuiFocusIsEdit()
    } catch {
        return false
    }
}

ClickSeqGui_BindHotkeys(pairs) {
    global g_ClickSeqGui, g_ClickSeqHotkeys
    ClickSeqGui_UnbindHotkeys()
    if (!IsObject(g_ClickSeqGui))
        return
    try HotIf(ClickSeqGui_HotIfKeys)
    catch {
        return
    }
    for p in pairs {
        try {
            Hotkey(p[1], p[2], "On I10")
            g_ClickSeqHotkeys.Push(p[1])
        } catch {
        }
    }
    try HotIf()
    catch {
    }
}

ClickSeqGui_Center(guiObj, w := 920, h := 560) {
    ml := 0
    mt := 0
    mr := 0
    mb := 0
    try GetActiveMonitorWorkArea_StandardBar(&ml, &mt, &mr, &mb)
    catch {
        MonitorGetWorkArea(MonitorGetPrimary(), &ml, &mt, &mr, &mb)
    }
    x := ml + ((mr - ml) - w) // 2
    y := mt + ((mb - mt) - h) // 2
    guiObj.Show("x" . x . " y" . y . " w" . w . " h" . h)
}

ClickSeqGui_Notify(msg, ms := 1600, accent := "") {
    if (accent = "")
        accent := BANNER_ACCENT_INFO
    try ShowCenteredOverlay_Utils(msg, ms, accent)
    catch {
    }
}

ClickSeqGui_OwnerHwnd() {
    global g_ClickSeqGui
    hwnd := 0
    try {
        if (IsObject(g_ClickSeqGui))
            hwnd := g_ClickSeqGui.Hwnd
    } catch {
        hwnd := 0
    }
    return hwnd
}

ClickSeqGui_OwnerOpt() {
    hwnd := ClickSeqGui_OwnerHwnd()
    return hwnd ? " Owner" . hwnd : ""
}

ClickSeqGui_GuiOwner() {
    hwnd := ClickSeqGui_OwnerHwnd()
    return hwnd ? " +Owner" . hwnd : ""
}

ClickSeqGui_DialogsBegin() {
    global g_ClickSeqGui
    try {
        if (IsObject(g_ClickSeqGui))
            g_ClickSeqGui.Opt("-AlwaysOnTop")
    } catch {
    }
}

ClickSeqGui_DialogsEnd() {
    global g_ClickSeqGui
    try {
        if (IsObject(g_ClickSeqGui))
            g_ClickSeqGui.Opt("+AlwaysOnTop")
    } catch {
    }
}

ClickSeqGui_Confirm(msg) {
    ClickSeqGui_DialogsBegin()
    result := MsgBox(msg, "Click Sequences", "YesNo Icon?" . ClickSeqGui_OwnerOpt())
    ClickSeqGui_DialogsEnd()
    return result = "Yes"
}

ClickSeqGui_HintForLevel(level) {
    switch level {
        case "macros":
            return "[Enter] slots   E rename   [R] rules   [V] map   F1/? help   Esc close"
        case "slots":
            return "[S] Sequence Group   [H] Hardcoded Script   E edit   Delete   U/J   Enter siblings   Backspace   Esc"
        case "sequences":
            return "[I]/[A]/Insert add Sibling   E edit   Delete   U/J   Enter clicks   Backspace   Esc"
        case "clicks":
            return "[I]/[A]/Insert add   E edit   Delete   U/J   Enter aliases   Backspace   Esc"
        case "selectors":
            return "[I]/[A]/Insert add Alias   E edit   Delete   U/J   Backspace   Esc"
        default:
            return "Esc close"
    }
}

ClickSeqGui_TitleForLevel() {
    global g_ClickSeqLevel, g_ClickSeqMacroId, g_ClickSeqSeqIndex, g_ClickSeqClickIndex, g_ClickSeqGroupId, g_ClickSeqSlotIndex
    if (g_ClickSeqLevel = "macros")
        return "Click Sequences"
    macro := ClickSeqData_MacroById(g_ClickSeqMacroId)
    mname := IsObject(macro) ? macro.name : g_ClickSeqMacroId
    if (g_ClickSeqLevel = "slots")
        return "Slots — " . mname
    if (g_ClickSeqLevel = "sequences")
        return "Sibling Sequences — " . mname . " / " . (g_ClickSeqGroupId != "" ? g_ClickSeqGroupId : "group")
    seq := ClickSeqGui_CurrentSequence()
    sname := IsObject(seq) ? seq.name : "#" . g_ClickSeqSeqIndex
    if (g_ClickSeqLevel = "clicks")
        return "Clicks — " . sname
    return "Aliases — " . sname . " click " . g_ClickSeqClickIndex
}

ClickSeqGui_CurrentMacro() {
    global g_ClickSeqMacroId
    return ClickSeqData_MacroById(g_ClickSeqMacroId)
}

ClickSeqGui_CurrentSequence() {
    global g_ClickSeqSeqIndex
    macro := ClickSeqGui_CurrentMacro()
    if (!IsObject(macro) || g_ClickSeqSeqIndex < 1 || g_ClickSeqSeqIndex > macro.sequences.Length)
        return ""
    return macro.sequences[g_ClickSeqSeqIndex]
}

ClickSeqGui_CurrentClick() {
    global g_ClickSeqClickIndex
    seq := ClickSeqGui_CurrentSequence()
    if (!IsObject(seq) || g_ClickSeqClickIndex < 1 || g_ClickSeqClickIndex > seq.clicks.Length)
        return ""
    return seq.clicks[g_ClickSeqClickIndex]
}

ClickSeqGui_CurrentSlot() {
    global g_ClickSeqSlotIndex
    macro := ClickSeqGui_CurrentMacro()
    if (!IsObject(macro) || !macro.HasProp("slots") || g_ClickSeqSlotIndex < 1 || g_ClickSeqSlotIndex > macro.slots.Length)
        return ""
    return macro.slots[g_ClickSeqSlotIndex]
}

ClickSeqGui_Rebuild() {
    global g_ClickSeqGui, g_ClickSeqLv, g_ClickSeqHint, g_ClickSeqHeader, g_ClickSeqActive, g_ClickSeqLevel
    ClickSeqGui_UnbindHotkeys()
    if (IsObject(g_ClickSeqGui)) {
        try g_ClickSeqGui.Destroy()
        catch {
        }
    }
    g_ClickSeqGui := Gui("+AlwaysOnTop +ToolWindow", ClickSeqGui_TitleForLevel())
    g_ClickSeqGui.SetFont("s10", "Segoe UI")
    g_ClickSeqHeader := g_ClickSeqGui.Add("Text", "x12 y10 w890 h22", ClickSeqGui_TitleForLevel())
    g_ClickSeqHint := g_ClickSeqGui.Add("Text", "x12 y34 w890 h22", ClickSeqGui_HintForLevel(g_ClickSeqLevel))
    cols := ["Col1", "Col2", "Col3", "Col4"]
    switch g_ClickSeqLevel {
        case "macros":
            cols := ["Name", "Trigger", "Slots", "Search order"]
        case "slots":
            cols := ["#", "Type", "Id", "Detail"]
        case "sequences":
            cols := ["#", "Context", "Name", "Preview"]
        case "clicks":
            cols := ["#", "Aliases", "Newest", "Settle ms"]
        case "selectors":
            cols := ["#", "Kind", "Match", "Value"]
    }
    g_ClickSeqLv := g_ClickSeqGui.Add("ListView", "x12 y62 w890 h460 Grid", cols)
    g_ClickSeqLv.OnEvent("DoubleClick", ClickSeqGui_OnEnter)
    g_ClickSeqGui.OnEvent("Close", (*) => ClickSeqGui_Close())
    g_ClickSeqGui.OnEvent("Escape", (*) => ClickSeqGui_Close())
    g_ClickSeqActive := true
    ClickSeqGui_Refresh()
    pairs := [
        ["$*Enter", ClickSeqGui_OnEnter],
        ["Enter", ClickSeqGui_OnEnter],
        ["$*Escape", (*) => ClickSeqGui_Close()],
        ["Escape", (*) => ClickSeqGui_Close()],
        ["$*Backspace", ClickSeqGui_OnBack],
        ["Backspace", ClickSeqGui_OnBack],
        ["$*i", ClickSeqGui_OnAdd],
        ["$*a", ClickSeqGui_OnAdd],
        ["Insert", ClickSeqGui_OnAdd],
        ["$*Insert", ClickSeqGui_OnAdd],
        ["$*e", ClickSeqGui_OnEdit],
        ["e", ClickSeqGui_OnEdit],
        ["Delete", ClickSeqGui_OnDelete],
        ["$*Delete", ClickSeqGui_OnDelete],
        ["$*u", (*) => ClickSeqGui_OnMove(-1)],
        ["u", (*) => ClickSeqGui_OnMove(-1)],
        ["$*j", (*) => ClickSeqGui_OnMove(1)],
        ["j", (*) => ClickSeqGui_OnMove(1)],
        ["$*s", ClickSeqGui_OnAddSeqGroup],
        ["s", ClickSeqGui_OnAddSeqGroup],
        ["$*h", ClickSeqGui_OnAddHardcoded],
        ["h", ClickSeqGui_OnAddHardcoded],
        ["$*r", ClickSeqGui_OnRules],
        ["r", ClickSeqGui_OnRules],
        ["$*v", ClickSeqGui_OnMap],
        ["v", ClickSeqGui_OnMap],
        ["$*F1", ClickSeqGui_ShowHelp],
        ["F1", ClickSeqGui_ShowHelp],
        ["$*?", ClickSeqGui_ShowHelp],
        ["?", ClickSeqGui_ShowHelp]
    ]
    ClickSeqGui_BindHotkeys(pairs)
    ClickSeqGui_Center(g_ClickSeqGui, 920, 560)
    try g_ClickSeqLv.Focus()
    catch {
    }
}

ClickSeqGui_Refresh() {
    global g_ClickSeqLv, g_ClickSeqRows, g_ClickSeqLevel, g_ClickSeqHeader, g_ClickSeqHint, g_ClickSeqGui
    if (!IsObject(g_ClickSeqLv))
        return
    if (IsObject(g_ClickSeqHeader))
        g_ClickSeqHeader.Value := ClickSeqGui_TitleForLevel()
    if (IsObject(g_ClickSeqHint))
        g_ClickSeqHint.Value := ClickSeqGui_HintForLevel(g_ClickSeqLevel)
    try g_ClickSeqGui.Title := ClickSeqGui_TitleForLevel()
    catch {
    }
    g_ClickSeqLv.Delete()
    g_ClickSeqRows := []
    switch g_ClickSeqLevel {
        case "macros":
            ClickSeqGui_FillMacros()
        case "slots":
            ClickSeqGui_FillSlots()
        case "sequences":
            ClickSeqGui_FillSequences()
        case "clicks":
            ClickSeqGui_FillClicks()
        case "selectors":
            ClickSeqGui_FillSelectors()
    }
    loop 4
        try g_ClickSeqLv.ModifyCol(A_Index, "AutoHdr")
    if (g_ClickSeqRows.Length > 0) {
        try g_ClickSeqLv.Modify(1, "Select Focus Vis")
        catch {
        }
    }
}

ClickSeqGui_FillMacros() {
    global g_ClickSeqLv, g_ClickSeqRows
    for macro in ClickSeqData_Load() {
        ClickSeqData_EnsureSlots(macro)
        n := (macro.HasProp("slots") && IsObject(macro.slots)) ? macro.slots.Length : 0
        order := "bottomUp"
        if (macro.HasProp("rules") && IsObject(macro.rules) && macro.rules.HasProp("searchOrder"))
            order := macro.rules.searchOrder
        g_ClickSeqRows.Push(macro)
        g_ClickSeqLv.Add("", macro.name, macro.trigger, n, order)
    }
}

ClickSeqGui_FillSlots() {
    global g_ClickSeqLv, g_ClickSeqRows
    macro := ClickSeqGui_CurrentMacro()
    if (!IsObject(macro))
        return
    ClickSeqData_EnsureSlots(macro)
    idx := 1
    for slot in macro.slots {
        g_ClickSeqRows.Push({ index: idx, slot: slot })
        if (slot.HasProp("type") && slot.type = "hardcoded") {
            sid := slot.HasProp("scriptId") ? slot.scriptId : ""
            detail := ClickSeqScript_Title(sid)
            g_ClickSeqLv.Add("", idx, "Hardcoded Script", sid, detail)
        } else {
            gid := slot.HasProp("groupId") ? slot.groupId : "clicks"
            n := 0
            gobj := ClickSeqData_SeqGroupById(macro, gid)
            if (IsObject(gobj) && gobj.HasProp("seqIndexes") && IsObject(gobj.seqIndexes))
                n := gobj.seqIndexes.Length
            g_ClickSeqLv.Add("", idx, "Sequence Group", gid, n . " sibling(s)")
        }
        idx += 1
    }
}

ClickSeqGui_FillSequences() {
    global g_ClickSeqLv, g_ClickSeqRows, g_ClickSeqGroupId
    macro := ClickSeqGui_CurrentMacro()
    if (!IsObject(macro))
        return
    for row in ClickSeqData_SequencesForGroup(macro, g_ClickSeqGroupId) {
        g_ClickSeqRows.Push({ index: row.index, seq: row.seq })
        preview := ClickSeqData_SequencePreview(row.seq)
        g_ClickSeqLv.Add("", row.index, ClickSeqData_ContextLabel(row.seq.context), row.seq.name, preview)
    }
}

ClickSeqGui_FillClicks() {
    global g_ClickSeqLv, g_ClickSeqRows
    seq := ClickSeqGui_CurrentSequence()
    if (!IsObject(seq))
        return
    idx := 1
    for click in seq.clicks {
        g_ClickSeqRows.Push({ index: idx, click: click })
        newest := (click.HasProp("preferNewest") && !click.preferNewest) ? "no" : "yes"
        settle := click.HasProp("settleMs") ? click.settleMs : CLICKSEQ_DEFAULT_SETTLE_MS
        g_ClickSeqLv.Add("", idx, ClickSeqData_ClickPreview(click), newest, settle)
        idx += 1
    }
}

ClickSeqGui_FillSelectors() {
    global g_ClickSeqLv, g_ClickSeqRows
    click := ClickSeqGui_CurrentClick()
    if (!IsObject(click))
        return
    idx := 1
    for sel in click.selectors {
        g_ClickSeqRows.Push({ index: idx, sel: sel })
        match := ClickSeqData_NormalizeMatch(sel.HasProp("match") ? sel.match : "exact")
        extra := (ClickSeqData_NormalizeKind(sel.kind) = "icon") ? " (not executed)" : ""
        g_ClickSeqLv.Add("", idx, ClickSeqData_KindLabel(sel.kind) . extra, match, sel.value)
        idx += 1
    }
}

ClickSeqGui_SelectedIndex() {
    global g_ClickSeqLv, g_ClickSeqRows
    if (!IsObject(g_ClickSeqLv))
        return 0
    row := 0
    try row := g_ClickSeqLv.GetNext(0)
    catch {
        return 0
    }
    if (!row || row > g_ClickSeqRows.Length)
        return 0
    return row
}

ClickSeqGui_OnEnter(*) {
    global g_ClickSeqLevel, g_ClickSeqMacroId, g_ClickSeqSeqIndex, g_ClickSeqClickIndex, g_ClickSeqRows, g_ClickSeqSlotIndex, g_ClickSeqGroupId
    row := ClickSeqGui_SelectedIndex()
    if (!row)
        return
    item := g_ClickSeqRows[row]
    if (g_ClickSeqLevel = "macros") {
        g_ClickSeqMacroId := item.id
        g_ClickSeqLevel := "slots"
        ClickSeqGui_Rebuild()
        return
    }
    if (g_ClickSeqLevel = "slots") {
        slot := item.slot
        g_ClickSeqSlotIndex := item.index
        if (slot.HasProp("type") && slot.type = "hardcoded") {
            ClickSeqGui_HardcodedForm(true)
            return
        }
        g_ClickSeqGroupId := slot.HasProp("groupId") ? slot.groupId : "clicks"
        g_ClickSeqLevel := "sequences"
        ClickSeqGui_Rebuild()
        return
    }
    if (g_ClickSeqLevel = "sequences") {
        g_ClickSeqSeqIndex := item.index
        g_ClickSeqLevel := "clicks"
        ClickSeqGui_Rebuild()
        return
    }
    if (g_ClickSeqLevel = "clicks") {
        g_ClickSeqClickIndex := item.index
        g_ClickSeqLevel := "selectors"
        ClickSeqGui_Rebuild()
    }
}

ClickSeqGui_OnBack(*) {
    global g_ClickSeqLevel
    if (g_ClickSeqLevel = "macros") {
        ClickSeqGui_Close()
        return
    }
    if (g_ClickSeqLevel = "slots")
        g_ClickSeqLevel := "macros"
    else if (g_ClickSeqLevel = "sequences")
        g_ClickSeqLevel := "slots"
    else if (g_ClickSeqLevel = "clicks")
        g_ClickSeqLevel := "sequences"
    else if (g_ClickSeqLevel = "selectors")
        g_ClickSeqLevel := "clicks"
    ClickSeqGui_Rebuild()
}

ClickSeqGui_OnAdd(*) {
    global g_ClickSeqLevel, g_ClickSeqMacroId, g_ClickSeqRows
    switch g_ClickSeqLevel {
        case "macros":
            row := ClickSeqGui_SelectedIndex()
            if (!row && g_ClickSeqRows.Length = 1)
                row := 1
            if (!row) {
                ClickSeqGui_Notify("Select a Shortcut, then Enter to open Slots.", 2200,
                    BANNER_ACCENT_INTERMEDIATE)
                return
            }
            g_ClickSeqMacroId := g_ClickSeqRows[row].id
            g_ClickSeqLevel := "slots"
            ClickSeqGui_Rebuild()
        case "slots":
            ClickSeqGui_Notify("Use [S] Sequence Group or [H] Hardcoded Script.", 2200,
                BANNER_ACCENT_INTERMEDIATE)
        case "sequences":
            ClickSeqGui_SequenceForm(false)
        case "clicks":
            ClickSeqGui_ClickForm(false)
        case "selectors":
            ClickSeqGui_SelectorForm(false)
    }
}

ClickSeqGui_OnEdit(*) {
    global g_ClickSeqLevel
    row := ClickSeqGui_SelectedIndex()
    if (!row && g_ClickSeqLevel != "macros") {
        ClickSeqGui_Notify("Select a row", 1200, BANNER_ACCENT_ERROR)
        return
    }
    switch g_ClickSeqLevel {
        case "macros":
            ClickSeqGui_MacroRename()
        case "slots":
            ClickSeqGui_EditSlot()
        case "sequences":
            ClickSeqGui_SequenceForm(true)
        case "clicks":
            ClickSeqGui_ClickForm(true)
        case "selectors":
            ClickSeqGui_SelectorForm(true)
    }
}

ClickSeqGui_OnDelete(*) {
    global g_ClickSeqLevel, g_ClickSeqRows
    if (g_ClickSeqLevel = "macros") {
        ClickSeqGui_Notify("Shortcuts cannot be deleted in this version.", 2000, BANNER_ACCENT_INTERMEDIATE)
        return
    }
    row := ClickSeqGui_SelectedIndex()
    if (!row)
        return
    item := g_ClickSeqRows[row]
    label := "this item"
    if (g_ClickSeqLevel = "slots") {
        if (item.slot.HasProp("type") && item.slot.type = "hardcoded")
            label := "Hardcoded Script " . (item.slot.HasProp("scriptId") ? item.slot.scriptId : "")
        else
            label := "Sequence Group " . (item.slot.HasProp("groupId") ? item.slot.groupId : "")
    } else if (g_ClickSeqLevel = "sequences")
        label := item.seq.name != "" ? item.seq.name : "sequence " . item.index
    else if (g_ClickSeqLevel = "clicks")
        label := "click " . item.index
    else
        label := "alias " . item.index
    if (!ClickSeqGui_Confirm("Delete " . label . "?"))
        return
    macro := ClickSeqGui_CurrentMacro()
    if (!IsObject(macro))
        return
    if (g_ClickSeqLevel = "slots") {
        slot := item.slot
        if (slot.HasProp("type") && slot.type = "seqGroup") {
            gid := slot.HasProp("groupId") ? slot.groupId : ""
            next := []
            if (macro.HasProp("seqGroups") && IsObject(macro.seqGroups)) {
                for g in macro.seqGroups {
                    if (g.HasProp("id") && g.id = gid)
                        continue
                    next.Push(g)
                }
                macro.seqGroups := next
            }
        }
        macro.slots.RemoveAt(item.index)
    } else if (g_ClickSeqLevel = "sequences") {
        ClickSeqData_RemoveSequenceAt(macro, item.index)
    } else if (g_ClickSeqLevel = "clicks") {
        seq := ClickSeqGui_CurrentSequence()
        if (!IsObject(seq))
            return
        seq.clicks.RemoveAt(item.index)
    } else {
        click := ClickSeqGui_CurrentClick()
        if (!IsObject(click))
            return
        click.selectors.RemoveAt(item.index)
    }
    ClickSeqData_ReplaceMacro(macro)
    ClickSeqGui_Refresh()
}

ClickSeqGui_OnMove(delta) {
    global g_ClickSeqLevel, g_ClickSeqGroupId
    if (g_ClickSeqLevel = "macros")
        return
    row := ClickSeqGui_SelectedIndex()
    if (!row)
        return
    dest := row + delta
    macro := ClickSeqGui_CurrentMacro()
    if (!IsObject(macro))
        return
    moved := false
    if (g_ClickSeqLevel = "slots") {
        moved := ClickSeqData_Swap(macro.slots, row, dest)
    } else if (g_ClickSeqLevel = "sequences") {
        g := ClickSeqData_SeqGroupById(macro, g_ClickSeqGroupId)
        if (IsObject(g) && g.HasProp("seqIndexes"))
            moved := ClickSeqData_Swap(g.seqIndexes, row, dest)
    } else if (g_ClickSeqLevel = "clicks") {
        seq := ClickSeqGui_CurrentSequence()
        if (IsObject(seq))
            moved := ClickSeqData_Swap(seq.clicks, row, dest)
    } else {
        click := ClickSeqGui_CurrentClick()
        if (IsObject(click))
            moved := ClickSeqData_Swap(click.selectors, row, dest)
    }
    if (!moved)
        return
    ClickSeqData_ReplaceMacro(macro)
    ClickSeqGui_Refresh()
    global g_ClickSeqLv
    try g_ClickSeqLv.Modify(dest, "Select Focus Vis")
    catch {
    }
}

ClickSeqGui_MacroRename() {
    global g_ClickSeqRows
    row := ClickSeqGui_SelectedIndex()
    if (!row)
        return
    macro := g_ClickSeqRows[row]
    owner := ClickSeqGui_GuiOwner()
    ClickSeqGui_DialogsBegin()
    g := Gui("+AlwaysOnTop +ToolWindow" . owner, "Rename macro")
    g.SetFont("s10", "Segoe UI")
    g.Add("Text", , "Name  (trigger stays " . macro.trigger . ")")
    eName := g.Add("Edit", "w360", macro.name)
    saved := false
    g.Add("Button", "y+16 w100 Default", "Save").OnEvent("Click", SaveMacro)
    g.Add("Button", "x+8 w100", "Cancel").OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())
    g.Show()
    try WinWaitClose("ahk_id " g.Hwnd)
    catch {
    }
    ClickSeqGui_DialogsEnd()
    if (saved)
        ClickSeqGui_Refresh()

    SaveMacro(*) {
        name := Trim(eName.Value)
        if (name = "") {
            MsgBox("Name is required.", "Click Sequences", "Icon!")
            return
        }
        list := ClickSeqData_Load()
        for m in list {
            if (m.id = macro.id) {
                m.name := name
                break
            }
        }
        ClickSeqData_Save(list)
        saved := true
        g.Destroy()
    }
}

ClickSeqGui_SequenceForm(isEdit) {
    global g_ClickSeqRows
    existing := ""
    if (isEdit) {
        row := ClickSeqGui_SelectedIndex()
        if (!row)
            return
        existing := g_ClickSeqRows[row].seq
    }
    ctxDefault := "enterprise"
    try ctxDefault := ResolveGlobalAICompanion()
    catch {
        ctxDefault := "enterprise"
    }
    owner := ClickSeqGui_GuiOwner()
    ClickSeqGui_DialogsBegin()
    g := Gui("+AlwaysOnTop +ToolWindow" . owner, isEdit ? "Edit sequence" : "Add sequence")
    g.SetFont("s10", "Segoe UI")
    g.Add("Text", , "Name")
    eName := g.Add("Edit", "w360", isEdit ? existing.name : "")
    g.Add("Text", "y+10", "Context  (gemini  |  enterprise  |  copilot  |  *)")
    eCtx := g.Add("Edit", "w360", isEdit ? existing.context : ctxDefault)
    saved := false
    g.Add("Button", "y+16 w100 Default", "Save").OnEvent("Click", SaveSeq)
    g.Add("Button", "x+8 w100", "Cancel").OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())
    g.Show()
    try WinWaitClose("ahk_id " g.Hwnd)
    catch {
    }
    ClickSeqGui_DialogsEnd()
    if (saved)
        ClickSeqGui_Refresh()

    SaveSeq(*) {
        global g_ClickSeqRows, g_ClickSeqGroupId
        name := Trim(eName.Value)
        if (name = "") {
            MsgBox("Name is required.", "Click Sequences", "Icon!")
            return
        }
        macro := ClickSeqGui_CurrentMacro()
        if (!IsObject(macro))
            return
        ctx := ClickSeqData_NormalizeContext(eCtx.Value)
        if (isEdit) {
            row := ClickSeqGui_SelectedIndex()
            seq := macro.sequences[g_ClickSeqRows[row].index]
            seq.name := name
            seq.context := ctx
        } else {
            seq := ClickSeqData_NewSequence(name, ctx)
            ClickSeqData_AddSequenceToGroup(macro, g_ClickSeqGroupId, seq)
        }
        ClickSeqData_ReplaceMacro(macro)
        saved := true
        g.Destroy()
    }
}

ClickSeqGui_ClickForm(isEdit) {
    global g_ClickSeqRows, g_ClickSeqClickIndex, g_ClickSeqLevel
    existing := ""
    if (isEdit) {
        row := ClickSeqGui_SelectedIndex()
        if (!row)
            return
        existing := g_ClickSeqRows[row].click
    }
    owner := ClickSeqGui_GuiOwner()
    ClickSeqGui_DialogsBegin()
    g := Gui("+AlwaysOnTop +ToolWindow" . owner, isEdit ? "Edit click" : "Add click")
    g.SetFont("s10", "Segoe UI")
    g.Add("Text", , "Control type  (Button, MenuItem, …)")
    eType := g.Add("Edit", "w360", isEdit ? existing.controlType : "Button")
    g.Add("Text", "y+10", "ClassContains AND-filter  (optional)")
    eClass := g.Add("Edit", "w360", isEdit ? existing.classContains : "")
    g.Add("Text", "y+10", "Settle ms after click")
    eSettle := g.Add("Edit", "w120", isEdit ? existing.settleMs : CLICKSEQ_DEFAULT_SETTLE_MS)
    chkNewest := g.Add("Checkbox", "y+12", "Prefer visually newest match (chat feeds)")
    chkNewest.Value := isEdit ? (existing.preferNewest ? 1 : 0) : 1
    saved := false
    addedNew := false
    g.Add("Button", "y+16 w100 Default", "Save").OnEvent("Click", SaveClick)
    g.Add("Button", "x+8 w100", "Cancel").OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())
    g.Show()
    try WinWaitClose("ahk_id " g.Hwnd)
    catch {
    }
    ClickSeqGui_DialogsEnd()
    if (saved) {
        ClickSeqGui_Refresh()
        if (addedNew) {
            seq := ClickSeqGui_CurrentSequence()
            if (IsObject(seq) && seq.clicks.Length > 0) {
                g_ClickSeqClickIndex := seq.clicks.Length
                g_ClickSeqLevel := "selectors"
                ClickSeqGui_Rebuild()
                ClickSeqGui_SelectorForm(false)
            }
        }
    }

    SaveClick(*) {
        seq := ClickSeqGui_CurrentSequence()
        if (!IsObject(seq))
            return
        settle := ClickSeqData_Int(eSettle.Value, CLICKSEQ_DEFAULT_SETTLE_MS)
        if (settle < 0)
            settle := 0
        ctype := Trim(eType.Value)
        if (ctype = "")
            ctype := "Button"
        if (isEdit) {
            row := ClickSeqGui_SelectedIndex()
            click := seq.clicks[row]
            click.controlType := ctype
            click.classContains := Trim(eClass.Value)
            click.settleMs := settle
            click.preferNewest := (chkNewest.Value = 1)
        } else {
            click := ClickSeqData_NewClick()
            click.controlType := ctype
            click.classContains := Trim(eClass.Value)
            click.settleMs := settle
            click.preferNewest := (chkNewest.Value = 1)
            seq.clicks.Push(click)
            addedNew := true
        }
        ClickSeqData_ReplaceMacro(ClickSeqGui_CurrentMacro())
        saved := true
        g.Destroy()
    }
}

ClickSeqGui_SelectorForm(isEdit) {
    global g_ClickSeqRows, g_ClickSeqFormKind, g_ClickSeqFormMatch, g_ClickSeqFormHwnd
    existing := ""
    if (isEdit) {
        row := ClickSeqGui_SelectedIndex()
        if (!row)
            return
        existing := g_ClickSeqRows[row].sel
        g_ClickSeqFormKind := ClickSeqData_NormalizeKind(existing.kind)
        g_ClickSeqFormMatch := ClickSeqData_NormalizeMatch(existing.match)
    } else {
        g_ClickSeqFormKind := "name"
        g_ClickSeqFormMatch := "exact"
    }
    owner := ClickSeqGui_GuiOwner()
    ClickSeqGui_DialogsBegin()
    g := Gui("+AlwaysOnTop +ToolWindow" . owner, isEdit ? "Edit Alias" : "Add Alias")
    g.SetFont("s10", "Segoe UI")
    g.Add("Text", "w420", "[N] name   [A] automationId   [C] classContains   [R] region   [I] icon")
    lblKind := g.Add("Text", "w420", "Kind: " . g_ClickSeqFormKind)
    g.Add("Text", "y+8 w420", "[X] exact   [B] substring  (name / automationId)")
    lblMatch := g.Add("Text", "w420", "Match: " . g_ClickSeqFormMatch)
    g.Add("Text", "y+10", "Value  (region: x,y,w,h relative to window)")
    eVal := g.Add("Edit", "w420", isEdit ? existing.value : "")
    saved := false
    g.Add("Button", "y+16 w100 Default", "Save").OnEvent("Click", SaveSel)
    g.Add("Button", "x+8 w100", "Cancel").OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())

    g.Show()
    g_ClickSeqFormHwnd := 0
    try g_ClickSeqFormHwnd := g.Hwnd
    catch {
        g_ClickSeqFormHwnd := 0
    }
    formKeys := ["n", "a", "c", "r", "i", "x", "b"]
    if (g_ClickSeqFormHwnd) {
        try HotIf(ClickSeqGui_FormHotIfKeys)
        try Hotkey("n", (*) => SetKind("name"), "On")
        try Hotkey("a", (*) => SetKind("automationId"), "On")
        try Hotkey("c", (*) => SetKind("classContains"), "On")
        try Hotkey("r", (*) => SetKind("region"), "On")
        try Hotkey("i", (*) => SetKind("icon"), "On")
        try Hotkey("x", (*) => SetMatch("exact"), "On")
        try Hotkey("b", (*) => SetMatch("substr"), "On")
        try HotIf()
    }
    try WinWaitClose("ahk_id " g.Hwnd)
    catch {
    }
    if (g_ClickSeqFormHwnd) {
        try HotIf(ClickSeqGui_FormHotIfKeys)
        for k in formKeys {
            try Hotkey(k, "Off")
        }
        try HotIf()
    }
    g_ClickSeqFormHwnd := 0
    ClickSeqGui_DialogsEnd()
    if (saved)
        ClickSeqGui_Refresh()

    SetKind(kind) {
        global g_ClickSeqFormKind
        g_ClickSeqFormKind := kind
        lblKind.Value := "Kind: " . kind
    }

    SetMatch(match) {
        global g_ClickSeqFormMatch
        g_ClickSeqFormMatch := match
        lblMatch.Value := "Match: " . match
    }

    SaveSel(*) {
        global g_ClickSeqFormKind, g_ClickSeqFormMatch
        value := Trim(eVal.Value)
        if (value = "") {
            MsgBox("Value is required.", "Click Sequences", "Icon!")
            return
        }
        click := ClickSeqGui_CurrentClick()
        if (!IsObject(click))
            return
        sel := ClickSeqData_NewSelector(g_ClickSeqFormKind, value, g_ClickSeqFormMatch)
        if (isEdit) {
            row := ClickSeqGui_SelectedIndex()
            click.selectors[row] := sel
        } else {
            click.selectors.Push(sel)
        }
        ClickSeqData_ReplaceMacro(ClickSeqGui_CurrentMacro())
        saved := true
        g.Destroy()
    }
}

ClickSeqGui_OnAddSeqGroup(*) {
    global g_ClickSeqLevel
    if (g_ClickSeqLevel != "slots")
        return
    macro := ClickSeqGui_CurrentMacro()
    if (!IsObject(macro))
        return
    ClickSeqData_EnsureSlots(macro)
    gid := ClickSeqData_UniqueGroupId(macro, "group")
    if (!macro.HasProp("seqGroups") || !IsObject(macro.seqGroups))
        macro.seqGroups := []
    macro.seqGroups.Push(ClickSeqData_NewSeqGroup(gid))
    macro.slots.Push(ClickSeqData_NewSeqGroupSlot(gid))
    ClickSeqData_ReplaceMacro(macro)
    ClickSeqGui_Refresh()
}

ClickSeqGui_OnAddHardcoded(*) {
    global g_ClickSeqLevel
    if (g_ClickSeqLevel != "slots")
        return
    ClickSeqGui_HardcodedForm(false)
}

ClickSeqGui_EditSlot() {
    global g_ClickSeqRows
    row := ClickSeqGui_SelectedIndex()
    if (!row)
        return
    slot := g_ClickSeqRows[row].slot
    if (slot.HasProp("type") && slot.type = "hardcoded") {
        ClickSeqGui_HardcodedForm(true)
        return
    }
    ClickSeqGui_SeqGroupRename()
}

ClickSeqGui_HardcodedForm(isEdit) {
    global g_ClickSeqRows
    catalog := ClickSeqScript_Catalog()
    owner := ClickSeqGui_GuiOwner()
    ClickSeqGui_DialogsBegin()
    g := Gui("+AlwaysOnTop +ToolWindow" . owner, isEdit ? "Edit Hardcoded Script" : "Add Hardcoded Script")
    g.SetFont("s10", "Segoe UI")
    g.Add("Text", "w420", "Pick a registered Hardcoded Script (not authored at runtime).")
    lv := g.Add("ListView", "w420 h220 Grid", ["Id", "Title"])
    for row in catalog
        lv.Add("", row.id, row.title)
    if (isEdit) {
        cur := g_ClickSeqRows[ClickSeqGui_SelectedIndex()].slot
        sid := cur.HasProp("scriptId") ? cur.scriptId : ""
        idx := 1
        for row in catalog {
            if (row.id = sid) {
                try lv.Modify(idx, "Select Focus Vis")
                catch {
                }
                break
            }
            idx += 1
        }
    } else if (catalog.Length > 0) {
        try lv.Modify(1, "Select Focus Vis")
        catch {
        }
    }
    saved := false
    g.Add("Button", "y+12 w100 Default", "Save").OnEvent("Click", SaveScript)
    g.Add("Button", "x+8 w100", "Cancel").OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())
    lv.OnEvent("DoubleClick", SaveScript)
    g.Show()
    try WinWaitClose("ahk_id " g.Hwnd)
    catch {
    }
    ClickSeqGui_DialogsEnd()
    if (saved)
        ClickSeqGui_Refresh()

    SaveScript(*) {
        pick := 0
        try pick := lv.GetNext(0)
        catch {
            pick := 0
        }
        if (!pick || pick > catalog.Length) {
            MsgBox("Select a Hardcoded Script.", "Click Sequences", "Icon!")
            return
        }
        sid := catalog[pick].id
        macro := ClickSeqGui_CurrentMacro()
        if (!IsObject(macro))
            return
        ClickSeqData_EnsureSlots(macro)
        if (isEdit) {
            row := ClickSeqGui_SelectedIndex()
            macro.slots[row] := ClickSeqData_NewHardcodedSlot(sid)
        } else {
            macro.slots.Push(ClickSeqData_NewHardcodedSlot(sid))
        }
        ClickSeqData_ReplaceMacro(macro)
        saved := true
        g.Destroy()
    }
}

ClickSeqGui_SeqGroupRename() {
    global g_ClickSeqRows
    row := ClickSeqGui_SelectedIndex()
    if (!row)
        return
    slot := g_ClickSeqRows[row].slot
    oldId := slot.HasProp("groupId") ? slot.groupId : "clicks"
    owner := ClickSeqGui_GuiOwner()
    ClickSeqGui_DialogsBegin()
    g := Gui("+AlwaysOnTop +ToolWindow" . owner, "Rename Sequence Group")
    g.SetFont("s10", "Segoe UI")
    g.Add("Text", , "Group id")
    eId := g.Add("Edit", "w280", oldId)
    saved := false
    g.Add("Button", "y+16 w100 Default", "Save").OnEvent("Click", SaveG)
    g.Add("Button", "x+8 w100", "Cancel").OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())
    g.Show()
    try WinWaitClose("ahk_id " g.Hwnd)
    catch {
    }
    ClickSeqGui_DialogsEnd()
    if (saved)
        ClickSeqGui_Refresh()

    SaveG(*) {
        nid := ClickSeqData_SanitizeId(eId.Value)
        if (nid = "") {
            MsgBox("Id is required.", "Click Sequences", "Icon!")
            return
        }
        macro := ClickSeqGui_CurrentMacro()
        if (!IsObject(macro))
            return
        grp := ClickSeqData_SeqGroupById(macro, oldId)
        if (IsObject(grp))
            grp.id := nid
        slot.groupId := nid
        ClickSeqData_ReplaceMacro(macro)
        saved := true
        g.Destroy()
    }
}

ClickSeqGui_OnRules(*) {
    global g_ClickSeqLevel, g_ClickSeqMacroId, g_ClickSeqRows
    macro := ""
    if (g_ClickSeqLevel = "macros") {
        row := ClickSeqGui_SelectedIndex()
        if (!row && g_ClickSeqRows.Length = 1)
            row := 1
        if (!row)
            return
        macro := g_ClickSeqRows[row]
        g_ClickSeqMacroId := macro.id
    } else {
        macro := ClickSeqGui_CurrentMacro()
    }
    if (!IsObject(macro))
        return
    ClickSeqData_EnsureSlots(macro)
    cur := "bottomUp"
    if (macro.HasProp("rules") && IsObject(macro.rules) && macro.rules.HasProp("searchOrder"))
        cur := macro.rules.searchOrder
    owner := ClickSeqGui_GuiOwner()
    ClickSeqGui_DialogsBegin()
    g := Gui("+AlwaysOnTop +ToolWindow" . owner, "Context Rules")
    g.SetFont("s10", "Segoe UI")
    g.Add("Text", "w400", "searchOrder  [B] bottomUp   [T] topDown   [F] firstMatch")
    lbl := g.Add("Text", "w400", "Current: " . cur)
    saved := false
    g.Add("Button", "y+16 w100 Default", "Save").OnEvent("Click", SaveRules)
    g.Add("Button", "x+8 w100", "Cancel").OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())
    g.Show()
    hwnd := 0
    try hwnd := g.Hwnd
    catch {
    }
    if (hwnd) {
        try HotIf((*) => WinActive("ahk_id " hwnd))
        try Hotkey("b", (*) => SetOrder("bottomUp"), "On")
        try Hotkey("t", (*) => SetOrder("topDown"), "On")
        try Hotkey("f", (*) => SetOrder("firstMatch"), "On")
        try HotIf()
    }
    try WinWaitClose("ahk_id " g.Hwnd)
    catch {
    }
    if (hwnd) {
        try HotIf((*) => WinActive("ahk_id " hwnd))
        try Hotkey("b", "Off")
        try Hotkey("t", "Off")
        try Hotkey("f", "Off")
        try HotIf()
    }
    ClickSeqGui_DialogsEnd()
    if (saved)
        ClickSeqGui_Refresh()

    SetOrder(v) {
        cur := v
        lbl.Value := "Current: " . v
    }

    SaveRules(*) {
        if (!macro.HasProp("rules") || !IsObject(macro.rules))
            macro.rules := ClickSeqData_DefaultRules()
        macro.rules.searchOrder := ClickSeqData_NormalizeSearchOrder(cur)
        ClickSeqData_ReplaceMacro(macro)
        saved := true
        g.Destroy()
    }
}

ClickSeqGui_ShowHelp(*) {
    path := ClickSeqData_VocabPath()
    text := ""
    try text := FileRead(path, "UTF-8")
    catch {
        text := "Vocabulary file not found:`n" . path
    }
    owner := ClickSeqGui_GuiOwner()
    ClickSeqGui_DialogsBegin()
    g := Gui("+AlwaysOnTop +ToolWindow" . owner, "Click Sequence vocabulary")
    g.SetFont("s10", "Segoe UI")
    g.Add("Edit", "ReadOnly w720 h480 -Wrap", text)
    g.Add("Button", "y+10 w100 Default", "Close").OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())
    g.Show()
    try WinWaitClose("ahk_id " g.Hwnd)
    catch {
    }
    ClickSeqGui_DialogsEnd()
}

ClickSeqGui_OnMap(*) {
    py := A_ScriptDir "\Utils\click_sequence_viz\generate_map.py"
    if (!FileExist(py)) {
        ClickSeqGui_Notify("generate_map.py not found", 2000, BANNER_ACCENT_ERROR)
        return
    }
    try StandardLoadingBar_Show("Building sequence map…", BANNER_ACCENT_INTERMEDIATE)
    catch {
    }
    cmd := 'python "' . py . '"'
    try {
        RunWait(cmd, A_ScriptDir, "Hide")
    } catch as e {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        ClickSeqGui_Notify("Python failed: " . e.Message, 2500, BANNER_ACCENT_ERROR)
        return
    }
    html := A_ScriptDir "\Utils\click_sequence_viz\output\map.html"
    try StandardLoadingBar_Hide(400)
    catch {
    }
    if (!FileExist(html)) {
        ClickSeqGui_Notify("map.html was not generated", 2200, BANNER_ACCENT_ERROR)
        return
    }
    Run('"' . html . '"')
}
