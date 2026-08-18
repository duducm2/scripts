; =============================================================================
; Utils module: click_sequence_data.ahk
; Persistent click-sequence registry (macro → sequence → click → selector aliases).
; Store: assets/data/click_sequences.ini  (same Load/Save mechanics as hotstring_data.ahk)
; Loaded at run/CRUD time, not by Act.
; =============================================================================

global g_ClickSeqMacros := []
global g_ClickSeqDataCacheReady := false
global g_ClickSeqDataCacheMtime := ""
global g_ClickSeqLastFailMessage := ""

CLICKSEQ_DEFAULT_MACRO_ID := "ai-quick-download"
CLICKSEQ_DEFAULT_SETTLE_MS := 600
CLICKSEQ_MAX_SEQ := 80
CLICKSEQ_MAX_CLICK := 40
CLICKSEQ_MAX_SELECTOR := 20

ClickSeqData_VocabPath() {
    return A_ScriptDir "\docs\click-sequence-vocabulary.md"
}

ClickSeqData_SlotLabel(slot) {
    if (!IsObject(slot))
        return ""
    if (slot.HasProp("type") && slot.type = "hardcoded")
        return "Hardcoded Script: " . (slot.HasProp("scriptId") ? slot.scriptId : "")
    return "Sequence Group: " . (slot.HasProp("groupId") ? slot.groupId : "clicks")
}

ClickSeqData_SequencesForGroup(macro, groupId) {
    out := []
    if (!IsObject(macro) || !macro.HasProp("sequences") || !IsObject(macro.sequences))
        return out
    g := ClickSeqData_SeqGroupById(macro, groupId)
    idxs := []
    if (IsObject(g) && g.HasProp("seqIndexes") && IsObject(g.seqIndexes) && g.seqIndexes.Length > 0)
        idxs := g.seqIndexes
    else {
        loop macro.sequences.Length
            idxs.Push(A_Index)
    }
    seen := Map()
    for n in idxs {
        if (n < 1 || n > macro.sequences.Length || seen.Has(n))
            continue
        seen[n] := true
        out.Push({ index: n, seq: macro.sequences[n] })
    }
    return out
}

ClickSeqData_IniPath() {
    return A_ScriptDir "\assets\data\click_sequences.ini"
}

ClickSeqData_NormalizeIniValue(val) {
    if (val = "" || val = "ERROR")
        return ""
    return val
}

ClickSeqData_Invalidate() {
    global g_ClickSeqMacros, g_ClickSeqDataCacheReady, g_ClickSeqDataCacheMtime
    g_ClickSeqMacros := []
    g_ClickSeqDataCacheReady := false
    g_ClickSeqDataCacheMtime := ""
}

ClickSeqData_FileMtime() {
    path := ClickSeqData_IniPath()
    if (!FileExist(path))
        return ""
    mtime := ""
    try mtime := FileGetTime(path, "M")
    catch {
        mtime := ""
    }
    return mtime
}

ClickSeqData_SanitizeId(id) {
    id := StrLower(Trim(id))
    id := RegExReplace(id, "[^a-z0-9\-_]", "")
    return id
}

ClickSeqData_MkClick(selValues, classContains := "", settleMs := unset) {
    if (!IsSet(settleMs))
        settleMs := CLICKSEQ_DEFAULT_SETTLE_MS
    click := ClickSeqData_NewClick()
    click.classContains := classContains
    click.settleMs := settleMs
    click.preferNewest := true
    click.selectors := []
    for v in selValues {
        if (IsObject(v))
            click.selectors.Push(v)
        else
            click.selectors.Push(ClickSeqData_NewSelector("name", v, "exact"))
    }
    return click
}

ClickSeqData_MkSeq(name, context, clicks) {
    seq := ClickSeqData_NewSequence(name, context)
    seq.clicks := clicks
    return seq
}

; Former #!+9 Quality Gates (personal Gemini A/B/file-or-code; Enterprise/Copilot labels from those stubs).
ClickSeqData_DefaultAiQuickDownloadSequences() {
    openSettle := 800
    seqs := []
    seqs.Push(ClickSeqData_MkSeq("Download code", "gemini", [ClickSeqData_MkClick(["Download code",
        "Baixar código", "Baixar codigo"])]))
    seqs.Push(ClickSeqData_MkSeq("Download code (icon)", "gemini", [ClickSeqData_MkClick([ClickSeqData_NewSelector(
        "name", "Download", "substr")], "mdc-icon-button")]))
    seqs.Push(ClickSeqData_MkSeq("Open then viewer Download", "gemini", [ClickSeqData_MkClick(["Open", "Abrir"],
    "open-button", openSettle), ClickSeqData_MkClick(["Download", "Baixar"], "drive-viewer")]))
    seqs.Push(ClickSeqData_MkSeq("Direct viewer Download", "gemini", [ClickSeqData_MkClick(["Download", "Baixar"],
    "drive-viewer")]))
    seqs.Push(ClickSeqData_MkSeq("Open then Download", "gemini", [ClickSeqData_MkClick(["Open", "Abrir"], "",
    openSettle), ClickSeqData_MkClick(["Download", "Baixar"])]))
    seqs.Push(ClickSeqData_MkSeq("Direct Download", "gemini", [ClickSeqData_MkClick(["Download", "Baixar"])]))
    seqs.Push(ClickSeqData_MkSeq("Open then Export", "enterprise", [ClickSeqData_MkClick(["Open", "Abrir"], "",
    openSettle), ClickSeqData_MkClick(["Export"])]))
    seqs.Push(ClickSeqData_MkSeq("Direct Export", "enterprise", [ClickSeqData_MkClick(["Export"])]))
    seqs.Push(ClickSeqData_MkSeq("Download code", "enterprise", [ClickSeqData_MkClick(["Download code",
        "Baixar código", "Baixar codigo"])]))
    seqs.Push(ClickSeqData_MkSeq("Open then Download", "enterprise", [ClickSeqData_MkClick(["Open", "Abrir"], "",
    openSettle), ClickSeqData_MkClick(["Download", "Baixar"])]))
    seqs.Push(ClickSeqData_MkSeq("Direct Download", "enterprise", [ClickSeqData_MkClick(["Download", "Baixar"])]))
    seqs.Push(ClickSeqData_MkSeq("Direct Download", "copilot", [ClickSeqData_MkClick(["Download", "Baixar"])]))
    seqs.Push(ClickSeqData_MkSeq("Download a copy", "copilot", [ClickSeqData_MkClick(["Download a copy", "Download",
        "Baixar"])]))
    idx := 1
    for seq in seqs {
        seq.order := idx
        idx += 1
    }
    return seqs
}

ClickSeqData_DefaultRules() {
    return { searchOrder: "bottomUp" }
}

ClickSeqData_DefaultSlots() {
    return [
        { type: "hardcoded", scriptId: "scrollFeedBottom" },
        { type: "seqGroup", groupId: "clicks" },
        { type: "hardcoded", scriptId: "desktopWait" },
        { type: "hardcoded", scriptId: "desktopCut" }
    ]
}

ClickSeqData_NormalizeSearchOrder(val) {
    v := StrLower(Trim(val))
    if (v = "topdown" || v = "top-down" || v = "top")
        return "topDown"
    if (v = "firstmatch" || v = "first" || v = "tree")
        return "firstMatch"
    return "bottomUp"
}

ClickSeqData_ParseRules(raw) {
    rules := ClickSeqData_DefaultRules()
    text := ClickSeqData_NormalizeIniValue(raw)
    if (text = "")
        return rules
    for part in StrSplit(text, "|") {
        part := Trim(part)
        pos := InStr(part, ":")
        if (!pos)
            continue
        k := StrLower(Trim(SubStr(part, 1, pos - 1)))
        v := Trim(SubStr(part, pos + 1))
        if (k = "searchorder")
            rules.searchOrder := ClickSeqData_NormalizeSearchOrder(v)
    }
    return rules
}

ClickSeqData_EncodeRules(rules) {
    order := "bottomUp"
    if (IsObject(rules) && rules.HasProp("searchOrder") && rules.searchOrder != "")
        order := rules.searchOrder
    return "searchOrder:" . order
}

ClickSeqData_ParseSlots(raw) {
    slots := []
    text := ClickSeqData_NormalizeIniValue(raw)
    if (text = "")
        return slots
    for part in StrSplit(text, "|") {
        part := Trim(part)
        pos := InStr(part, ":")
        if (!pos)
            continue
        kind := StrLower(Trim(SubStr(part, 1, pos - 1)))
        rest := Trim(SubStr(part, pos + 1))
        if (rest = "")
            continue
        if (kind = "hardcoded" || kind = "script")
            slots.Push({ type: "hardcoded", scriptId: ClickSeqData_SanitizeId(rest) })
        else if (kind = "seqgroup" || kind = "group")
            slots.Push({ type: "seqGroup", groupId: ClickSeqData_SanitizeId(rest) })
    }
    return slots
}

ClickSeqData_EncodeSlots(slots) {
    s := ""
    if (!IsObject(slots))
        return s
    for slot in slots {
        token := ""
        if (slot.HasProp("type") && slot.type = "hardcoded")
            token := "hardcoded:" . (slot.HasProp("scriptId") ? slot.scriptId : "")
        else
            token := "seqGroup:" . (slot.HasProp("groupId") ? slot.groupId : "clicks")
        if (token = "hardcoded:" || token = "seqGroup:")
            continue
        if (s != "")
            s .= "|"
        s .= token
    }
    return s
}

ClickSeqData_SeqGroupById(macro, groupId) {
    if (!IsObject(macro) || !macro.HasProp("seqGroups") || !IsObject(macro.seqGroups))
        return ""
    id := ClickSeqData_SanitizeId(groupId)
    for g in macro.seqGroups {
        if (g.HasProp("id") && g.id = id)
            return g
    }
    return ""
}

ClickSeqData_EnsureSlots(macro) {
    changed := false
    if (!IsObject(macro))
        return changed
    if (!macro.HasProp("rules") || !IsObject(macro.rules)) {
        macro.rules := ClickSeqData_DefaultRules()
        changed := true
    }
    slotsWereEmpty := (!macro.HasProp("slots") || !IsObject(macro.slots) || macro.slots.Length = 0)
    if (slotsWereEmpty) {
        macro.slots := ClickSeqData_DefaultSlots()
        changed := true
    }
    if (!macro.HasProp("seqGroups") || !IsObject(macro.seqGroups)) {
        macro.seqGroups := []
        changed := true
    }
    n := (macro.HasProp("sequences") && IsObject(macro.sequences)) ? macro.sequences.Length : 0
    for slot in macro.slots {
        if (!(slot.HasProp("type") && slot.type = "seqGroup"))
            continue
        gid := slot.HasProp("groupId") && slot.groupId != "" ? slot.groupId : "clicks"
        slot.groupId := gid
        if (IsObject(ClickSeqData_SeqGroupById(macro, gid)))
            continue
        idxs := []
        if (slotsWereEmpty && gid = "clicks") {
            loop n
                idxs.Push(A_Index)
        }
        macro.seqGroups.Push({ id: gid, seqIndexes: idxs })
        changed := true
    }
    g := ClickSeqData_SeqGroupById(macro, "clicks")
    if (slotsWereEmpty && IsObject(g) && n > 0 && (!g.HasProp("seqIndexes") || !IsObject(g.seqIndexes)
        || g.seqIndexes.Length = 0)) {
        g.seqIndexes := []
        loop n
            g.seqIndexes.Push(A_Index)
        changed := true
    }
    return changed
}

ClickSeqData_UniqueGroupId(macro, base := "group") {
    stem := ClickSeqData_SanitizeId(base)
    if (stem = "")
        stem := "group"
    n := 1
    loop 80 {
        cand := (n = 1) ? stem : stem . n
        if (!IsObject(ClickSeqData_SeqGroupById(macro, cand)))
            return cand
        n += 1
    }
    return stem . A_TickCount
}

ClickSeqData_AddSequenceToGroup(macro, groupId, seq) {
    if (!IsObject(macro) || !IsObject(seq))
        return 0
    if (!macro.HasProp("sequences") || !IsObject(macro.sequences))
        macro.sequences := []
    seq.order := macro.sequences.Length + 1
    macro.sequences.Push(seq)
    idx := macro.sequences.Length
    g := ClickSeqData_SeqGroupById(macro, groupId)
    if (IsObject(g)) {
        if (!g.HasProp("seqIndexes") || !IsObject(g.seqIndexes))
            g.seqIndexes := []
        g.seqIndexes.Push(idx)
    }
    return idx
}

ClickSeqData_RemoveSequenceAt(macro, seqIndex) {
    if (!IsObject(macro) || !macro.HasProp("sequences") || seqIndex < 1 || seqIndex > macro.sequences.Length)
        return false
    macro.sequences.RemoveAt(seqIndex)
    idx := 1
    for seq in macro.sequences {
        seq.order := idx
        idx += 1
    }
    if (!macro.HasProp("seqGroups") || !IsObject(macro.seqGroups))
        return true
    for g in macro.seqGroups {
        if (!g.HasProp("seqIndexes") || !IsObject(g.seqIndexes))
            continue
        next := []
        for n in g.seqIndexes {
            if (n = seqIndex)
                continue
            next.Push(n > seqIndex ? n - 1 : n)
        }
        g.seqIndexes := next
    }
    return true
}

ClickSeqData_NewHardcodedSlot(scriptId) {
    return { type: "hardcoded", scriptId: ClickSeqData_SanitizeId(scriptId) }
}

ClickSeqData_NewSeqGroupSlot(groupId) {
    return { type: "seqGroup", groupId: ClickSeqData_SanitizeId(groupId) }
}

ClickSeqData_NewSeqGroup(groupId) {
    return { id: ClickSeqData_SanitizeId(groupId), seqIndexes: [] }
}

ClickSeqData_DefaultSeqGroups(seqCount) {
    idxs := []
    loop seqCount
        idxs.Push(A_Index)
    return [{ id: "clicks", seqIndexes: idxs }]
}

ClickSeqData_DefaultMacros() {
    seqs := ClickSeqData_DefaultAiQuickDownloadSequences()
    return [{
        id: CLICKSEQ_DEFAULT_MACRO_ID,
        name: "AI Quick Download",
        trigger: "#!+9",
        postAction: "desktopCutNewest",
        rules: ClickSeqData_DefaultRules(),
        slots: ClickSeqData_DefaultSlots(),
        seqGroups: ClickSeqData_DefaultSeqGroups(seqs.Length),
        sequences: seqs
    }]
}

ClickSeqData_ApplyDefaultSequencesIfEmpty(list) {
    seeded := false
    defSeqs := ""
    for macro in list {
        if (macro.id != CLICKSEQ_DEFAULT_MACRO_ID)
            continue
        n := (macro.HasProp("sequences") && IsObject(macro.sequences)) ? macro.sequences.Length : 0
        if (n > 0)
            continue
        if (!IsObject(defSeqs))
            defSeqs := ClickSeqData_DefaultAiQuickDownloadSequences()
        macro.sequences := defSeqs
        ClickSeqData_EnsureSlots(macro)
        seeded := true
    }
    return seeded
}

ClickSeqData_NewSequence(name := "", context := "enterprise") {
    return {
        name: name,
        context: ClickSeqData_NormalizeContext(context),
        order: 1,
        clicks: []
    }
}

ClickSeqData_NewClick() {
    return {
        preferNewest: true,
        settleMs: CLICKSEQ_DEFAULT_SETTLE_MS,
        controlType: "Button",
        classContains: "",
        selectors: []
    }
}

ClickSeqData_NewSelector(kind := "name", value := "", match := "exact") {
    return {
        kind: ClickSeqData_NormalizeKind(kind),
        match: ClickSeqData_NormalizeMatch(match),
        value: Trim(value)
    }
}

ClickSeqData_NormalizeContext(ctx) {
    c := StrLower(Trim(ctx))
    if (c = "*" || c = "any" || c = "all")
        return "*"
    if (c = "enterprise" || c = "gemini enterprise")
        return "enterprise"
    if (c = "copilot" || c = "copilot web")
        return "copilot"
    if (c = "gemini" || c = "personal")
        return "gemini"
    if (c = "")
        return "enterprise"
    return c
}

ClickSeqData_NormalizeKind(kind) {
    k := StrLower(Trim(kind))
    if (k = "nameid" || k = "automationid" || k = "automation_id" || k = "id")
        return "automationId"
    if (k = "class" || k = "classname" || k = "classcontains")
        return "classContains"
    if (k = "region" || k = "area" || k = "rect")
        return "region"
    if (k = "icon" || k = "image")
        return "icon"
    return "name"
}

ClickSeqData_NormalizeMatch(match) {
    m := StrLower(Trim(match))
    if (m = "substr" || m = "substring" || m = "contains")
        return "substr"
    return "exact"
}

ClickSeqData_KindLabel(kind) {
    switch ClickSeqData_NormalizeKind(kind) {
        case "automationId":
            return "automationId"
        case "classContains":
            return "classContains"
        case "region":
            return "region"
        case "icon":
            return "icon"
        default:
            return "name"
    }
}

ClickSeqData_ContextLabel(ctx) {
    switch ClickSeqData_NormalizeContext(ctx) {
        case "enterprise":
            return "Gemini Enterprise"
        case "copilot":
            return "Copilot"
        case "gemini":
            return "Gemini"
        case "*":
            return "Any companion"
        default:
            return ctx
    }
}

ClickSeqData_EncodeSelector(sel) {
    if (!IsObject(sel))
        return ""
    kind := ClickSeqData_NormalizeKind(sel.HasProp("kind") ? sel.kind : "name")
    match := ClickSeqData_NormalizeMatch(sel.HasProp("match") ? sel.match : "exact")
    value := sel.HasProp("value") ? Trim(sel.value) : ""
    value := StrReplace(value, "|", " ")
    if ((kind = "name" || kind = "automationId") && match = "substr")
        return kind . ":substr:" . value
    return kind . ":" . value
}

ClickSeqData_ParseSelector(token) {
    token := Trim(token)
    if (token = "")
        return ""
    pos := InStr(token, ":")
    if (!pos)
        return ClickSeqData_NewSelector("name", token, "exact")
    kind := ClickSeqData_NormalizeKind(SubStr(token, 1, pos - 1))
    rest := SubStr(token, pos + 1)
    match := "exact"
    if (kind = "name" || kind = "automationId") {
        pos2 := InStr(rest, ":")
        if (pos2) {
            maybeMatch := StrLower(Trim(SubStr(rest, 1, pos2 - 1)))
            if (maybeMatch = "exact" || maybeMatch = "substr" || maybeMatch = "substring") {
                match := ClickSeqData_NormalizeMatch(maybeMatch)
                rest := SubStr(rest, pos2 + 1)
            }
        }
    }
    return ClickSeqData_NewSelector(kind, rest, match)
}

ClickSeqData_JoinSelectors(arr) {
    s := ""
    if (!IsObject(arr))
        return s
    for sel in arr {
        enc := ClickSeqData_EncodeSelector(sel)
        if (enc = "")
            continue
        if (s != "")
            s .= "|"
        s .= enc
    }
    return s
}

ClickSeqData_ParseSelectors(raw) {
    out := []
    text := Trim(ClickSeqData_NormalizeIniValue(raw))
    if (text = "")
        return out
    for part in StrSplit(text, "|") {
        sel := ClickSeqData_ParseSelector(part)
        if (IsObject(sel) && sel.value != "")
            out.Push(sel)
    }
    return out
}

ClickSeqData_SelectorPreview(sel) {
    if (!IsObject(sel))
        return ""
    kind := ClickSeqData_KindLabel(sel.kind)
    match := ClickSeqData_NormalizeMatch(sel.HasProp("match") ? sel.match : "exact")
    extra := (match = "substr") ? " ~" : ""
    return kind . extra . ": " . sel.value
}

ClickSeqData_ClickPreview(click) {
    if (!IsObject(click) || !IsObject(click.selectors) || click.selectors.Length = 0)
        return "(no selectors)"
    parts := []
    for sel in click.selectors {
        v := Trim(sel.value)
        if (v != "")
            parts.Push(v)
        if (parts.Length >= 3)
            break
    }
    s := ""
    for p in parts {
        if (s != "")
            s .= " | "
        s .= p
    }
    return s
}

ClickSeqData_SequencePreview(seq) {
    if (!IsObject(seq) || !IsObject(seq.clicks) || seq.clicks.Length = 0)
        return "(no clicks)"
    parts := []
    for click in seq.clicks
        parts.Push(ClickSeqData_ClickPreview(click))
    s := ""
    for p in parts {
        if (s != "")
            s .= " → "
        s .= p
    }
    return s
}

ClickSeqData_Int(val, fallback := 0) {
    n := fallback
    try n := Integer(val)
    catch {
        n := fallback
    }
    return n
}

ClickSeqData_Bool01(val, fallback := true) {
    if (val = "" || val = "ERROR")
        return fallback
    v := StrLower(Trim(val))
    if (v = "0" || v = "false" || v = "no")
        return false
    return true
}

ClickSeqData_ReadClick(path, section) {
    prefer := ""
    settle := ""
    ctype := ""
    classC := ""
    sels := ""
    try prefer := IniRead(path, section, "PreferNewest", "1")
    try settle := IniRead(path, section, "SettleMs", CLICKSEQ_DEFAULT_SETTLE_MS)
    try ctype := IniRead(path, section, "ControlType", "Button")
    try classC := IniRead(path, section, "ClassContains", "")
    try sels := IniRead(path, section, "Selectors", "")
    ctype := ClickSeqData_NormalizeIniValue(ctype)
    if (ctype = "")
        ctype := "Button"
    return {
        preferNewest: ClickSeqData_Bool01(prefer, true),
        settleMs: ClickSeqData_Int(settle, CLICKSEQ_DEFAULT_SETTLE_MS),
        controlType: ctype,
        classContains: ClickSeqData_NormalizeIniValue(classC),
        selectors: ClickSeqData_ParseSelectors(sels)
    }
}

ClickSeqData_ReadSequence(path, macroId, seqIdx) {
    section := "Seq_" . macroId . "_" . seqIdx
    name := ""
    try name := IniRead(path, section, "Name", "")
    catch {
        return ""
    }
    name := ClickSeqData_NormalizeIniValue(name)
    ctx := ""
    order := ""
    try ctx := IniRead(path, section, "Context", "enterprise")
    try order := IniRead(path, section, "Order", seqIdx)
    clicks := []
    clickIdx := 1
    loop CLICKSEQ_MAX_CLICK {
        csec := "Click_" . macroId . "_" . seqIdx . "_" . clickIdx
        has := ""
        try has := IniRead(path, csec, "Selectors", "")
        catch {
            break
        }
        preferKey := ""
        try preferKey := IniRead(path, csec, "PreferNewest", "")
        catch {
            preferKey := ""
        }
        if (ClickSeqData_NormalizeIniValue(has) = "" && ClickSeqData_NormalizeIniValue(preferKey) = "")
            break
        clicks.Push(ClickSeqData_ReadClick(path, csec))
        clickIdx += 1
    }
    return {
        name: name,
        context: ClickSeqData_NormalizeContext(ctx),
        order: ClickSeqData_Int(order, seqIdx),
        clicks: clicks
    }
}

ClickSeqData_EnsureDefaultMacro(list) {
    found := false
    for m in list {
        if (m.id = CLICKSEQ_DEFAULT_MACRO_ID) {
            found := true
            break
        }
    }
    if (!found) {
        def := ClickSeqData_DefaultMacros()
        list.InsertAt(1, def[1])
    }
    return list
}

ClickSeqData_Load(force := false, skipMtime := false) {
    global g_ClickSeqMacros, g_ClickSeqDataCacheReady, g_ClickSeqDataCacheMtime
    if (!force && skipMtime && g_ClickSeqDataCacheReady)
        return g_ClickSeqMacros
    path := ClickSeqData_IniPath()
    if (!FileExist(path)) {
        if (!ClickSeqData_Save(ClickSeqData_DefaultMacros()))
            return g_ClickSeqMacros
        force := true
    }
    mtime := ClickSeqData_FileMtime()
    if (!force && g_ClickSeqDataCacheReady && mtime = g_ClickSeqDataCacheMtime)
        return g_ClickSeqMacros

    list := []
    rawIds := ""
    try rawIds := IniRead(path, "Index", "Macros", "")
    catch {
        rawIds := ""
    }
    rawIds := ClickSeqData_NormalizeIniValue(rawIds)
    ids := []
    if (rawIds != "") {
        for part in StrSplit(rawIds, "|") {
            id := ClickSeqData_SanitizeId(part)
            if (id != "")
                ids.Push(id)
        }
    }
    if (ids.Length = 0)
        ids.Push(CLICKSEQ_DEFAULT_MACRO_ID)

    needMigrate := false
    seen := Map()
    for id in ids {
        if (seen.Has(id))
            continue
        seen[id] := true
        section := "Macro_" . id
        name := ""
        try name := IniRead(path, section, "Name", "")
        catch {
            name := ""
        }
        name := ClickSeqData_NormalizeIniValue(name)
        trigger := ""
        postAction := ""
        try trigger := IniRead(path, section, "Trigger", "")
        try postAction := IniRead(path, section, "PostAction", "")
        trigger := ClickSeqData_NormalizeIniValue(trigger)
        postAction := ClickSeqData_NormalizeIniValue(postAction)
        rulesRaw := ""
        slotsRaw := ""
        try rulesRaw := IniRead(path, section, "Rules", "")
        try slotsRaw := IniRead(path, section, "Slots", "")
        rulesRaw := ClickSeqData_NormalizeIniValue(rulesRaw)
        slotsRaw := ClickSeqData_NormalizeIniValue(slotsRaw)
        if (name = "" && id = CLICKSEQ_DEFAULT_MACRO_ID) {
            name := "AI Quick Download"
            if (trigger = "")
                trigger := "#!+9"
            if (postAction = "")
                postAction := "desktopCutNewest"
        }
        sequences := []
        seqIdx := 1
        loop CLICKSEQ_MAX_SEQ {
            seq := ClickSeqData_ReadSequence(path, id, seqIdx)
            if (!IsObject(seq))
                break
            if (seq.name = "" && seq.clicks.Length = 0)
                break
            sequences.Push(seq)
            seqIdx += 1
        }
        ClickSeqData_SortSequences(sequences)
        seqGroups := []
        gIdx := 1
        loop 20 {
            gsec := "SeqGroup_" . id . "_" . gIdx
            gid := ""
            try gid := IniRead(path, gsec, "Id", "")
            catch {
                break
            }
            gid := ClickSeqData_SanitizeId(ClickSeqData_NormalizeIniValue(gid))
            seqLine := ""
            try seqLine := IniRead(path, gsec, "Sequences", "")
            if (gid = "" && ClickSeqData_NormalizeIniValue(seqLine) = "")
                break
            if (gid = "")
                gid := "clicks"
            idxs := []
            for tok in StrSplit(ClickSeqData_NormalizeIniValue(seqLine), "|") {
                n := ClickSeqData_Int(Trim(tok), 0)
                if (n > 0)
                    idxs.Push(n)
            }
            seqGroups.Push({ id: gid, seqIndexes: idxs })
            gIdx += 1
        }
        macroObj := {
            id: id,
            name: name != "" ? name : id,
            trigger: trigger,
            postAction: postAction,
            rules: ClickSeqData_ParseRules(rulesRaw),
            slots: ClickSeqData_ParseSlots(slotsRaw),
            seqGroups: seqGroups,
            sequences: sequences
        }
        if (slotsRaw = "")
            needMigrate := true
        ClickSeqData_EnsureSlots(macroObj)
        list.Push(macroObj)
    }

    list := ClickSeqData_EnsureDefaultMacro(list)
    if (ClickSeqData_ApplyDefaultSequencesIfEmpty(list) || needMigrate) {
        ClickSeqData_Save(list)
        return g_ClickSeqMacros
    }
    g_ClickSeqMacros := list
    g_ClickSeqDataCacheReady := true
    g_ClickSeqDataCacheMtime := mtime
    return g_ClickSeqMacros
}

ClickSeqData_SortSequences(sequences) {
    if (!IsObject(sequences) || sequences.Length < 2)
        return
    i := 1
    while (i < sequences.Length) {
        j := i + 1
        while (j <= sequences.Length) {
            oi := ClickSeqData_Int(sequences[i].order, i)
            oj := ClickSeqData_Int(sequences[j].order, j)
            if (oj < oi) {
                tmp := sequences[i]
                sequences[i] := sequences[j]
                sequences[j] := tmp
            }
            j += 1
        }
        i += 1
    }
    idx := 1
    for seq in sequences {
        seq.order := idx
        idx += 1
    }
}

ClickSeqData_Save(list) {
    global g_ClickSeqMacros, g_ClickSeqDataCacheReady, g_ClickSeqDataCacheMtime
    path := ClickSeqData_IniPath()
    try DirCreate(A_ScriptDir "\assets\data")
    catch {
    }
    try FileDelete(path)
    catch {
    }
    if (!IsObject(list))
        list := []
    list := ClickSeqData_EnsureDefaultMacro(list)
    try {
        idLine := ""
        for macro in list {
            if (idLine != "")
                idLine .= "|"
            idLine .= macro.id
        }
        IniWrite(idLine, path, "Index", "Macros")
        for macro in list {
            msec := "Macro_" . macro.id
            IniWrite(macro.HasProp("name") ? macro.name : macro.id, path, msec, "Name")
            IniWrite(macro.HasProp("trigger") ? macro.trigger : "", path, msec, "Trigger")
            IniWrite(macro.HasProp("postAction") ? macro.postAction : "", path, msec, "PostAction")
            ClickSeqData_EnsureSlots(macro)
            IniWrite(ClickSeqData_EncodeRules(macro.rules), path, msec, "Rules")
            IniWrite(ClickSeqData_EncodeSlots(macro.slots), path, msec, "Slots")
            gIdx := 1
            groups := macro.HasProp("seqGroups") ? macro.seqGroups : []
            for grp in groups {
                gsec := "SeqGroup_" . macro.id . "_" . gIdx
                IniWrite(grp.HasProp("id") ? grp.id : "clicks", path, gsec, "Id")
                idxLine := ""
                if (grp.HasProp("seqIndexes") && IsObject(grp.seqIndexes)) {
                    for n in grp.seqIndexes {
                        if (idxLine != "")
                            idxLine .= "|"
                        idxLine .= n
                    }
                }
                IniWrite(idxLine, path, gsec, "Sequences")
                gIdx += 1
            }
            seqs := macro.HasProp("sequences") ? macro.sequences : []
            ClickSeqData_SortSequences(seqs)
            seqIdx := 1
            for seq in seqs {
                ssec := "Seq_" . macro.id . "_" . seqIdx
                IniWrite(seq.HasProp("name") ? seq.name : "", path, ssec, "Name")
                IniWrite(seq.HasProp("context") ? seq.context : "enterprise", path, ssec, "Context")
                IniWrite(seqIdx, path, ssec, "Order")
                clickIdx := 1
                clicks := seq.HasProp("clicks") ? seq.clicks : []
                for click in clicks {
                    csec := "Click_" . macro.id . "_" . seqIdx . "_" . clickIdx
                    prefer := (click.HasProp("preferNewest") && !click.preferNewest) ? "0" : "1"
                    settle := click.HasProp("settleMs") ? click.settleMs : CLICKSEQ_DEFAULT_SETTLE_MS
                    ctype := click.HasProp("controlType") && click.controlType != "" ? click.controlType : "Button"
                    classC := click.HasProp("classContains") ? click.classContains : ""
                    IniWrite(prefer, path, csec, "PreferNewest")
                    IniWrite(settle, path, csec, "SettleMs")
                    IniWrite(ctype, path, csec, "ControlType")
                    IniWrite(classC, path, csec, "ClassContains")
                    IniWrite(ClickSeqData_JoinSelectors(click.HasProp("selectors") ? click.selectors : []), path, csec,
                    "Selectors")
                    clickIdx += 1
                }
                seqIdx += 1
            }
        }
    } catch {
        ClickSeqData_Invalidate()
        return false
    }
    g_ClickSeqMacros := list
    g_ClickSeqDataCacheReady := true
    g_ClickSeqDataCacheMtime := ClickSeqData_FileMtime()
    return true
}

ClickSeqData_MacroById(macroId) {
    id := ClickSeqData_SanitizeId(macroId)
    for macro in ClickSeqData_Load() {
        if (macro.id = id)
            return macro
    }
    return ""
}

ClickSeqData_SequenceCount() {
    n := 0
    for macro in ClickSeqData_Load(false, true) {
        if (macro.HasProp("sequences") && IsObject(macro.sequences))
            n += macro.sequences.Length
    }
    return n
}

ClickSeqData_CloneSelector(sel) {
    return ClickSeqData_NewSelector(
        sel.HasProp("kind") ? sel.kind : "name",
        sel.HasProp("value") ? sel.value : "",
        sel.HasProp("match") ? sel.match : "exact")
}

ClickSeqData_CloneClick(click) {
    out := ClickSeqData_NewClick()
    if (!IsObject(click))
        return out
    out.preferNewest := click.HasProp("preferNewest") ? !!click.preferNewest : true
    settleVal := click.HasProp("settleMs") ? click.settleMs : CLICKSEQ_DEFAULT_SETTLE_MS
    out.settleMs := ClickSeqData_Int(settleVal, CLICKSEQ_DEFAULT_SETTLE_MS)
    out.controlType := click.HasProp("controlType") && click.controlType != "" ? click.controlType : "Button"
    out.classContains := click.HasProp("classContains") ? click.classContains : ""
    out.selectors := []
    if (click.HasProp("selectors") && IsObject(click.selectors)) {
        for sel in click.selectors
            out.selectors.Push(ClickSeqData_CloneSelector(sel))
    }
    return out
}

ClickSeqData_CloneSequence(seq) {
    out := ClickSeqData_NewSequence()
    if (!IsObject(seq))
        return out
    out.name := seq.HasProp("name") ? seq.name : ""
    out.context := ClickSeqData_NormalizeContext(seq.HasProp("context") ? seq.context : "enterprise")
    out.order := seq.HasProp("order") ? ClickSeqData_Int(seq.order, 1) : 1
    out.clicks := []
    if (seq.HasProp("clicks") && IsObject(seq.clicks)) {
        for click in seq.clicks
            out.clicks.Push(ClickSeqData_CloneClick(click))
    }
    return out
}

ClickSeqData_ReplaceMacro(updated) {
    list := ClickSeqData_Load()
    out := []
    found := false
    for macro in list {
        if (macro.id = updated.id) {
            out.Push(updated)
            found := true
        } else {
            out.Push(macro)
        }
    }
    if (!found)
        out.Push(updated)
    return ClickSeqData_Save(out)
}

ClickSeqData_Swap(arr, i, j) {
    if (i < 1 || j < 1 || i > arr.Length || j > arr.Length)
        return false
    tmp := arr[i]
    arr[i] := arr[j]
    arr[j] := tmp
    return true
}
