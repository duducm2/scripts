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

ClickSeqData_DefaultMacros() {
    return [{
        id: CLICKSEQ_DEFAULT_MACRO_ID,
        name: "AI Quick Download",
        trigger: "#!+9",
        postAction: "desktopCutNewest",
        sequences: []
    }]
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
        list.Push({
            id: id,
            name: name != "" ? name : id,
            trigger: trigger,
            postAction: postAction,
            sequences: sequences
        })
    }

    list := ClickSeqData_EnsureDefaultMacro(list)
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
