; =============================================================================
; Utils module: click_sequence_runner.ahk
; Execute configured click sequences against a window via UIA-v2.
; Slot chain runner: Hardcoded Scripts + Sequence Groups (sibling fallbacks).
; =============================================================================

global g_ClickSeqSearchOrder := "bottomUp"

ClickSeq_SetFail(message) {
    global g_ClickSeqLastFailMessage
    g_ClickSeqLastFailMessage := message
}

ClickSeq_LastFailMessage() {
    global g_ClickSeqLastFailMessage
    return g_ClickSeqLastFailMessage
}

ClickSeq_CompanionLabel(companion) {
    return ClickSeqData_ContextLabel(companion)
}

ClickSeq_TypeSpec(controlType) {
    t := Trim(controlType)
    if (t = "" || t = "Button" || t = "50000")
        return 50000
    if (t = "MenuItem" || t = "50011")
        return 50011
    if (t = "Hyperlink" || t = "50005")
        return 50005
    if (t = "ListItem" || t = "50007")
        return 50007
    n := 0
    try n := Integer(t)
    catch {
        n := 0
    }
    if (n > 0)
        return n
    return t
}

ClickSeq_ClassContains(el, needle) {
    if (!IsObject(el) || needle = "")
        return false
    try {
        cn := el.ClassName
        return (cn != "" && InStr(cn, needle))
    } catch {
        return false
    }
}

ClickSeq_ElName(el) {
    if (!IsObject(el))
        return ""
    n := ""
    try n := el.Name
    catch {
        return ""
    }
    return n
}

ClickSeq_ElAutomationId(el) {
    if (!IsObject(el))
        return ""
    id := ""
    try id := el.AutomationId
    catch {
        return ""
    }
    return id
}

ClickSeq_SelectorMatches(el, click, sel) {
    if (!IsObject(el) || !IsObject(sel))
        return false
    classAnd := click.HasProp("classContains") ? Trim(click.classContains) : ""
    if (classAnd != "" && !ClickSeq_ClassContains(el, classAnd))
        return false
    kind := ClickSeqData_NormalizeKind(sel.HasProp("kind") ? sel.kind : "name")
    match := ClickSeqData_NormalizeMatch(sel.HasProp("match") ? sel.match : "exact")
    want := sel.HasProp("value") ? Trim(sel.value) : ""
    if (want = "")
        return false
    switch kind {
        case "name":
            n := ClickSeq_ElName(el)
            if (n = "")
                return false
            if (match = "substr")
                return InStr(n, want)
            return (n = want)
        case "automationId":
            id := ClickSeq_ElAutomationId(el)
            if (id = "")
                return false
            if (match = "substr")
                return InStr(id, want)
            return (id = want)
        case "classContains":
            return ClickSeq_ClassContains(el, want)
        default:
            return false
    }
}

ClickSeq_ElementTop(el) {
    if (!IsObject(el))
        return ""
    try {
        br := el.BoundingRectangle
        if (!IsObject(br))
            return ""
        if ((br.r - br.l) <= 0 || (br.b - br.t) <= 0)
            return ""
        return br.t
    } catch {
        return ""
    }
}

ClickSeq_PickNewest(matches) {
    if (!IsObject(matches) || matches.Length = 0)
        return 0
    lastEl := 0
    lastTop := ""
    for el in matches {
        top := ClickSeq_ElementTop(el)
        if (top = "")
            continue
        if (lastEl = 0 || top >= lastTop) {
            lastEl := el
            lastTop := top
        }
    }
    return lastEl ? lastEl : matches[matches.Length]
}

ClickSeq_PickTopmost(matches) {
    if (!IsObject(matches) || matches.Length = 0)
        return 0
    firstEl := 0
    firstTop := ""
    for el in matches {
        top := ClickSeq_ElementTop(el)
        if (top = "")
            continue
        if (firstEl = 0 || top < firstTop) {
            firstEl := el
            firstTop := top
        }
    }
    return firstEl ? firstEl : matches[1]
}

ClickSeq_SearchOrder() {
    global g_ClickSeqSearchOrder
    if (!IsSet(g_ClickSeqSearchOrder) || g_ClickSeqSearchOrder = "")
        return "bottomUp"
    return g_ClickSeqSearchOrder
}

ClickSeq_FindAllOfType(uia, typeSpec) {
    if (!IsObject(uia))
        return []
    list := []
    try {
        found := uia.FindAll({ Type: typeSpec })
        if (found) {
            for el in found
                list.Push(el)
        }
    } catch {
        if (typeSpec = 50000) {
            try {
                found := uia.FindAll({ Type: "Button" })
                if (found) {
                    for el in found
                        list.Push(el)
                }
            } catch {
            }
        }
    }
    return list
}

ClickSeq_FindElement(uia, click, sel) {
    if (!IsObject(uia) || !IsObject(sel))
        return 0
    kind := ClickSeqData_NormalizeKind(sel.HasProp("kind") ? sel.kind : "name")
    if (kind = "icon" || kind = "region")
        return 0
    typeSpec := ClickSeq_TypeSpec(click.HasProp("controlType") ? click.controlType : "Button")
    preferNewest := !(click.HasProp("preferNewest") && !click.preferNewest)
    searchOrder := ClickSeq_SearchOrder()
    classAnd := click.HasProp("classContains") ? Trim(click.classContains) : ""
    match := ClickSeqData_NormalizeMatch(sel.HasProp("match") ? sel.match : "exact")
    want := sel.HasProp("value") ? Trim(sel.value) : ""

    if (!preferNewest && kind = "name" && match = "exact" && want != "" && classAnd = "" && searchOrder = "firstMatch") {
        el := 0
        try el := uia.FindFirst({ Type: typeSpec, Name: want })
        catch {
            el := 0
        }
        if (IsObject(el) && ClickSeq_SelectorMatches(el, click, sel))
            return el
    }

    matches := []
    for el in ClickSeq_FindAllOfType(uia, typeSpec) {
        try {
            if (ClickSeq_SelectorMatches(el, click, sel))
                matches.Push(el)
        } catch {
        }
    }
    if (matches.Length = 0)
        return 0
    if (!preferNewest || searchOrder = "firstMatch")
        return matches[1]
    if (searchOrder = "topDown")
        return ClickSeq_PickTopmost(matches)
    return ClickSeq_PickNewest(matches)
}

ClickSeq_Invoke(el) {
    if (!IsObject(el))
        return false
    try {
        if el.GetPropertyValue(UIA.Property.IsInvokePatternAvailable) {
            el.InvokePattern.Invoke()
            return true
        }
    } catch {
    }
    try {
        el.Invoke()
        return true
    } catch {
    }
    try {
        el.Click()
        return true
    } catch {
    }
    return false
}

ClickSeq_ParseRegion(value) {
    parts := StrSplit(value, ",")
    if (parts.Length < 2)
        return ""
    x := ClickSeqData_Int(Trim(parts[1]), -1)
    y := ClickSeqData_Int(Trim(parts[2]), -1)
    w := parts.Length >= 3 ? ClickSeqData_Int(Trim(parts[3]), 1) : 1
    h := parts.Length >= 4 ? ClickSeqData_Int(Trim(parts[4]), 1) : 1
    if (x < 0 || y < 0)
        return ""
    if (w < 1)
        w := 1
    if (h < 1)
        h := 1
    return { x: x, y: y, w: w, h: h }
}

ClickSeq_ClickRegion(hwnd, value) {
    if (!hwnd)
        return false
    rect := ClickSeq_ParseRegion(value)
    if (!IsObject(rect))
        return false
    wx := 0
    wy := 0
    try WinGetPos(&wx, &wy, , , "ahk_id " hwnd)
    catch {
        return false
    }
    cx := wx + rect.x + rect.w // 2
    cy := wy + rect.y + rect.h // 2
    prevMode := A_CoordModeMouse
    CoordMode("Mouse", "Screen")
    try {
        Click(cx . " " . cy)
    } catch {
        CoordMode("Mouse", prevMode)
        return false
    }
    CoordMode("Mouse", prevMode)
    return true
}

ClickSeq_AttachUia(hwnd) {
    if (!hwnd)
        return 0
    uia := 0
    try uia := UIA_Browser("ahk_id " hwnd)
    catch {
        uia := 0
    }
    return IsObject(uia) ? uia : 0
}

ClickSeq_ExecuteClick(uia, hwnd, click) {
    if (!IsObject(click) || !click.HasProp("selectors") || !IsObject(click.selectors))
        return false
    for sel in click.selectors {
        kind := ClickSeqData_NormalizeKind(sel.HasProp("kind") ? sel.kind : "name")
        if (kind = "icon")
            continue
        if (kind = "region") {
            if (ClickSeq_ClickRegion(hwnd, sel.HasProp("value") ? sel.value : ""))
                return true
            continue
        }
        el := ClickSeq_FindElement(uia, click, sel)
        if (IsObject(el) && ClickSeq_Invoke(el))
            return true
    }
    return false
}

ClickSeq_SequencesForCompanion(macro, companion) {
    out := []
    if (!IsObject(macro) || !macro.HasProp("sequences") || !IsObject(macro.sequences))
        return out
    ctx := ClickSeqData_NormalizeContext(companion)
    specific := []
    wildcard := []
    for seq in macro.sequences {
        sc := ClickSeqData_NormalizeContext(seq.HasProp("context") ? seq.context : "")
        if (sc = ctx)
            specific.Push(seq)
        else if (sc = "*")
            wildcard.Push(seq)
    }
    ClickSeqData_SortSequences(specific)
    ClickSeqData_SortSequences(wildcard)
    for seq in specific
        out.Push(seq)
    for seq in wildcard
        out.Push(seq)
    return out
}

; Returns true if every Slot in the Shortcut chain succeeds.
; extras: { doCut, desktopPath, beforePath, beforeStamp, seqAttempts }
ClickSeq_TrySiblingSequences(seqs, hwnd) {
    for seq in seqs {
        clicks := seq.HasProp("clicks") ? seq.clicks : []
        if (!IsObject(clicks) || clicks.Length = 0)
            continue
        uia := ClickSeq_AttachUia(hwnd)
        if (!IsObject(uia))
            continue
        ok := true
        for click in clicks {
            if (!ClickSeq_ExecuteClick(uia, hwnd, click)) {
                ok := false
                break
            }
            settleVal := click.HasProp("settleMs") ? click.settleMs : CLICKSEQ_DEFAULT_SETTLE_MS
            settle := ClickSeqData_Int(settleVal, CLICKSEQ_DEFAULT_SETTLE_MS)
            if (settle < 0)
                settle := 0
            if (settle > 0)
                Sleep settle
            uia := ClickSeq_AttachUia(hwnd)
            if (!IsObject(uia)) {
                ok := false
                break
            }
        }
        if (ok)
            return true
    }
    return false
}

ClickSeq_RunSeqGroup(macro, groupId, companion, hwnd) {
    rows := ClickSeqData_SequencesForGroup(macro, groupId)
    ctx := ClickSeqData_NormalizeContext(companion)
    specific := []
    wildcard := []
    for row in rows {
        seq := row.seq
        sc := ClickSeqData_NormalizeContext(seq.HasProp("context") ? seq.context : "")
        if (sc = ctx)
            specific.Push(seq)
        else if (sc = "*")
            wildcard.Push(seq)
    }
    ClickSeqData_SortSequences(specific)
    ClickSeqData_SortSequences(wildcard)
    seqs := []
    for seq in specific
        seqs.Push(seq)
    for seq in wildcard
        seqs.Push(seq)
    if (seqs.Length = 0) {
        ClickSeq_SetFail("❌ No Sibling Sequences configured for " . ClickSeq_CompanionLabel(companion)
            . ". Open Utility Shortcuts → Sequences.")
        return false
    }
    if (ClickSeq_TrySiblingSequences(seqs, hwnd))
        return true
    ClickSeq_SetFail("❌ Quick Download: click sequence failed for " . ClickSeq_CompanionLabel(companion))
    return false
}

ClickSeq_RunMacro(macroId, companion, hwnd, extras := unset) {
    global g_ClickSeqSearchOrder, g_ClickSeqRunCtx
    ClickSeq_SetFail("")
    label := ClickSeq_CompanionLabel(companion)
    macro := ClickSeqData_MacroById(macroId)
    if (!IsObject(macro)) {
        ClickSeq_SetFail("❌ No click sequences configured for " . label . ". Open Utility Shortcuts → Sequences.")
        return false
    }
    ClickSeqData_EnsureSlots(macro)
    if (!hwnd) {
        ClickSeq_SetFail("❌ Click sequence: target window missing for " . label . ".")
        return false
    }

    doCut := true
    seqAttempts := 1
    desktopPath := ""
    beforePath := ""
    beforeStamp := ""
    if (IsSet(extras) && IsObject(extras)) {
        if (extras.HasProp("doCut"))
            doCut := !!extras.doCut
        if (extras.HasProp("seqAttempts"))
            seqAttempts := ClickSeqData_Int(extras.seqAttempts, 1)
        if (extras.HasProp("desktopPath"))
            desktopPath := extras.desktopPath
        if (extras.HasProp("beforePath"))
            beforePath := extras.beforePath
        if (extras.HasProp("beforeStamp"))
            beforeStamp := extras.beforeStamp
    }
    if (seqAttempts < 1)
        seqAttempts := 1

    g_ClickSeqSearchOrder := "bottomUp"
    if (macro.HasProp("rules") && IsObject(macro.rules) && macro.rules.HasProp("searchOrder"))
        g_ClickSeqSearchOrder := ClickSeqData_NormalizeSearchOrder(macro.rules.searchOrder)

    g_ClickSeqRunCtx := {
        hwnd: hwnd,
        companion: companion,
        desktopPath: desktopPath,
        beforePath: beforePath,
        beforeStamp: beforeStamp,
        lastPath: "",
        doCut: doCut
    }

    for slot in macro.slots {
        if (slot.HasProp("type") && slot.type = "hardcoded") {
            sid := slot.HasProp("scriptId") ? slot.scriptId : ""
            if (sid = "desktopCut" && !doCut)
                continue
            if (!ClickSeqScript_Run(sid))
                return false
            continue
        }
        groupId := slot.HasProp("groupId") ? slot.groupId : "clicks"
        ok := false
        loop seqAttempts {
            try StandardLoadingBar_Update("⏳ Finding download control…", BANNER_ACCENT_INTERMEDIATE)
            catch {
            }
            if (ClickSeq_RunSeqGroup(macro, groupId, companion, hwnd)) {
                ok := true
                break
            }
            failMsg := ClickSeq_LastFailMessage()
            if (InStr(failMsg, "No Sibling Sequences") || InStr(failMsg, "No click sequences"))
                break
            if (A_Index < seqAttempts)
                Sleep 1200
        }
        if (!ok)
            return false
    }
    return true
}
