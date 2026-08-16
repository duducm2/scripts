; =============================================================================
; lib/MobillsUia.ahk
; Hotkey-free UIA helpers for Mobills web (Utils process). Prefixed MobillsAuto_
; so they never collide with Mobills_* functions in the Shift keys process.
; =============================================================================

global MOBILLS_STEP_MS := 350
global MOBILLS_TYPE_DELAY_MS := 40
global MOBILLS_GATE_TIMEOUT_MS := 8000
global MOBILLS_CARD_NAME := "Mercado Pago"

MobillsAuto_FindHwnd() {
    for exe in ["ahk_exe chrome.exe", "ahk_exe msedge.exe"] {
        try {
            hwnds := WinGetList(exe)
            for hwnd in hwnds {
                try {
                    t := WinGetTitle("ahk_id " hwnd)
                    if (t != "" && InStr(t, "Mobills"))
                        return hwnd
                } catch {
                }
            }
        } catch {
        }
    }
    return 0
}

MobillsAuto_ActivateHwnd(hwnd) {
    if (!hwnd)
        return false
    try {
        WinActivate("ahk_id " hwnd)
        return !!WinWaitActive("ahk_id " hwnd, , 2)
    } catch {
        return false
    }
}

MobillsAuto_AttachBrowser() {
    hwnd := MobillsAuto_FindHwnd()
    if (!hwnd) {
        try Run("https://web.mobills.com.br/dashboard")
        catch {
            return { uia: "", hwnd: 0, error: "Could not launch Mobills URL." }
        }
        endTick := A_TickCount + 20000
        while (A_TickCount < endTick) {
            hwnd := MobillsAuto_FindHwnd()
            if (hwnd)
                break
            Sleep 400
        }
    }
    if (!hwnd)
        return { uia: "", hwnd: 0, error: "Mobills window not found (Chrome/Edge title)." }
    MobillsAuto_ActivateHwnd(hwnd)
    Sleep MOBILLS_STEP_MS
    uia := ""
    try uia := UIA_Browser("ahk_id " hwnd)
    catch {
        uia := ""
    }
    if (!uia) {
        MobillsAuto_ActivateHwnd(hwnd)
        Sleep 400
        try uia := UIA_Browser("ahk_id " hwnd)
        catch {
            uia := ""
        }
    }
    if (!uia)
        return { uia: "", hwnd: hwnd, error: "UIA_Browser attach failed for Mobills hwnd." }
    return { uia: uia, hwnd: hwnd, error: "" }
}

; Approach A AutomationId, B Name+Type, C ClassName substring, D first matching Type.
; candidates: array of objects { AutomationId?, Name?, Type?, ClassName?, matchmode? }
MobillsAuto_Resolve(scope, candidates, attempted := unset) {
    if (!IsSet(attempted))
        attempted := []
    if !scope
        return ""
    if (!IsObject(candidates))
        return ""
    for , c in candidates {
        label := MobillsAuto_CandidateLabel(c)
        attempted.Push(label)
        el := MobillsAuto_TryOne(scope, c)
        if el
            return el
    }
    return ""
}

MobillsAuto_CandidateLabel(c) {
    parts := []
    if (c.HasProp("AutomationId") && c.AutomationId != "")
        parts.Push("AutomationId=" c.AutomationId)
    if (c.HasProp("Name") && c.Name != "")
        parts.Push("Name=" c.Name)
    if (c.HasProp("Type") && c.Type != "")
        parts.Push("Type=" c.Type)
    if (c.HasProp("ClassName") && c.ClassName != "")
        parts.Push("ClassName~=" c.ClassName)
    if (!parts.Length)
        return "(empty)"
    out := parts[1]
    i := 2
    while (i <= parts.Length) {
        out .= ", " parts[i]
        i++
    }
    return out
}

MobillsAuto_TryOne(scope, c) {
    try {
        el := ""
        try el := scope.FindElement(c)
        catch {
        }
        if (el && c.HasProp("ClassName") && c.ClassName != "") {
            cls := ""
            try cls := el.ClassName
            if !InStr(cls, c.ClassName)
                el := ""
        }
        if el
            return el
    } catch {
    }
    if (c.HasProp("ClassName") && c.ClassName != "" && c.HasProp("Type")) {
        try {
            all := scope.FindAll({ Type: c.Type })
            if all {
                for item in all {
                    try {
                        if InStr(item.ClassName, c.ClassName)
                            return item
                    } catch {
                    }
                }
            }
        } catch {
        }
    }
    return ""
}

MobillsAuto_IsDisabled(el) {
    if !el
        return true
    try {
        if !el.GetPropertyValue(UIA.Property.IsEnabled)
            return true
        if el.GetPropertyValue(UIA.Property.IsOffscreen)
            return true
        cls := ""
        try cls := el.GetPropertyValue(UIA.Property.ClassName)
        if (cls = "") {
            try cls := el.ClassName
        }
        if (cls != "" && InStr(cls, "Mui-disabled"))
            return true
    } catch {
    }
    return false
}

MobillsAuto_WaitFor(fn, timeoutMs := unset, pollMs := 150) {
    if (!IsSet(timeoutMs))
        timeoutMs := MOBILLS_GATE_TIMEOUT_MS
    endTick := A_TickCount + timeoutMs
    while (A_TickCount < endTick) {
        try {
            result := fn.Call()
            if (result)
                return result
        } catch {
        }
        Sleep pollMs
    }
    return ""
}

MobillsAuto_Click(el) {
    if !el
        return false
    try {
        try {
            if (el.GetPropertyValue(UIA.Property.IsInvokePatternAvailable)) {
                el.Invoke()
                return true
            }
        } catch {
        }
        el.Click()
        return true
    } catch {
        try {
            el.Click()
            return true
        } catch {
            return false
        }
    }
}

MobillsAuto_ElementValue(el) {
    if !el
        return ""
    try {
        v := el.Value
        if (v != "")
            return v
    } catch {
    }
    try return el.Name
    catch {
        return ""
    }
}

MobillsAuto_Digits(s) {
    return RegExReplace(s, "[^\d]", "")
}

MobillsAuto_SetEditVerified(el, text) {
    if !el
        return { ok: false, got: "", attempted: ["element missing"] }
    attempted := ["Click", "SetFocus", "^a", "SendText", "read Value"]
    loop 2 {
        try el.Click()
        catch {
        }
        Sleep 80
        try el.SetFocus()
        catch {
        }
        Sleep 80
        Send "^a"
        Sleep MOBILLS_TYPE_DELAY_MS
        prevDelay := A_KeyDelay
        SetKeyDelay MOBILLS_TYPE_DELAY_MS
        try SendText text
        catch {
            Send "{Text}" text
        }
        SetKeyDelay prevDelay
        Sleep MOBILLS_STEP_MS
        got := MobillsAuto_ElementValue(el)
        if (got = text || InStr(got, text))
            return { ok: true, got: got, attempted: attempted }
        if (MobillsAuto_Digits(got) != "" && MobillsAuto_Digits(got) = MobillsAuto_Digits(text))
            return { ok: true, got: got, attempted: attempted }
    }
    return { ok: false, got: MobillsAuto_ElementValue(el), attempted: attempted }
}

MobillsAuto_ComboChipText(combo) {
    if !combo
        return ""
    try {
        texts := combo.FindAll({ Type: 50020 })
        if texts {
            for t in texts {
                try {
                    nm := Trim(t.Name)
                    if (nm != "" && nm != Chr(160) && nm != " ")
                        return nm
                } catch {
                }
            }
        }
    } catch {
    }
    return ""
}

MobillsAuto_ChipMatches(got, wanted) {
    if (got = "" || wanted = "")
        return false
    return (got = wanted || InStr(got, wanted) || InStr(wanted, got))
}

; MUI Autocomplete list is often portaled outside the ComboBox. Prefer Click over Enter.
MobillsAuto_FindOpenListbox(root) {
    if !root
        return ""
    for , cls in ["MuiAutocomplete-listbox", "MuiAutocomplete-popper", "MuiMenu-list"] {
        try {
            el := root.FindElement({ ClassName: cls, matchmode: "Substring" })
            if el
                return el
        } catch {
        }
    }
    for , t in [50008, 50009] {
        try {
            all := root.FindAll({ Type: t })
            if all {
                for item in all {
                    try {
                        if !item.GetPropertyValue(UIA.Property.IsOffscreen)
                            return item
                    } catch {
                    }
                }
            }
        } catch {
        }
    }
    return ""
}

MobillsAuto_FindAutocompleteOption(root, wanted) {
    if !root || wanted = ""
        return ""
    list := MobillsAuto_FindOpenListbox(root)
    scope := list ? list : ""
    if !scope
        return ""
    types := [50007, 50011, 50000, 50020]
    for , t in types {
        try {
            el := scope.FindElement({ Type: t, Name: wanted })
            if el
                return el
        } catch {
        }
    }
    for , t in types {
        try {
            el := scope.FindElement({ Type: t, Name: wanted, matchmode: "Substring" })
            if el {
                nm := ""
                try nm := Trim(el.Name)
                if (nm = wanted)
                    return el
            }
        } catch {
        }
    }
    return ""
}

MobillsAuto_ClickAutocompleteOption(el) {
    if !el
        return false
    target := el
    p := el
    loop 8 {
        ty := 0
        try ty := p.Type
        catch {
            break
        }
        if (ty = 50007 || ty = 50011 || ty = 50000) {
            target := p
            break
        }
        try p := p.WalkTree("p")
        catch {
            break
        }
        if !p
            break
    }
    try target.ScrollIntoView()
    catch {
    }
    ; UIA Click() with no args Invokes; MUI only commits on a physical click.
    try {
        target.Click("left")
        return true
    } catch {
        try {
            target.Click("left")
            return true
        } catch {
            return false
        }
    }
}

MobillsAuto_PickAutocomplete(combo, wanted, root := "") {
    attempted := []
    if !combo
        return { ok: false, got: "", attempted: ["combo missing"] }
    wanted := Trim(wanted)
    if (wanted = "")
        return { ok: false, got: "", attempted: ["wanted empty"] }
    if (root = "")
        root := combo

    chipBefore := MobillsAuto_ComboChipText(combo)
    if MobillsAuto_ChipMatches(chipBefore, wanted) {
        attempted.Push("already selected")
        return { ok: true, got: chipBefore, attempted: attempted }
    }

    openBtn := ""
    try openBtn := combo.FindElement({ Type: 50000, Name: "Open" })
    catch {
    }
    if (!openBtn) {
        try {
            btns := combo.FindAll({ Type: 50000 })
            if (btns && btns.Length)
                openBtn := btns[btns.Length]
        } catch {
        }
    }
    attempted.Push(openBtn ? "Open button" : "Open button missing — click combo")
    if (openBtn)
        MobillsAuto_Click(openBtn)
    else
        MobillsAuto_Click(combo)
    Sleep MOBILLS_STEP_MS

    edit := ""
    try edit := combo.FindElement({ Type: 50004 })
    catch {
    }
    if (edit) {
        attempted.Push("Combo Edit + SendText")
        try edit.Click()
        catch {
        }
        Sleep 80
        try edit.SetFocus()
        catch {
        }
        Send "^a"
        Sleep MOBILLS_TYPE_DELAY_MS
        prevDelay := A_KeyDelay
        SetKeyDelay MOBILLS_TYPE_DELAY_MS
        try SendText wanted
        catch {
            Send "{Text}" wanted
        }
        SetKeyDelay prevDelay
        Sleep 120
        typed := MobillsAuto_ElementValue(edit)
        if (typed = "" || !InStr(typed, wanted)) {
            try edit.Value := wanted
            catch {
            }
            Sleep 80
        }
    } else {
        attempted.Push("no Edit — SendText to focused control")
        SendText wanted
    }
    Sleep 500

    opt := MobillsAuto_FindAutocompleteOption(root, wanted)
    if (!opt)
        opt := MobillsAuto_FindAutocompleteOption(combo, wanted)
    if (opt) {
        attempted.Push("Click list option")
        MobillsAuto_ClickAutocompleteOption(opt)
    } else {
        attempted.Push("list option missing — Enter fallback")
        Send "{Enter}"
    }
    Sleep MOBILLS_STEP_MS

    got := MobillsAuto_ComboChipText(combo)
    if MobillsAuto_ChipMatches(got, wanted)
        return { ok: true, got: got, attempted: attempted }
    return { ok: false, got: got, attempted: attempted }
}

MobillsAuto_FindDialog(uia) {
    if !uia
        return ""
    titles := ["New expense", "New income", "New credit card expense", "New transfer"]
    for , title in titles {
        try {
            t := uia.FindElement({ Type: 50020, Name: title })
            if t {
                try {
                    p := t.WalkTree("p")
                    loop 6 {
                        if !p
                            break
                        try {
                            if (p.Type = 50032 || p.LocalizedType = "caixa de diálogo")
                                return p
                        } catch {
                        }
                        try p := p.WalkTree("p")
                        catch {
                            break
                        }
                    }
                } catch {
                }
                return t
            }
        } catch {
        }
    }
    try {
        dlg := uia.FindElement({ Type: 50032, ClassName: "MuiDialog-paper", matchmode: "Substring" })
        if dlg
            return dlg
    } catch {
    }
    return ""
}

MobillsAuto_DialogTitle(scope) {
    if !scope
        return ""
    titles := ["New expense", "New income", "New credit card expense", "New transfer"]
    for , title in titles {
        try {
            t := scope.FindFirst({ Type: 50020, Name: title })
            if t
                return title
        } catch {
        }
        try {
            t := scope.FindElement({ Type: 50020, Name: title })
            if t
                return title
        } catch {
        }
    }
    return ""
}

MobillsAuto_SelectNewMenuItem(uia, itemName) {
    attempted := []
    actionBtn := MobillsAuto_Resolve(uia, [{ Type: 50000, AutomationId: "action-button" }, { Type: 50000, Name: "New",
        matchmode: "Substring" }], attempted)
    if !actionBtn
        return { ok: false, attempted: attempted }
    clickOk := MobillsAuto_Click(actionBtn)
    if !clickOk
        return { ok: false, attempted: attempted }
    Sleep 280
    menuItem := MobillsAuto_Resolve(uia, [{ Type: 50011, Name: itemName, matchmode: "Substring" }, { Type: 50000, Name: itemName,
        matchmode: "Substring" }], attempted)
    if !menuItem
        return { ok: false, attempted: attempted }
    if !MobillsAuto_Click(menuItem)
        return { ok: false, attempted: attempted }
    return { ok: true, attempted: attempted }
}

MobillsAuto_ClickMenu(uia, autoId, btnName) {
    attempted := []
    btn := MobillsAuto_Resolve(uia, [{ Type: 50000, AutomationId: autoId }, { Type: 50000, Name: btnName }], attempted)
    if !btn
        return { ok: false, attempted: attempted }
    if !MobillsAuto_Click(btn)
        return { ok: false, attempted: attempted }
    return { ok: true, attempted: attempted }
}

MobillsAuto_FindAmountEdit(scope) {
    if !scope
        return ""
    try {
        edits := scope.FindAll({ Type: 50004 })
        if edits {
            for e in edits {
                try {
                    nm := e.Name
                    if (nm = "Description" || nm = "Origin account" || nm = "Destination account")
                        continue
                } catch {
                }
                try {
                    v := e.Value
                    if (v != "" && RegExMatch(v, "^[\d.,]+$"))
                        return e
                } catch {
                }
            }
        }
    } catch {
    }
    return ""
}

MobillsAuto_FindNamedEdit(scope, name) {
    if !scope
        return ""
    try {
        el := scope.FindFirst({ Type: 50004, Name: name })
        if el
            return el
    } catch {
    }
    try {
        el := scope.FindElement({ Type: 50004, Name: name, matchmode: "Substring" })
        if el
            return el
    } catch {
    }
    return ""
}

MobillsAuto_DialogCombos(scope) {
    list := []
    if !scope
        return list
    try {
        found := scope.FindAll({ Type: 50003 })
        if found {
            for c in found
                list.Push(c)
        }
    } catch {
    }
    return list
}

MobillsAuto_FindSaveButton(scope) {
    if !scope
        return ""
    try {
        btn := scope.FindFirst({ Type: 50000, Name: "SAVE" })
        if btn
            return btn
    } catch {
    }
    try {
        btns := scope.FindAll({ Type: 50000 })
        if btns {
            for b in btns {
                try {
                    if (b.Name = "SAVE")
                        return b
                } catch {
                }
            }
        }
    } catch {
    }
    return ""
}

MobillsAuto_ClickLeft(el) {
    if !el
        return false
    try {
        el.Click("left")
        return true
    } catch {
        try {
            el.Click("left")
            return true
        } catch {
            return false
        }
    }
}

MobillsAuto_SelectedPageNumber(uia) {
    if !uia
        return 0
    try {
        nav := uia.FindElement({ Name: "pagination navigation", Type: 50026, matchmode: "Substring" })
        if !nav
            return 0
        buttons := nav.FindAll({ Type: 50000 })
        if !buttons
            return 0
        for b in buttons {
            cls := ""
            try cls := b.ClassName
            if !InStr(cls, "Mui-selected")
                continue
            nm := ""
            try nm := b.Name
            if RegExMatch(nm, "(\d+)", &m)
                return Integer(m[1])
        }
    } catch {
    }
    return 0
}

MobillsAuto_FindPageButton(uia, pageN) {
    if !uia || pageN < 1
        return ""
    try {
        el := uia.FindElement({ Type: 50000, Name: "Go to page " pageN })
        if el
            return el
    } catch {
    }
    try {
        el := uia.FindElement({ Type: 50000, Name: "page " pageN })
        if el
            return el
    } catch {
    }
    return ""
}

MobillsAuto_NextPage(uia, fingerprint := "") {
    attempted := []
    cur := MobillsAuto_SelectedPageNumber(uia)
    nextN := (cur > 0) ? cur + 1 : 2
    btn := MobillsAuto_FindPageButton(uia, nextN)
    if btn
        attempted.Push("Go to page " nextN)
    if !btn {
        btn := MobillsAuto_Resolve(uia, [{ Type: 50000, Name: "Go to next page" }, { Type: 50000, Name: "next page",
            matchmode: "Substring" }], attempted)
    }
    if !btn {
        try {
            nav := uia.FindElement({ Name: "pagination navigation", Type: 50026, matchmode: "Substring" })
            if nav {
                buttons := nav.FindAll({ Type: 50000 })
                if (buttons && buttons.Length)
                    btn := buttons[buttons.Length]
            }
        } catch {
        }
        attempted.Push("pagination navigation last button")
    }
    if !btn
        return { ok: false, done: true, attempted: attempted }
    if MobillsAuto_IsDisabled(btn)
        return { ok: true, done: true, attempted: attempted }
    if !MobillsAuto_ClickLeft(btn)
        return { ok: false, done: false, attempted: attempted }
    Sleep MOBILLS_STEP_MS * 2
    uia := MobillsAuto_RefreshUia(uia)
    sel := MobillsAuto_SelectedPageNumber(uia)
    if (sel = cur || sel = 0) {
        uia := MobillsAuto_RefreshUia(uia)
        btn2 := MobillsAuto_FindPageButton(uia, nextN)
        if !btn2
            btn2 := btn
        if !MobillsAuto_ClickLeft(btn2)
            return { ok: false, done: false, attempted: attempted }
        attempted.Push("second left click")
        Sleep MOBILLS_STEP_MS * 2
        uia := MobillsAuto_RefreshUia(uia)
    }
    endTick := A_TickCount + 4000
    while (A_TickCount < endTick) {
        uia := MobillsAuto_RefreshUia(uia)
        sel := MobillsAuto_SelectedPageNumber(uia)
        if (sel = nextN)
            return { ok: true, done: false, attempted: attempted }
        Sleep 150
    }
    sel := MobillsAuto_SelectedPageNumber(uia)
    if (sel = nextN)
        return { ok: true, done: false, attempted: attempted }
    if (sel = cur || sel = 0)
        return { ok: false, done: true, attempted: attempted }
    return { ok: true, done: false, attempted: attempted }
}

MobillsAuto_RefreshUia(uia) {
    try {
        att := MobillsAuto_AttachBrowser()
        if (att.uia)
            return att.uia
    } catch {
    }
    return uia
}
