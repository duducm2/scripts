; =============================================================================
; Shift keys module: mobills_pagination.ahk
; Mobills pagination unified
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

; =============================================================================
; Mobills pagination (unified)
; =============================================================================
Mobills_GetContext(uia) {
    try {
        url := uia.GetCurrentURL()
        url := StrLower(url)
    } catch {
        url := ""
    }
    if InStr(url, "/transactions")
        return "transactions"
    if InStr(url, "/accounts")
        return "accounts"
    ; Budgets has its own pager controls (avoid selecting month/year text)
    if InStr(url, "/budgets")
        return "budgets"
    if InStr(url, "/planning")
        return "planning"
    return "unknown"
}

Mobills_IsDisabled(el) {
    try {
        ; Gate 3: Visibility & Interactivity Checks
        if !el.GetPropertyValue(UIA.Property.IsEnabled)
            return true

        if el.GetPropertyValue(UIA.Property.IsOffscreen)
            return true

        cls := ""
        try cls := el.GetPropertyValue(UIA.Property.ClassName)
        if (cls != "" && InStr(cls, "Mui-disabled"))
            return true

        ; Check aria-hidden or other attributes if possible
        try {
            ariaHidden := el.GetPropertyValue(UIA.Property.AriaProperties)
            if (ariaHidden != "" && InStr(ariaHidden, "hidden=true"))
                return true
        } catch {
        }
    } catch {
    }
    return false
}

Mobills_IsButton(el) {
    if !el
        return false
    try {
        ct := el.GetPropertyValue(UIA.Property.ControlType)
        return (ct = UIA.Type.Button || ct = 50000)
    } catch {
        ; Some wrappers expose .Type directly
        try {
            return (el.Type = 50000 || el.Type = UIA.Type.Button)
        } catch {
            return false
        }
    }
}

Mobills_FindElementByCandidates(uia, candidates) {
    if !uia
        return ""
    for , c in candidates {
        try {
            el := uia.FindElement(c)
            if el
                return el
        } catch {
        }
    }
    return ""
}

Mobills_FindOpenButton(uia, index := 1) {
    if !uia
        return ""
    try {
        openButtons := uia.FindAll({ Name: "Open", Type: 50000 })
        if (openButtons && openButtons.Length >= index)
            return openButtons[index]
    } catch {
    }
    return ""
}

; Budgets: resolve BOTH arrows by index, then pick by dir.
; Target:
;   Prev => {T:30}, {T:26}, {T:0, i:7}
;   Next => {T:30}, {T:26}, {T:0, i:8}
Mobills_GetBudgetsPrevNext(uia, &prevBtn, &nextBtn) {
    prevBtn := ""
    nextBtn := ""

    prev := ""
    next := ""
    try {
        prev := uia.ElementFromPath({ Type: 30 }, { Type: 26 }, { Type: 0, i: 7 })
    } catch {
        prev := ""
    }
    try {
        next := uia.ElementFromPath({ Type: 30 }, { Type: 26 }, { Type: 0, i: 8 })
    } catch {
        next := ""
    }

    if (prev && Mobills_IsButton(prev) && !Mobills_IsDisabled(prev))
        prevBtn := prev
    if (next && Mobills_IsButton(next) && !Mobills_IsDisabled(next))
        nextBtn := next

    ; Sanity check: ensure prev is physically left of next (swap if needed).
    if (prevBtn && nextBtn) {
        try {
            p := prevBtn.Location
            n := nextBtn.Location
            if (p.x > n.x) {
                tmp := prevBtn
                prevBtn := nextBtn
                nextBtn := tmp
            }
        } catch {
        }
    }

    return (prevBtn || nextBtn)
}

Mobills_FindPagerByName(uia, dir) {
    ; Try common labels (EN/PT). Substring match.
    namesPrev := ["Go to previous page", "previous page", "Previous", "Prev", "Anterior", "Página anterior",
        "Ir para a página anterior"]
    namesNext := ["Go to next page", "next page", "Next", "Próximo", "Proximo", "Página seguinte",
        "Ir para a próxima página",
        "Ir para a proxima página"]
    names := (dir = "Prev") ? namesPrev : namesNext

    for , nm in names {
        try {
            btn := uia.FindElement({ Name: nm, Type: 50000, matchmode: "Substring" })
            if btn && !Mobills_IsDisabled(btn)
                return btn
        } catch {
        }
    }
    return ""
}

Mobills_FindPagerByStructure(uia, dir) {
    ; Gate 4: Fallback Spatial & Structural Discovery
    ; Look for Group named "pagination navigation" or similar
    try {
        navGroup := uia.FindElement({ Name: "pagination navigation", Type: "Group", matchmode: "Substring" })
        if navGroup {
            buttons := navGroup.FindAll({ Type: "Button" })
            if (IsObject(buttons) && buttons.Length > 0) {
                if (dir = "Prev") {
                    btn := buttons[1]
                    if !Mobills_IsDisabled(btn)
                        return btn
                } else {
                    btn := buttons[buttons.Length]
                    if !Mobills_IsDisabled(btn)
                        return btn
                }
            }
        }
    } catch {
    }

    ; Look for List -> ListItem -> Button with Pagination classes
    try {
        lists := uia.FindAll({ Type: "List" })
        if IsObject(lists) {
            for , lst in lists {
                try {
                    cls := lst.GetPropertyValue(UIA.Property.ClassName)
                    if InStr(cls, "Pagination") {
                        buttons := lst.FindAll({ Type: "Button" })
                        if (IsObject(buttons) && buttons.Length > 0) {
                            if (dir = "Prev") {
                                btn := buttons[1]
                                if !Mobills_IsDisabled(btn)
                                    return btn
                            } else {
                                btn := buttons[buttons.Length]
                                if !Mobills_IsDisabled(btn)
                                    return btn
                            }
                        }
                    }
                } catch {
                }
            }
        }
    } catch {
    }
    return ""
}

Mobills_FindPagerByMonthHeader(uia, dir) {
    grp := ""
    try grp := FindMonthGroup(uia)
    if !grp
        return ""

    step := (dir = "Prev") ? "-1" : "+1"
    try {
        btn := grp.WalkTree(step, { Type: "Button" })
        if btn && !Mobills_IsDisabled(btn)
            return btn
    } catch {
    }

    ; fallback: scan sibling buttons in parent and choose closest left/right
    try parent := UIA.TreeWalkerTrue.GetParentElement(grp)
    if !parent
        return ""

    try grpPos := grp.Location
    best := ""
    bestX := ""
    try {
        for , el in parent.FindAll({ Type: "Button" }) {
            if Mobills_IsDisabled(el)
                continue
            pos := el.Location
            sameRow := (pos.y >= grpPos.y - 10 && pos.y <= grpPos.y + grpPos.h + 10)
            if !sameRow
                continue
            if (dir = "Prev") {
                if (pos.x < grpPos.x) {
                    if (best = "" || pos.x > bestX) {
                        best := el
                        bestX := pos.x
                    }
                }
            } else {
                if (pos.x > grpPos.x + grpPos.w) {
                    if (best = "" || pos.x < bestX) {
                        best := el
                        bestX := pos.x
                    }
                }
            }
        }
    } catch {
    }
    return best
}

; Transactions month navigation: click the arrows adjacent to the month/year display
; (e.g. "<  February 2026  >"), NOT the table pagination ("previous page / next page").
Mobills_FindMonthNavByMonthYear(uia, dir) {
    if !uia
        return ""

    ; Find the visible month label (Text) in any supported language
    months := ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
        "November",
        "December",
        "Janeiro", "Fevereiro", "Março", "Marco", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro",
        "Outubro",
        "Novembro", "Dezembro"]

    monthEl := ""
    for , m in months {
        try {
            el := uia.FindElement({ Name: m, Type: "Text", mm: 1, cs: false })
            if el {
                monthEl := el
                break
            }
        } catch {
        }
    }
    if !monthEl
        return ""

    ; Use the month element bounds as the anchor; choose closest button left/right on the same row.
    try mPos := monthEl.Location

    ; Search for buttons in a broader scope (e.g., the parent of the parent)
    header := ""
    try header := monthEl.WalkTree("p", { Type: "Group" })
    if header {
        try {
            parentHeader := UIA.TreeWalkerTrue.GetParentElement(header)
            if parentHeader
                header := parentHeader
        } catch {
        }
    }

    buttons := ""
    if header {
        try buttons := header.FindAll({ Type: "Button" })
    }
    if (!IsObject(buttons) || buttons.Length = 0) {
        ; Fallback to all buttons in the document if header search fails
        try buttons := uia.FindAll({ Type: "Button" })
    }
    if (!IsObject(buttons) || buttons.Length = 0)
        return ""

    best := ""
    bestDist := ""
    try {
        for , b in buttons {
            if Mobills_IsDisabled(b)
                continue
            pos := b.Location
            ; same row as the month label
            if !(pos.y <= (mPos.y + mPos.h + 12) && (pos.y + pos.h) >= (mPos.y - 12))
                continue

            if (dir = "Prev") {
                if (pos.x + pos.w) >= mPos.x
                    continue
                dist := mPos.x - (pos.x + pos.w)
            } else {
                if pos.x <= (mPos.x + mPos.w)
                    continue
                dist := pos.x - (mPos.x + mPos.w)
            }

            if (best = "" || dist < bestDist) {
                best := b
                bestDist := dist
            }
        }
    } catch {
    }

    return best
}

Mobills_FindPagerByPath(uia, dir, context) {
    ; Budgets page: force the known arrow BUTTONs and avoid adjacent Text elements.
    ; Target:
    ;   Next  => {T:30}, {T:26}, {T:0, i:8}
    ;   Prev  => {T:30}, {T:26}, {T:0, i:7}
    if (context = "budgets") {
        if Mobills_GetBudgetsPrevNext(uia, &prevBtn, &nextBtn) {
            if (dir = "Prev")
                return prevBtn ? prevBtn : ""
            return nextBtn ? nextBtn : ""
        }
        ; IMPORTANT: Do not fall through on budgets (prevents picking the wrong control)
        return ""
    }

    ; Legacy paths (worked across some pages previously)
    try {
        if (dir = "Prev") {
            btn := uia.ElementFromPath({ Type: 30 }, { Type: 26 }, { Type: 26 }, { Type: 8 }, { Type: 7 }, { Type: 0 })
        } else {
            btn := uia.ElementFromPath({ Type: 30 }, { Type: 26 }, { Type: 26 }, { Type: 8 }, { Type: 7, i: -1 }, { Type: 0 })
        }
        if btn && Mobills_IsButton(btn) && !Mobills_IsDisabled(btn)
            return btn
    } catch {
    }

    ; Accounts: prefer toolbar-ish container if present (avoid hardcoding a single index)
    if (context = "accounts") {
        try {
            toolbar := uia.ElementFromPath({ Type: 30 }, { Type: 26 })
            if toolbar {
                buttons := toolbar.FindAll({ Type: "Button" })
                best := ""
                bestKey := ""
                for , b in buttons {
                    if Mobills_IsDisabled(b)
                        continue
                    pos := b.Location
                    ; prefer top-most row
                    key := pos.y * 100000 + pos.x
                    if (dir = "Prev") {
                        ; left-most among top candidates
                        if (best = "" || key < bestKey) {
                            best := b
                            bestKey := key
                        }
                    } else {
                        ; right-most among top candidates (approx: invert x)
                        key2 := pos.y * 100000 - pos.x
                        if (best = "" || key2 < bestKey) {
                            best := b
                            bestKey := key2
                        }
                    }
                }
                if best
                    return best
            }
        } catch {
        }
    }

    ; Planning: sometimes the "button" is a Text element
    if (context = "planning") {
        try {
            el := uia.ElementFromPath({ Type: 30 }, { Type: 26 }, { Type: 20, i: 2 })
            if el
                return el
        } catch {
        }
    }

    return ""
}

Mobills_ClickPager(el) {
    if !el
        return false
    try {
        ; Prefer Invoke when available, otherwise click.
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
        return false
    }
}
