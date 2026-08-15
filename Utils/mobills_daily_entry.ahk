; =============================================================================
; Utils module: mobills_daily_entry.ahk
; #!+U Macros [m] — parse Desktop MOBILLS_V1 file, confirm, enter in Mobills, scrape.
; =============================================================================

global g_MobillsDailyConfirmResult := ""
global g_MobillsDailyConfirmGui := ""
global g_MobillsDailyReviewResult := ""
global g_MobillsDailyReviewGui := ""

MobillsDaily_LoadIniCatalog(path) {
    catalog := Map()
    if (!FileExist(path))
        return catalog
    raw := ""
    try raw := FileRead(path, "UTF-8")
    catch {
        return catalog
    }
    current := ""
    loop parse raw, "`n", "`r" {
        line := Trim(A_LoopField)
        if (line = "" || SubStr(line, 1, 1) = ";" || SubStr(line, 1, 1) = "#")
            continue
        if (RegExMatch(line, "^\[(.+)\]$", &m)) {
            current := m[1]
            catalog[current] := { desc: "", subs: Map() }
            continue
        }
        if (current = "")
            continue
        eq := InStr(line, "=")
        if (!eq)
            continue
        k := Trim(SubStr(line, 1, eq - 1))
        v := Trim(SubStr(line, eq + 1))
        if (k = "geral")
            catalog[current].desc := v
        else
            catalog[current].subs[k] := v
    }
    return catalog
}

MobillsDaily_RenderIniCatalog(path, heading) {
    catalog := MobillsDaily_LoadIniCatalog(path)
    if (!catalog.Count)
        return heading . ":`n(file missing: " path ")"
    out := heading . ":"
    for section, rec in catalog {
        out .= "`n- " . section
        if (rec.desc != "")
            out .= " — " . rec.desc
        for sub, sdesc in rec.subs {
            out .= "`n  - " . section . " > " . sub
            if (sdesc != "")
                out .= " — " . sdesc
        }
    }
    return out
}

MobillsDaily_CatalogHasSection(catalog, name) {
    if (name = "" || !catalog)
        return false
    if catalog.Has(name)
        return true
    for section, rec in catalog {
        if (StrLower(section) = StrLower(name))
            return true
        for sub, _ in rec.subs {
            if (StrLower(sub) = StrLower(name))
                return true
        }
    }
    return false
}

MobillsDaily_MatchAccountName(catalog, uiName) {
    uiName := Trim(uiName)
    if (uiName = "")
        return ""
    uiName := RegExReplace(uiName, "\.{2,}$", "")
    uiName := Trim(uiName)
    if catalog.Has(uiName)
        return uiName
    best := ""
    bestLen := 0
    for section, _ in catalog {
        if (InStr(section, uiName) = 1 || InStr(uiName, section) = 1) {
            if (StrLen(section) > bestLen) {
                best := section
                bestLen := StrLen(section)
            }
        }
    }
    return best
}

MobillsDaily_DesktopPath() {
    p := ""
    try p := GetDesktopToRecyclePath()
    catch {
        p := ""
    }
    if (!p || p = "" || !DirExist(p))
        p := A_Desktop
    return p
}

MobillsDaily_FindLatestFile() {
    desktop := MobillsDaily_DesktopPath()
    if (!DirExist(desktop))
        return ""
    newestPath := ""
    newestStamp := ""
    loop files desktop "\*", "F" {
        if (StrLower(A_LoopFileName) = "desktop.ini")
            continue
        stamp := ""
        try {
            tC := FileGetTime(A_LoopFileFullPath, "C")
            tM := FileGetTime(A_LoopFileFullPath, "M")
            stamp := (tC >= tM) ? tC : tM
        } catch {
            continue
        }
        first := ""
        try {
            raw := FileRead(A_LoopFileFullPath, "UTF-8")
            loop parse raw, "`n", "`r" {
                first := Trim(A_LoopField)
                break
            }
        } catch {
            continue
        }
        if (first != "MOBILLS_V1")
            continue
        if (newestStamp = "" || stamp > newestStamp) {
            newestStamp := stamp
            newestPath := A_LoopFileFullPath
        }
    }
    return newestPath
}

MobillsDaily_ParseFile(path) {
    result := { rows: [], unmapped: [], error: "" }
    raw := ""
    try raw := FileRead(path, "UTF-8")
    catch as e {
        result.error := "Could not read file: " e.Message
        return result
    }
    accounts := MobillsDaily_LoadIniCatalog(A_ScriptDir "\accounts.ini")
    expenses := MobillsDaily_LoadIniCatalog(A_ScriptDir "\categories-expenses.ini")
    incomes := MobillsDaily_LoadIniCatalog(A_ScriptDir "\categories-income.ini")
    started := false
    loop parse raw, "`n", "`r" {
        line := Trim(A_LoopField)
        if (line = "")
            continue
        if (line = "MOBILLS_V1") {
            started := true
            continue
        }
        if (!started)
            continue
        if (line = "END")
            break
        if (SubStr(line, 1, 1) = "#") {
            result.unmapped.Push(line)
            continue
        }
        if (InStr(line, "TYPE|DESCRIPTION|VALUE") = 1)
            continue
        parts := StrSplit(line, "|")
        if (parts.Length < 7) {
            result.unmapped.Push("# UNMAPPED: bad column count: " line)
            continue
        }
        typ := StrUpper(Trim(parts[1]))
        desc := Trim(parts[2])
        val := Trim(parts[3])
        source := Trim(parts[4])
        target := Trim(parts[5])
        cat := Trim(parts[6])
        sub := Trim(parts[7])
        flags := []
        if !(typ = "EXPENSE" || typ = "INCOME" || typ = "CARD" || typ = "TRANSFER")
            flags.Push("bad type")
        if !RegExMatch(val, "^[0-9]+([.][0-9]+)?$")
            flags.Push("bad value")
        if (typ = "EXPENSE" || typ = "TRANSFER") {
            if (source != "" && !MobillsDaily_CatalogHasSection(accounts, source))
                flags.Push("unknown source account")
        }
        if (typ = "INCOME") {
            dest := (target != "") ? target : source
            if (dest != "" && !MobillsDaily_CatalogHasSection(accounts, dest))
                flags.Push("unknown dest account")
        }
        if (typ = "TRANSFER") {
            if (target != "" && !MobillsDaily_CatalogHasSection(accounts, target))
                flags.Push("unknown dest account")
        }
        if (typ = "CARD" && source = "")
            source := MOBILLS_CARD_NAME
        if (typ = "EXPENSE" || typ = "CARD") {
            if (cat != "" && !MobillsDaily_CatalogHasSection(expenses, cat) && !MobillsDaily_CatalogHasSection(expenses,
                sub))
                flags.Push("unknown expense category")
        }
        if (typ = "INCOME") {
            if (cat != "" && !MobillsDaily_CatalogHasSection(incomes, cat) && !MobillsDaily_CatalogHasSection(incomes,
                sub))
                flags.Push("unknown income category")
        }
        result.rows.Push({
            type: typ, description: desc, value: val, source: source, target: target,
            category: cat, subcategory: sub, flags: flags, raw: line
        })
    }
    if (!started)
        result.error := "File does not contain MOBILLS_V1."
    return result
}

MobillsDaily_CategoryLabel(row) {
    if (row.subcategory != "")
        return row.category . " > " . row.subcategory
    return row.category
}

MobillsDaily_ShowConfirmTable(rows, unmapped, filePath) {
    global g_MobillsDailyConfirmResult, g_MobillsDailyConfirmGui
    g_MobillsDailyConfirmResult := ""
    if (IsObject(g_MobillsDailyConfirmGui)) {
        try g_MobillsDailyConfirmGui.Destroy()
        catch {
        }
    }
    hint := "File: " filePath
    if (unmapped.Length)
        hint .= "`nUnmapped lines: " unmapped.Length . " (see rows marked ! )"
    dlg := Gui("+AlwaysOnTop +ToolWindow", "Mobills daily entry")
    dlg.SetFont("s10", "Segoe UI")
    dlg.Add("Text", "w900", hint)
    lv := dlg.Add("ListView", "w900 h360 -Multi", ["#", "Type", "Description", "Value", "Source", "Target", "Category"])
    i := 1
    for row in rows {
        mark := row.flags.Length ? "! " : ""
        lv.Add("", mark i, row.type, row.description, row.value, row.source, row.target, MobillsDaily_CategoryLabel(row
        ))
        i++
    }
    try lv.ModifyCol(1, 40)
    try lv.ModifyCol(2, 90)
    try lv.ModifyCol(3, 220)
    try lv.ModifyCol(4, 80)
    try lv.ModifyCol(5, 180)
    try lv.ModifyCol(6, 160)
    try lv.ModifyCol(7, 160)
    dlg.Add("Button", "w100 Section Default", "OK").OnEvent("Click", MobillsDaily_ConfirmOk)
    dlg.Add("Button", "w100 ys", "Cancel").OnEvent("Click", MobillsDaily_ConfirmCancel)
    dlg.OnEvent("Close", MobillsDaily_ConfirmCancel)
    dlg.OnEvent("Escape", MobillsDaily_ConfirmCancel)
    g_MobillsDailyConfirmGui := dlg
    dlg.Show()
    try lv.Focus()
    catch {
    }
    start := A_TickCount
    while (g_MobillsDailyConfirmResult = "") {
        if ((A_TickCount - start) >= 300000) {
            g_MobillsDailyConfirmResult := "cancel"
            break
        }
        Sleep 50
    }
    res := g_MobillsDailyConfirmResult
    g_MobillsDailyConfirmResult := ""
    try dlg.Destroy()
    catch {
    }
    g_MobillsDailyConfirmGui := ""
    return res = "ok"
}

MobillsDaily_ConfirmOk(*) {
    global g_MobillsDailyConfirmResult
    if (g_MobillsDailyConfirmResult != "")
        return
    g_MobillsDailyConfirmResult := "ok"
}

MobillsDaily_ConfirmCancel(*) {
    global g_MobillsDailyConfirmResult
    if (g_MobillsDailyConfirmResult != "")
        return
    g_MobillsDailyConfirmResult := "cancel"
}

MobillsDaily_Halt(msg) {
    try FileAppend(FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss") "`t" msg "`n", A_ScriptDir "\mobills-run-error.log",
    "UTF-8")
    catch {
    }
    dlg := Gui("+AlwaysOnTop +ToolWindow", "Mobills daily entry — halted")
    dlg.SetFont("s10", "Segoe UI")
    dlg.Add("Text", "w720", "The run stopped. Nothing after this transaction was entered.")
    edit := dlg.Add("Edit", "w720 h260 ReadOnly -Wrap", msg)
    dlg.Add("Button", "w100 Default", "Close").OnEvent("Click", (*) => dlg.Destroy())
    dlg.OnEvent("Close", (*) => dlg.Destroy())
    dlg.OnEvent("Escape", (*) => dlg.Destroy())
    dlg.Show()
    try edit.Focus()
    catch {
    }
    WinWaitClose("ahk_id " dlg.Hwnd)
}

MobillsDaily_Fail(idx, field, gate, attempted, extra := "") {
    att := ""
    if (IsObject(attempted)) {
        for a in attempted
            att .= "`n  - " a
    } else if (attempted != "") {
        att := "`n  - " attempted
    }
    msg := "Transaction #" idx "`nField: " field "`nGate: " gate
    if (extra != "")
        msg .= "`nDetail: " extra
    if (att != "")
        msg .= "`nSelectors attempted:" att
    MobillsDaily_Halt(msg)
}

MobillsDaily_ExpectedTitle(typ) {
    switch typ {
        case "EXPENSE": return "New expense"
        case "INCOME": return "New income"
        case "CARD": return "New credit card expense"
        case "TRANSFER": return "New transfer"
        default: return ""
    }
}

MobillsDaily_MenuItem(typ) {
    switch typ {
        case "EXPENSE": return "Expense"
        case "INCOME": return "Income"
        case "CARD": return "Credit card expense"
        case "TRANSFER": return "Transfer"
        default: return ""
    }
}

MobillsDaily_DialogGone(uia) {
    return (MobillsAuto_DialogTitle(uia) = "") ? true : ""
}

MobillsDaily_TitleIfExpected(uia, expected) {
    d := MobillsAuto_FindDialog(uia)
    t := MobillsAuto_DialogTitle(d)
    if (t = "")
        t := MobillsAuto_DialogTitle(uia)
    return (t = expected) ? expected : ""
}

MobillsDaily_SaveIfEnabled(scope) {
    b := MobillsAuto_FindSaveButton(scope)
    if (b && !MobillsAuto_IsDisabled(b))
        return b
    return ""
}

MobillsDaily_EnterRow(uia, row, idx, total) {
    StandardLoadingBar_Update("⏳ " idx "/" total " — " row.type ": " row.description " R$ " row.value)
    menuName := MobillsDaily_MenuItem(row.type)
    opened := MobillsAuto_SelectNewMenuItem(uia, menuName)
    if (!opened.ok) {
        MobillsDaily_Fail(idx, "New menu", "Could not click + then " menuName, opened.attempted)
        return false
    }
    expected := MobillsDaily_ExpectedTitle(row.type)
    title := MobillsAuto_WaitFor(MobillsDaily_TitleIfExpected.Bind(uia, expected))
    if (title != expected) {
        MobillsDaily_Fail(idx, "Dialog title", "Expected '" expected "'", ["Text Name=" expected], "got '" title "'")
        return false
    }
    dialog := MobillsAuto_FindDialog(uia)
    scope := dialog ? dialog : uia
    Sleep MOBILLS_STEP_MS

    amount := MobillsAuto_FindAmountEdit(scope)
    if !amount {
        MobillsDaily_Fail(idx, "Value", "Amount edit not found", ["Edit with numeric Value", "sibling of R$"])
        return false
    }
    setAmt := MobillsAuto_SetEditVerified(amount, row.value)
    if (!setAmt.ok) {
        MobillsDaily_Fail(idx, "Value", "Read-back mismatch", setAmt.attempted, "wanted " row.value " got " setAmt.got)
        return false
    }

    if (row.type = "TRANSFER") {
        combos := MobillsAuto_DialogCombos(scope)
        originEdit := MobillsAuto_FindNamedEdit(scope, "Origin account")
        destEdit := MobillsAuto_FindNamedEdit(scope, "Destination account")
        originCombo := (combos.Length >= 1) ? combos[1] : ""
        destCombo := (combos.Length >= 2) ? combos[2] : ""
        if (originCombo) {
            pk := MobillsAuto_PickAutocomplete(originCombo, row.source)
            if (!pk.ok) {
                MobillsDaily_Fail(idx, "Origin account", "Autocomplete mismatch", pk.attempted, "wanted " row.source " got " pk
                    .got)
                return false
            }
        } else if (originEdit) {
            setO := MobillsAuto_SetEditVerified(originEdit, row.source)
            if (!setO.ok) {
                MobillsDaily_Fail(idx, "Origin account", "Edit mismatch", setO.attempted, "wanted " row.source " got " setO
                    .got)
                return false
            }
        } else {
            MobillsDaily_Fail(idx, "Origin account", "Field not found", ["Name=Origin account", "ComboBox 1"])
            return false
        }
        if (destCombo) {
            pk := MobillsAuto_PickAutocomplete(destCombo, row.target)
            if (!pk.ok) {
                MobillsDaily_Fail(idx, "Destination account", "Autocomplete mismatch", pk.attempted, "wanted " row.target " got " pk
                    .got)
                return false
            }
        } else if (destEdit) {
            setD := MobillsAuto_SetEditVerified(destEdit, row.target)
            if (!setD.ok) {
                MobillsDaily_Fail(idx, "Destination account", "Edit mismatch", setD.attempted, "wanted " row.target " got " setD
                    .got)
                return false
            }
        } else {
            MobillsDaily_Fail(idx, "Destination account", "Field not found", ["Name=Destination account", "ComboBox 2"])
            return false
        }
    } else {
        descEl := MobillsAuto_FindNamedEdit(scope, "Description")
        if !descEl {
            MobillsDaily_Fail(idx, "Description", "Edit not found", ["Name=Description"])
            return false
        }
        setD := MobillsAuto_SetEditVerified(descEl, row.description)
        if (!setD.ok) {
            MobillsDaily_Fail(idx, "Description", "Read-back mismatch", setD.attempted, "wanted " row.description " got " setD
                .got)
            return false
        }
        combos := MobillsAuto_DialogCombos(scope)
        catCombo := (combos.Length >= 2) ? combos[2] : ""
        acctCombo := (combos.Length >= 3) ? combos[3] : ""
        catWanted := row.subcategory != "" ? row.subcategory : row.category
        if (catCombo && catWanted != "") {
            pk := MobillsAuto_PickAutocomplete(catCombo, catWanted)
            if (!pk.ok && row.subcategory != "" && row.category != "")
                pk := MobillsAuto_PickAutocomplete(catCombo, row.category)
            if (!pk.ok) {
                MobillsDaily_Fail(idx, "Category", "Autocomplete mismatch", pk.attempted, "wanted " catWanted " got " pk
                    .got)
                return false
            }
        }
        acctWanted := (row.type = "CARD") ? MOBILLS_CARD_NAME : ((row.type = "INCOME") ? (row.target != "" ? row.target :
            row.source) : row.source)
        if (acctCombo && acctWanted != "") {
            pk := MobillsAuto_PickAutocomplete(acctCombo, acctWanted)
            if (!pk.ok) {
                MobillsDaily_Fail(idx, (row.type = "CARD") ? "Card" : "Account", "Autocomplete mismatch", pk.attempted,
                "wanted " acctWanted " got " pk.got)
                return false
            }
        }
    }

    saveBtn := MobillsAuto_WaitFor(MobillsDaily_SaveIfEnabled.Bind(scope))
    if !saveBtn {
        MobillsDaily_Fail(idx, "SAVE", "Button stayed disabled or missing", ["Name=SAVE", "Mui-disabled class"])
        return false
    }
    if !MobillsAuto_Click(saveBtn) {
        MobillsDaily_Fail(idx, "SAVE", "Click failed", ["Name=SAVE Invoke/Click"])
        return false
    }
    gone := MobillsAuto_WaitFor(MobillsDaily_DialogGone.Bind(uia))
    if !gone {
        MobillsDaily_Fail(idx, "SAVE", "Dialog still open after SAVE", ["heading gone: " expected])
        return false
    }
    Sleep MOBILLS_STEP_MS
    return true
}

MobillsDaily_IsMoneyText(nm) {
    return (InStr(nm, "R$") = 1 || RegExMatch(nm, "^R\$"))
}

MobillsDaily_ScrapeAccounts(uia) {
    catalog := MobillsDaily_LoadIniCatalog(A_ScriptDir "\accounts.ini")
    nav := MobillsAuto_ClickMenu(uia, "menu-accounts-item", "Accounts")
    if (!nav.ok)
        return { ok: false, rows: [], error: "Could not open Accounts page", attempted: nav.attempted }
    Sleep MOBILLS_STEP_MS * 2
    all := []
    seenPages := 0
    loop 8 {
        seenPages++
        StandardLoadingBar_Update("⏳ Scraping accounts page " seenPages "...")
        pageRows := MobillsDaily_ScrapeAccountsPage(uia, catalog)
        for r in pageRows
            all.Push(r)
        nxt := MobillsAuto_NextPage(uia)
        if (nxt.done || !nxt.ok)
            break
        Sleep MOBILLS_STEP_MS * 2
    }
    return { ok: true, rows: all, error: "" }
}

MobillsDaily_ScrapeAccountsPage(uia, catalog) {
    rows := []
    texts := ""
    try texts := uia.FindAll({ Type: 50020 })
    catch {
        return rows
    }
    if !texts
        return rows
    names := []
    classes := []
    i := 1
    while (i <= texts.Length) {
        nm := ""
        cls := ""
        try nm := texts[i].Name
        try cls := texts[i].ClassName
        names.Push(nm)
        classes.Push(cls)
        i++
    }
    i := 1
    while (i <= names.Length) {
        if (names[i] = "Current balance") {
            if InStr(classes[i], "MuiTypography-h5") {
                i++
                continue
            }
            money := ""
            if (i < names.Length && MobillsDaily_IsMoneyText(names[i + 1]))
                money := names[i + 1]
            acct := ""
            j := i - 1
            while (j >= 1) {
                cand := MobillsDaily_MatchAccountName(catalog, names[j])
                if (cand != "") {
                    acct := cand
                    break
                }
                j--
            }
            if (acct != "" && money != "")
                rows.Push({ kind: "account", name: acct, label: "Current balance", value: money })
        }
        i++
    }
    return rows
}

MobillsDaily_ScrapeCards(uia) {
    nav := MobillsAuto_ClickMenu(uia, "menu-creditCards-item", "Credit cards")
    if (!nav.ok)
        return { ok: false, rows: [], error: "Could not open Credit cards page", attempted: nav.attempted }
    Sleep MOBILLS_STEP_MS * 2
    StandardLoadingBar_Update("⏳ Scraping credit cards...")
    rows := []
    texts := ""
    try texts := uia.FindAll({ Type: 50020 })
    catch {
        return { ok: true, rows: rows, error: "" }
    }
    if !texts
        return { ok: true, rows: rows, error: "" }
    names := []
    for t in texts {
        try names.Push(t.Name)
        catch {
            names.Push("")
        }
    }
    card := ""
    sawOpen := false
    sawPartial := false
    openVal := ""
    partialVal := ""
    i := 1
    while (i <= names.Length) {
        nm := names[i]
        if (nm = MOBILLS_CARD_NAME && card = "")
            card := nm
        if (nm = "Open invoice") {
            sawOpen := true
            sawPartial := false
        }
        if (nm = "Partial value")
            sawPartial := true
        if (MobillsDaily_IsMoneyText(nm)) {
            if (sawPartial && partialVal = "") {
                partialVal := nm
                if (openVal = "")
                    openVal := nm
            } else if (sawOpen && openVal = "") {
                openVal := nm
            }
        }
        if (nm = "Available Limit" || nm = "PAY INVOICE" || nm = "Total amount")
            break
        i++
    }
    if (card = "")
        card := MOBILLS_CARD_NAME
    if (openVal != "")
        rows.Push({ kind: "card", name: card, label: "Open invoice", value: openVal })
    if (partialVal != "")
        rows.Push({ kind: "card", name: card, label: "Partial value", value: partialVal })
    return { ok: true, rows: rows, error: "" }
}

MobillsDaily_WriteBalancesCsv(rows) {
    desktop := MobillsDaily_DesktopPath()
    name := "mobills-balances-" FormatTime(A_Now, "yyyy-MM-dd-HHmmss") ".csv"
    path := desktop "\" name
    body := "Kind,Name,Label,Value`n"
    for r in rows {
        body .= r.kind "," '"' StrReplace(r.name, '"', '""') '"' "," '"' StrReplace(r.label, '"', '""') '"' "," '"' StrReplace(
            r.value, '"', '""') '"`n'
    }
    try FileAppend(body, path, "UTF-8")
    catch as e {
        return { ok: false, path: "", error: e.Message }
    }
    return { ok: true, path: path, error: "" }
}

MobillsDaily_ShowReviewTable(rows, csvPath) {
    global g_MobillsDailyReviewResult, g_MobillsDailyReviewGui
    g_MobillsDailyReviewResult := ""
    dlg := Gui("+AlwaysOnTop +ToolWindow", "Mobills balances")
    dlg.SetFont("s10", "Segoe UI")
    dlg.Add("Text", "w720", "Saved: " csvPath)
    lv := dlg.Add("ListView", "w720 h360 -Multi", ["Kind", "Name", "Label", "Value"])
    for r in rows
        lv.Add("", r.kind, r.name, r.label, r.value)
    try lv.ModifyCol(1, 80)
    try lv.ModifyCol(2, 220)
    try lv.ModifyCol(3, 160)
    try lv.ModifyCol(4, 140)
    dlg.Add("Button", "w100 Default", "Close").OnEvent("Click", MobillsDaily_ReviewClose)
    dlg.OnEvent("Close", MobillsDaily_ReviewClose)
    dlg.OnEvent("Escape", MobillsDaily_ReviewClose)
    g_MobillsDailyReviewGui := dlg
    dlg.Show()
    start := A_TickCount
    while (g_MobillsDailyReviewResult = "") {
        if ((A_TickCount - start) >= 300000)
            break
        Sleep 50
    }
    g_MobillsDailyReviewResult := ""
    try dlg.Destroy()
    catch {
    }
    g_MobillsDailyReviewGui := ""
}

MobillsDaily_ReviewClose(*) {
    global g_MobillsDailyReviewResult
    g_MobillsDailyReviewResult := "done"
}

MobillsDaily_Run() {
    path := MobillsDaily_FindLatestFile()
    if (path = "") {
        ShowCenteredOverlay_Utils("❌ No MOBILLS_V1 file on Desktop", 2500, BANNER_ACCENT_ERROR)
        return
    }
    parsed := MobillsDaily_ParseFile(path)
    if (parsed.error != "") {
        ShowCenteredOverlay_Utils("❌ " parsed.error, 2500, BANNER_ACCENT_ERROR)
        return
    }
    if (!parsed.rows.Length) {
        ShowCenteredOverlay_Utils("❌ No transactions in " path, 2500, BANNER_ACCENT_ERROR)
        return
    }
    if !MobillsDaily_ShowConfirmTable(parsed.rows, parsed.unmapped, path) {
        ShowCenteredOverlay_Utils("Mobills daily entry cancelled", 1500, BANNER_ACCENT_INTERMEDIATE)
        return
    }

    StandardLoadingBar_Show("⏳ Opening Mobills...", BANNER_ACCENT_INTERMEDIATE, { fontSize: 17, trackActiveMonitor: true })
    barOwned := true
    try {
        att := MobillsAuto_AttachBrowser()
        if (att.error != "") {
            StandardLoadingBar_Hide(0)
            barOwned := false
            MobillsDaily_Halt(att.error)
            return
        }
        uia := att.uia
        total := parsed.rows.Length
        i := 1
        for row in parsed.rows {
            if !MobillsDaily_EnterRow(uia, row, i, total) {
                StandardLoadingBar_Hide(0)
                barOwned := false
                return
            }
            i++
            try uia := MobillsAuto_AttachBrowser().uia
            catch {
            }
        }

        StandardLoadingBar_Update("⏳ Scraping accounts...")
        acc := MobillsDaily_ScrapeAccounts(uia)
        if (!acc.ok) {
            StandardLoadingBar_Hide(0)
            barOwned := false
            MobillsDaily_Halt("Accounts scrape failed: " acc.error)
            return
        }
        try uia := MobillsAuto_AttachBrowser().uia
        catch {
        }
        cards := MobillsDaily_ScrapeCards(uia)
        if (!cards.ok) {
            StandardLoadingBar_Hide(0)
            barOwned := false
            MobillsDaily_Halt("Credit cards scrape failed: " cards.error)
            return
        }
        scraped := []
        for r in acc.rows
            scraped.Push(r)
        for r in cards.rows
            scraped.Push(r)
        written := MobillsDaily_WriteBalancesCsv(scraped)
        StandardLoadingBar_Hide(0)
        barOwned := false
        if (!written.ok) {
            MobillsDaily_Halt("Could not write balances CSV: " written.error)
            return
        }
        ShowCenteredOverlay_Utils("✅ Mobills daily entry done", 1500, BANNER_ACCENT_SUCCESS)
        MobillsDaily_ShowReviewTable(scraped, written.path)
    } catch as e {
        try StandardLoadingBar_Hide(0)
        barOwned := false
        MobillsDaily_Halt("Unexpected error: " e.Message)
    } finally {
        if (barOwned) {
            try StandardLoadingBar_Hide(0)
        }
    }
}
