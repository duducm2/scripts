; =============================================================================
; Utils module: finance_credit_card.ahk
; Credit cards ListView CRUD, mark paid, primary selection, usage chart
; =============================================================================

global g_FinanceCardLv := false
global g_FinanceCardRows := []
global g_FinanceCardHeader := false
global g_FinanceCardChartLabel := false
global g_FinanceCardChartBar := false

Finance_ShowCreditCard() {
    global g_FinanceGui, g_FinanceCardLv, g_FinanceCardHeader, g_FinanceCardChartLabel, g_FinanceCardChartBar
    Finance_CloseGui()
    Finance_EnsureData()
    g_FinanceGui := Gui("+AlwaysOnTop +ToolWindow", "Credit cards")
    g_FinanceGui.SetFont("s10", "Segoe UI")
    g_FinanceCardHeader := g_FinanceGui.Add("Text", "x12 y10 w860 h24")
    g_FinanceGui.Add("Text", "x12 y36 w860",
        "[A]/Insert add   [E] edit   Delete   [P] pay   [R] primary   Backspace menu")
    g_FinanceCardChartLabel := g_FinanceGui.Add("Text", "x12 y64 w860 h36")
    g_FinanceCardChartBar := g_FinanceGui.Add("Progress", "x12 y104 w860 h18 c2ECC71 Background333333 Range0-100", 0)
    g_FinanceCardLv := g_FinanceGui.Add("ListView", "x12 y132 w860 h390 Grid",
        ["Primary", "Name", "Limit", "Spent", "Available", "%", "Linked", "Close day"])
    g_FinanceCardLv.OnEvent("DoubleClick", (*) => Finance_CardEdit())
    g_FinanceCardLv.OnEvent("ItemSelect", (*) => Finance_CardUpdateChart())
    g_FinanceGui.OnEvent("Close", (*) => Finance_CloseGui())
    g_FinanceGui.OnEvent("Escape", (*) => Finance_ShowMainMenu())
    Finance_CardRefresh()
    Finance_BindHotkeys([
        ["a", (*) => Finance_CardAdd()],
        ["Insert", (*) => Finance_CardAdd()],
        ["e", (*) => Finance_CardEdit()],
        ["Delete", (*) => Finance_CardDelete()],
        ["p", (*) => Finance_CardPaySelected()],
        ["r", (*) => Finance_CardSetPrimary()],
        ["Backspace", (*) => Finance_ShowMainMenu()],
        ["Escape", (*) => Finance_ShowMainMenu()]
    ])
    Finance_CenterGui(g_FinanceGui, 890, 560)
}

Finance_CardRefresh() {
    global g_FinanceCardLv, g_FinanceCardRows, g_FinanceCardHeader
    if (!IsObject(g_FinanceCardLv))
        return
    cards := Finance_Load("credit_cards")
    accs := Finance_Load("accounts")
    primaryId := Finance_Setting("General", "PrimaryCardId", "")
    g_FinanceCardLv.Delete()
    g_FinanceCardRows := []
    totLim := 0.0
    totSpent := 0.0
    selectRow := 0
    for c in cards {
        g_FinanceCardRows.Push(c)
        lim := Finance_ParseDecimal(c["limit"])
        spent := Finance_ParseDecimal(c["current_spent"])
        avail := lim - spent
        pct := lim > 0 ? Round(spent / lim * 100, 1) : 0
        totLim += lim
        totSpent += spent
        star := (c["id"] = primaryId) ? "*" : ""
        g_FinanceCardLv.Add("", star, c["name"], Finance_FormatBrl(lim), Finance_FormatBrl(spent),
        Finance_FormatBrl(avail), pct . "%", Finance_AccName(accs, c["linked_account_id"]), c["closing_day"])
        if (primaryId != "" && c["id"] = primaryId)
            selectRow := g_FinanceCardRows.Length
    }
    if (!selectRow && g_FinanceCardRows.Length)
        selectRow := 1
    totAvail := totLim - totSpent
    g_FinanceCardHeader.Value := "Limit " . Finance_FormatBrl(totLim)
    . "  ·  Spent " . Finance_FormatBrl(totSpent)
    . "  ·  Available " . Finance_FormatBrl(totAvail)
    loop 8
        g_FinanceCardLv.ModifyCol(A_Index, "AutoHdr")
    if (selectRow) {
        g_FinanceCardLv.Modify(selectRow, "Select Focus Vis")
    }
    Finance_CardUpdateChart()
}

Finance_CardUpdateChart(*) {
    global g_FinanceCardChartLabel, g_FinanceCardChartBar
    if (!IsObject(g_FinanceCardChartLabel) || !IsObject(g_FinanceCardChartBar))
        return
    c := Finance_CardSelected()
    if (!c) {
        g_FinanceCardChartLabel.Value := "Select a credit card to see limit usage"
        g_FinanceCardChartBar.Value := 0
        try g_FinanceCardChartBar.Opt("c2ECC71")
        catch {
        }
        return
    }
    lim := Finance_ParseDecimal(c["limit"])
    spent := Finance_ParseDecimal(c["current_spent"])
    avail := lim - spent
    pct := lim > 0 ? Round(spent / lim * 100, 1) : 0
    barPct := lim > 0 ? Min(100, Round(spent / lim * 100)) : 0
    warn := Finance_ParseDecimal(Finance_Setting("General", "CardUsageWarnPct", "80"))
    g_FinanceCardChartLabel.Value := c["name"] . "`nSpent " . Finance_FormatBrl(spent)
    . " / Limit " . Finance_FormatBrl(lim)
    . "  ·  Available " . Finance_FormatBrl(avail)
    . "  ·  " . pct . "%"
    g_FinanceCardChartBar.Value := barPct
    color := (pct >= warn || (lim > 0 && spent > lim)) ? "cE74C3C" : "c2ECC71"
    try g_FinanceCardChartBar.Opt(color)
    catch {
    }
}

Finance_CardSelected() {
    global g_FinanceCardLv, g_FinanceCardRows
    row := g_FinanceCardLv.GetNext()
    if (!row || row > g_FinanceCardRows.Length)
        return false
    return g_FinanceCardRows[row]
}

Finance_CardAdd(*) {
    Finance_CardForm(false)
}

Finance_CardEdit(*) {
    c := Finance_CardSelected()
    if (!c) {
        Finance_Notify("Select a credit card", 1200, BANNER_ACCENT_ERROR)
        return
    }
    Finance_CardForm(c)
}

Finance_CardDelete(*) {
    c := Finance_CardSelected()
    if (!c)
        return
    if (!Finance_Confirm("Delete " . c["name"] . "?", "Credit cards"))
        return
    cards := Finance_Load("credit_cards")
    out := []
    for r in cards {
        if (r["id"] != c["id"])
            out.Push(r)
    }
    Finance_Save("credit_cards", out)
    primaryId := Finance_Setting("General", "PrimaryCardId", "")
    if (primaryId = c["id"]) {
        if (out.Length)
            Finance_SetSetting("General", "PrimaryCardId", out[1]["id"])
        else
            Finance_SetSetting("General", "PrimaryCardId", "")
    }
    Finance_CardRefresh()
}

Finance_CardPaySelected(*) {
    c := Finance_CardSelected()
    if (!c) {
        Finance_Notify("Select a credit card", 1200, BANNER_ACCENT_ERROR)
        return
    }
    Finance_CardMarkPaid(c["id"])
}

Finance_CardSetPrimary(*) {
    c := Finance_CardSelected()
    if (!c) {
        Finance_Notify("Select a credit card", 1200, BANNER_ACCENT_ERROR)
        return
    }
    Finance_SetSetting("General", "PrimaryCardId", c["id"])
    Finance_Notify(c["name"] . " is primary", 1400, BANNER_ACCENT_SUCCESS)
    Finance_CardRefresh()
}

Finance_CardMarkPaid(cardId) {
    cards := Finance_Load("credit_cards")
    accs := Finance_Load("accounts")
    card := Finance_FindById(cards, cardId)
    if (!card)
        return
    spent := Finance_ParseDecimal(card["current_spent"])
    if (spent <= 0) {
        Finance_Notify("Nothing to pay", 1400, BANNER_ACCENT_INFO)
        return
    }
    acc := Finance_FindById(accs, card["linked_account_id"])
    if (!acc) {
        Finance_Notify("Linked account missing", 1800, BANNER_ACCENT_ERROR)
        return
    }
    msg := "Pay " . Finance_FormatBrl(spent) . " from " . acc["name"] . " and reset the card?"
    if (!Finance_Confirm(msg, "Mark as paid"))
        return
    txs := Finance_Load("transactions")
    tx := Map(
        "id", Finance_NextId("TX", txs),
        "date", Finance_Today(),
        "description", "Invoice payment — " . card["name"],
        "amount", Finance_FormatCsvDecimal(spent),
        "type", "expense",
        "category_id", Finance_CatIdByName("Banking"),
        "subcategory", "",
        "account_id", acc["id"],
        "card_id", card["id"],
        "transfer_account_id", ""
    )
    if (tx["category_id"] = "")
        tx["category_id"] := Finance_CatIdByName("Other")
    txs.Push(tx)
    Finance_Save("transactions", txs)
    Finance_AdjustAccount(accs, acc["id"], -spent)
    card["current_spent"] := "0,00"
    card["initial_spent"] := Finance_FormatCsvDecimal(0 - Finance_CardNetFromTransactions(card["id"]))
    Finance_Save("accounts", accs)
    Finance_Save("credit_cards", cards)
    Finance_Notify("Invoice paid", 1600, BANNER_ACCENT_SUCCESS)
    global g_FinanceCardLv
    if (IsObject(g_FinanceCardLv))
        Finance_CardRefresh()
    else
        Finance_ShowCreditCard()
}

Finance_CardForm(existing) {
    global g_FinanceGui, g_FinanceCardLv
    accs := Finance_Load("accounts")
    isEdit := IsObject(existing)
    owner := ""
    try {
        if (IsObject(g_FinanceGui))
            owner := " +Owner" . g_FinanceGui.Hwnd
    } catch {
        owner := ""
    }
    Finance_DialogsBegin()
    g := Gui("+AlwaysOnTop +ToolWindow" . owner, isEdit ? "Edit credit card" : "Add credit card")
    g.SetFont("s10", "Segoe UI")
    g.Add("Text", , "Name")
    eName := g.Add("Edit", "w280", isEdit ? existing["name"] : "")
    g.Add("Text", "y+8", "Limit")
    eLim := g.Add("Edit", "w160", isEdit ? existing["limit"] : "0,00")
    g.Add("Text", "y+8", "Current spent")
    eSpent := g.Add("Edit", "w160", isEdit ? existing["current_spent"] : "0,00")
    g.Add("Text", "y+8", "Closing day (1-31)")
    eClose := g.Add("Edit", "w80", isEdit ? existing["closing_day"] : "9")
    accCombo := Finance_ComboFromRows(accs)
    accIdx := Finance_ComboIndex(accCombo.ids, isEdit ? existing["linked_account_id"] : "")
    g.Add("Text", "y+8", "Linked account")
    ddAcc := g.Add("DropDownList", "w280 Choose" . accIdx, accCombo.names)
    saved := false
    g.Add("Button", "y+16 w100 Default", "Save").OnEvent("Click", SaveCard)
    g.Add("Button", "x+8 w100", "Cancel").OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())
    g.Show()
    try WinWaitClose("ahk_id " g.Hwnd)
    catch {
    }
    Finance_DialogsEnd()
    if (saved) {
        if (IsObject(g_FinanceCardLv))
            Finance_CardRefresh()
        else
            Finance_ShowCreditCard()
    }

    SaveCard(*) {
        name := Trim(eName.Value)
        if (name = "") {
            Finance_Alert("Name is required.", "Credit cards")
            return
        }
        cards := Finance_Load("credit_cards")
        wasEmpty := cards.Length = 0
        spentStr := Finance_FormatCsvDecimal(Finance_ParseDecimal(eSpent.Value))
        if (isEdit) {
            initSpent := Finance_FormatCsvDecimal(
                Finance_ParseDecimal(eSpent.Value) - Finance_CardNetFromTransactions(existing["id"]))
            row := Map(
                "id", existing["id"],
                "name", name,
                "limit", Finance_FormatCsvDecimal(Finance_ParseDecimal(eLim.Value)),
                "initial_spent", initSpent,
                "current_spent", spentStr,
                "linked_account_id", accCombo.ids[ddAcc.Value],
                "closing_day", Integer(eClose.Value || 1))
            out := []
            for r in cards {
                if (r["id"] = existing["id"])
                    out.Push(row)
                else
                    out.Push(r)
            }
            cards := out
        } else {
            row := Map(
                "id", Finance_SlugId("CARD_", name, cards),
                "name", name,
                "limit", Finance_FormatCsvDecimal(Finance_ParseDecimal(eLim.Value)),
                "initial_spent", spentStr,
                "current_spent", spentStr,
                "linked_account_id", accCombo.ids[ddAcc.Value],
                "closing_day", Integer(eClose.Value || 1))
            cards.Push(row)
        }
        Finance_Save("credit_cards", cards)
        if (!isEdit && wasEmpty)
            Finance_SetSetting("General", "PrimaryCardId", row["id"])
        saved := true
        g.Destroy()
    }
}
