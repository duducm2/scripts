; =============================================================================
; Utils module: finance_transactions.ahk
; Filterable transaction list, month nav, add/edit/delete
; =============================================================================

global g_FinanceTxLv := false
global g_FinanceTxFilter := "all"
global g_FinanceTxHeader := false
global g_FinanceTxHint := false
global g_FinanceTxRows := []

Finance_ShowTransactions() {
    global g_FinanceGui, g_FinanceMonth, g_FinanceTxLv, g_FinanceTxFilter, g_FinanceTxHeader, g_FinanceTxHint
    Finance_CloseGui()
    Finance_EnsureData()
    if (g_FinanceMonth = "")
        g_FinanceMonth := Finance_CurrentYearMonth()

    g_FinanceGui := Gui("+AlwaysOnTop +ToolWindow", "Transactions")
    g_FinanceGui.SetFont("s10", "Segoe UI")
    g_FinanceTxHeader := g_FinanceGui.Add("Text", "x12 y10 w880 h36")
    g_FinanceTxHint := g_FinanceGui.Add("Text", "x12 y48 w880",
        "[G] all  [X] expenses  [N] incomes  [F] transfers  [,] prev month  [.] next  [A]/Insert add  [E] edit  Delete  Backspace menu"
    )
    g_FinanceTxLv := g_FinanceGui.Add("ListView", "x12 y74 w896 h500 Grid",
        ["Amount", "Account", "Category", "Description", "Date", "Type"])
    g_FinanceTxLv.OnEvent("DoubleClick", (*) => Finance_TxEdit())
    g_FinanceGui.OnEvent("Close", (*) => Finance_CloseGui())
    g_FinanceGui.OnEvent("Escape", (*) => Finance_ShowMainMenu())
    Finance_TxRefresh()
    Finance_BindHotkeys([
        ["g", (*) => Finance_TxSetFilter("all")],
        ["x", (*) => Finance_TxSetFilter("expense")],
        ["n", (*) => Finance_TxSetFilter("income")],
        ["f", (*) => Finance_TxSetFilter("transfer")],
        ["vkBC", (*) => Finance_TxShiftMonth(-1)],
        ["vkBE", (*) => Finance_TxShiftMonth(1)],
        ["a", (*) => Finance_TxAdd()],
        ["Insert", (*) => Finance_TxAdd()],
        ["e", (*) => Finance_TxEdit()],
        ["Delete", (*) => Finance_TxDelete()],
        ["Backspace", (*) => Finance_ShowMainMenu()],
        ["Escape", (*) => Finance_ShowMainMenu()]
    ])
    Finance_CenterGui(g_FinanceGui, 920, 600)
}

Finance_TxSetFilter(f) {
    global g_FinanceTxFilter
    g_FinanceTxFilter := f
    Finance_TxRefresh()
}

Finance_TxShiftMonth(d) {
    global g_FinanceMonth
    g_FinanceMonth := Finance_ShiftMonth(g_FinanceMonth, d)
    Finance_TxRefresh()
}

Finance_TxMatchesFilter(tx, filt) {
    t := tx["type"]
    if (filt = "all")
        return true
    if (filt = "expense")
        return (t = "expense" || t = "card_expense")
    if (filt = "income")
        return (t = "income")
    if (filt = "transfer")
        return (t = "transfer")
    return true
}

Finance_TxRefresh() {
    global g_FinanceTxLv, g_FinanceTxFilter, g_FinanceMonth, g_FinanceTxHeader, g_FinanceTxRows
    if (!IsObject(g_FinanceTxLv))
        return
    accs := Finance_Load("accounts")
    cats := Finance_Load("categories")
    tot := Finance_MonthTotals(g_FinanceMonth)
    filtLabel := Map("all", "General", "expense", "Expenses", "income", "Incomes", "transfer", "Transfers")
    fl := filtLabel.Has(g_FinanceTxFilter) ? filtLabel[g_FinanceTxFilter] : g_FinanceTxFilter
    g_FinanceTxHeader.Value := Finance_MonthLabel(g_FinanceMonth) . "  ·  " . fl
    . "  ·  In " . Finance_FormatBrl(tot.income)
    . "  ·  Out " . Finance_FormatBrl(tot.expense)
    . "  ·  Balance " . Finance_FormatBrl(tot.balance)
    g_FinanceTxLv.Delete()
    g_FinanceTxRows := []
    txs := Finance_Load("transactions")
    ; newest first
    loop txs.Length {
        tx := txs[txs.Length - A_Index + 1]
        if (!Finance_TxInMonth(tx, g_FinanceMonth))
            continue
        if (!Finance_TxMatchesFilter(tx, g_FinanceTxFilter))
            continue
        g_FinanceTxRows.Push(tx)
        cat := Finance_CatName(cats, tx["category_id"])
        if (tx["subcategory"] != "")
            cat .= " / " . Finance_SubcatLabel(cats, tx["category_id"], tx["subcategory"])
        acc := Finance_AccName(accs, tx["account_id"])
        g_FinanceTxLv.Add("", Finance_FormatBrl(Finance_ParseDecimal(tx["amount"])), acc, cat,
        tx["description"], tx["date"], Finance_TypeLabel(tx["type"]))
    }
    g_FinanceTxLv.ModifyCol(1, 110)
    g_FinanceTxLv.ModifyCol(2, 160)
    g_FinanceTxLv.ModifyCol(3, 170)
    g_FinanceTxLv.ModifyCol(4, 220)
    g_FinanceTxLv.ModifyCol(5, 100)
    g_FinanceTxLv.ModifyCol(6, 100)
    if (g_FinanceTxRows.Length = 0)
        g_FinanceTxHeader.Value .= "  ·  No transactions this month"
}

Finance_TxSelected() {
    global g_FinanceTxLv, g_FinanceTxRows
    row := g_FinanceTxLv.GetNext()
    if (!row || row > g_FinanceTxRows.Length)
        return false
    return g_FinanceTxRows[row]
}

Finance_TxAdd(*) {
    Finance_TxForm(false)
}

Finance_TxEdit(*) {
    tx := Finance_TxSelected()
    if (!tx) {
        Finance_Notify("Select a transaction", 1200, BANNER_ACCENT_ERROR)
        return
    }
    Finance_TxForm(tx)
}

Finance_TxDelete(*) {
    tx := Finance_TxSelected()
    if (!tx)
        return
    if (!Finance_Confirm("Delete " . tx["description"] . "?", "Transactions"))
        return
    rows := Finance_Load("transactions")
    out := []
    for r in rows {
        if (r["id"] != tx["id"])
            out.Push(r)
    }
    Finance_ApplyTransactionToBalances(tx, true)
    Finance_Save("transactions", out)
    Finance_RecomputeBudgetSpent(SubStr(tx["date"], 1, 7))
    Finance_TxRefresh()
}

Finance_TxForm(existing) {
    global g_FinanceGui
    owner := IsObject(g_FinanceGui) ? " +Owner" . g_FinanceGui.Hwnd : ""
    accs := Finance_Load("accounts")
    cats := Finance_Load("categories")
    cards := Finance_Load("credit_cards")
    isEdit := IsObject(existing)
    Finance_DialogsBegin()
    g := Gui("+AlwaysOnTop +ToolWindow" . owner, isEdit ? "Edit transaction" : "Add transaction")
    g.SetFont("s10", "Segoe UI")
    g.Add("Text", "x10 y10", "Description")
    eDesc := g.Add("Edit", "x10 y28 w420", isEdit ? existing["description"] : "")
    g.Add("Text", "x10 y58", "Amount (comma decimal)")
    eAmt := g.Add("Edit", "x10 y76 w200", isEdit ? existing["amount"] : "")
    types := ["expense", "income", "transfer", "card_expense", "adjustment"]
    curType := isEdit ? existing["type"] : "expense"
    typeIdx := 1
    loop types.Length {
        if (types[A_Index] = curType)
            typeIdx := A_Index
    }
    g.Add("Text", "x230 y58", "Type")
    ddType := g.Add("DropDownList", "x230 y76 w200 Choose" . typeIdx, ["Expense", "Income", "Transfer", "Credit card",
        "Adjustment"])

    y1 := 112
    y1c := 130
    y2 := 168
    y2c := 186
    catCombo := Finance_ComboFromRows(Finance_MainCategories(cats), "id", "name", true, "icon")
    catIdx := Finance_ComboIndex(catCombo.ids, isEdit ? existing["category_id"] : "")
    lblCat := g.Add("Text", "x10 y" . y1, "Category")
    ddCat := g.Add("DropDownList", "x10 y" . y1c . " w220 Choose" . catIdx, catCombo.names)
    lblSub := g.Add("Text", "x242 y" . y1, "Subcategory")
    eSub := g.Add("Edit", "x242 y" . y1c . " w180", isEdit ? existing["subcategory"] : "")

    accCombo := Finance_ComboFromRows(accs)
    accIdx := Finance_ComboIndex(accCombo.ids, isEdit ? existing["account_id"] : Finance_Setting("General",
        "DefaultAccountId", ""))
    lblAcc := g.Add("Text", "x10 y" . y2, "Account")
    ddAcc := g.Add("DropDownList", "x10 y" . y2c . " w220 Choose" . accIdx, accCombo.names)

    destCombo := Finance_ComboFromRows(accs, "id", "name", true)
    destIdx := Finance_ComboIndex(destCombo.ids, isEdit ? existing["transfer_account_id"] : "")
    lblDest := g.Add("Text", "x242 y" . y2, "To account")
    ddDest := g.Add("DropDownList", "x242 y" . y2c . " w180 Choose" . destIdx, destCombo.names)

    cardCombo := Finance_ComboFromRows(cards, "id", "name", true)
    cardIdx := Finance_ComboIndex(cardCombo.ids, isEdit ? existing["card_id"] : "")
    lblCard := g.Add("Text", "x10 y" . y2, "Credit card")
    ddCard := g.Add("DropDownList", "x10 y" . y2c . " w220 Choose" . cardIdx, cardCombo.names)

    g.Add("Text", "x10 y230 w420", "Date is always the system entry date (today) for new rows.")
    saved := false
    g.Add("Button", "x10 y256 w100 Default", "Save").OnEvent("Click", SaveTx)
    g.Add("Button", "x118 y256 w100", "Cancel").OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())
    ddType.OnEvent("Change", (*) => ApplyTxTypeFields(types[ddType.Value]))
    ApplyTxTypeFields(curType)
    g.Show("w460 h300")
    try WinWaitClose("ahk_id " g.Hwnd)
    catch {
    }
    Finance_DialogsEnd()
    if (saved)
        Finance_TxRefresh()

    ApplyTxTypeFields(t) {
        showCat := (t != "transfer")
        showAcc := (t != "card_expense")
        showDest := (t = "transfer")
        showCard := (t = "card_expense")
        lblCat.Visible := showCat
        ddCat.Visible := showCat
        lblSub.Visible := showCat
        eSub.Visible := showCat
        lblAcc.Visible := showAcc
        ddAcc.Visible := showAcc
        lblDest.Visible := showDest
        ddDest.Visible := showDest
        lblCard.Visible := showCard
        ddCard.Visible := showCard
        lblAcc.Text := (t = "transfer") ? "From account" : "Account"
        if (t = "transfer") {
            lblAcc.Move(10, y1)
            ddAcc.Move(10, y1c)
            lblDest.Move(242, y2)
            ddDest.Move(242, y2c)
        } else {
            lblAcc.Move(10, y2)
            ddAcc.Move(10, y2c)
        }
        catFilter := ""
        if (t = "expense" || t = "card_expense")
            catFilter := "expense"
        else if (t = "income")
            catFilter := "income"
        keepId := ""
        try keepId := catCombo.ids[ddCat.Value]
        catch {
            keepId := ""
        }
        catCombo := Finance_ComboFromRows(Finance_MainCategories(cats, catFilter), "id", "name", true, "icon")
        ddCat.Delete()
        ddCat.Add(catCombo.names)
        ddCat.Choose(Finance_ComboIndex(catCombo.ids, keepId))
    }

    SaveTx(*) {
        desc := Trim(eDesc.Value)
        if (desc = "") {
            Finance_Alert("Description is required.", "Transactions")
            return
        }
        rawAmt := Finance_NormalizeDot(eAmt.Value)
        amt := Finance_ParseDecimal(rawAmt)
        if (amt < 0)
            amt := -amt
        t := types[ddType.Value]
        catId := ""
        subVal := ""
        accId := ""
        destId := ""
        cardId := ""
        if (t != "transfer") {
            catId := catCombo.ids[ddCat.Value]
            subVal := Trim(eSub.Value)
        }
        if (t != "card_expense")
            accId := accCombo.ids[ddAcc.Value]
        if (t = "transfer")
            destId := destCombo.ids[ddDest.Value]
        if (t = "card_expense") {
            cardId := cardCombo.ids[ddCard.Value]
            if (cardId = "")
                cardId := Finance_Setting("General", "PrimaryCardId", "CARD_MP")
        }
        date := isEdit ? existing["date"] : Finance_Today()
        txs := Finance_Load("transactions")
        newTx := Map(
            "id", isEdit ? existing["id"] : Finance_NextId("TX", txs),
        "date", date,
        "description", desc,
        "amount", Finance_FormatCsvDecimal(amt),
        "type", t,
        "category_id", catId,
        "subcategory", subVal,
        "account_id", accId,
        "card_id", cardId,
        "transfer_account_id", destId)
        if (isEdit) {
            out := []
            for r in txs {
                if (r["id"] = existing["id"])
                    out.Push(newTx)
                else
                    out.Push(r)
            }
            txs := out
            Finance_ReplaceTransaction(existing, newTx)
        } else {
            txs.Push(newTx)
            Finance_ApplyTransactionToBalances(newTx, false)
        }
        Finance_Save("transactions", txs)
        ymNew := SubStr(date, 1, 7)
        Finance_RecomputeBudgetSpent(ymNew)
        if (isEdit) {
            ymOld := SubStr(existing["date"], 1, 7)
            if (ymOld != "" && ymOld != ymNew)
                Finance_RecomputeBudgetSpent(ymOld)
        }
        saved := true
        g.Destroy()
    }
}
