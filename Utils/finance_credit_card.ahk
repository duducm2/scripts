; =============================================================================
; Utils module: finance_credit_card.ahk
; Primary credit card, mark as paid
; =============================================================================

Finance_ShowCreditCard() {
    global g_FinanceGui
    Finance_CloseGui()
    Finance_EnsureData()
    cards := Finance_Load("credit_cards")
    accs := Finance_Load("accounts")
    primaryId := Finance_Setting("General", "PrimaryCardId", "CARD_MP")
    card := Finance_FindById(cards, primaryId)
    if (!card && cards.Length)
        card := cards[1]
    g_FinanceGui := Gui("+AlwaysOnTop +ToolWindow", "Credit card")
    g_FinanceGui.SetFont("s10", "Segoe UI")
    if (!card) {
        g_FinanceGui.Add("Text", "w480", "No credit card yet.")
        g_FinanceGui.Add("Button", "y+12 w140", "Add card").OnEvent("Click", (*) => Finance_CardForm(false))
        g_FinanceGui.Add("Button", "x+8 w140", "Back").OnEvent("Click", (*) => Finance_ShowMainMenu())
        g_FinanceGui.OnEvent("Escape", (*) => Finance_ShowMainMenu())
        Finance_CenterGui(g_FinanceGui, 520, 180)
        return
    }
    lim := Finance_ParseDecimal(card["limit"])
    spent := Finance_ParseDecimal(card["current_spent"])
    avail := lim - spent
    pct := lim > 0 ? Round(spent / lim * 100, 2) : 0
    accName := Finance_AccName(accs, card["linked_account_id"])
    g_FinanceGui.SetFont("s16 Bold", "Segoe UI")
    g_FinanceGui.Add("Text", "x20 y16 w520", card["name"])
    g_FinanceGui.SetFont("s11 Norm", "Segoe UI")
    g_FinanceGui.Add("Text", "x20 y56 w520", "Id  " . card["id"])
    g_FinanceGui.Add("Text", "x20 y84 w520", "Limit  " . Finance_FormatBrl(lim))
    g_FinanceGui.Add("Text", "x20 y112 w520", "Current spent  " . Finance_FormatBrl(spent))
    g_FinanceGui.Add("Text", "x20 y140 w520", "Available  " . Finance_FormatBrl(avail) . "  (" . pct . "%)")
    g_FinanceGui.Add("Text", "x20 y168 w520", "Linked account  " . accName)
    g_FinanceGui.Add("Text", "x20 y196 w520", "Closing day  " . card["closing_day"])
    g_FinanceGui.Add("Button", "x20 y240 w180 Default", "[P] Mark as paid").OnEvent("Click", (*) =>
        Finance_CardMarkPaid(card["id"]))
    g_FinanceGui.Add("Button", "x+12 w140", "[E] Edit").OnEvent("Click", (*) => Finance_CardForm(card))
    g_FinanceGui.Add("Button", "x+12 w140", "Back").OnEvent("Click", (*) => Finance_ShowMainMenu())
    g_FinanceGui.OnEvent("Close", (*) => Finance_CloseGui())
    g_FinanceGui.OnEvent("Escape", (*) => Finance_ShowMainMenu())
    Finance_BindHotkeys([
        ["p", (*) => Finance_CardMarkPaid(card["id"])],
        ["e", (*) => Finance_CardForm(card)],
        ["Backspace", (*) => Finance_ShowMainMenu()],
        ["Escape", (*) => Finance_ShowMainMenu()]
    ])
    Finance_CenterGui(g_FinanceGui, 560, 340)
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
        "category_id", Finance_CatIdByName("Operação bancária"),
        "subcategory", "",
        "account_id", acc["id"],
        "card_id", card["id"],
        "transfer_account_id", ""
    )
    if (tx["category_id"] = "")
        tx["category_id"] := Finance_CatIdByName("Outros")
    txs.Push(tx)
    Finance_Save("transactions", txs)
    Finance_AdjustAccount(accs, acc["id"], -spent)
    card["current_spent"] := "0,00"
    Finance_Save("accounts", accs)
    Finance_Save("credit_cards", cards)
    Finance_Notify("Invoice paid", 1600, BANNER_ACCENT_SUCCESS)
    Finance_ShowCreditCard()
}

Finance_CardForm(existing) {
    global g_FinanceGui
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
    if (saved)
        Finance_ShowCreditCard()

    SaveCard(*) {
        name := Trim(eName.Value)
        if (name = "") {
            Finance_Alert("Name is required.", "Credit card")
            return
        }
        cards := Finance_Load("credit_cards")
        row := Map(
            "id", isEdit ? existing["id"] : Finance_SlugId("CARD_", name, cards),
        "name", name,
        "limit", Finance_FormatCsvDecimal(Finance_ParseDecimal(eLim.Value)),
        "current_spent", Finance_FormatCsvDecimal(Finance_ParseDecimal(eSpent.Value)),
        "linked_account_id", accCombo.ids[ddAcc.Value],
        "closing_day", Integer(eClose.Value || 1))
        if (isEdit) {
            out := []
            for r in cards {
                if (r["id"] = existing["id"])
                    out.Push(row)
                else
                    out.Push(r)
            }
            cards := out
        } else {
            cards.Push(row)
        }
        Finance_Save("credit_cards", cards)
        Finance_SetSetting("General", "PrimaryCardId", row["id"])
        saved := true
        g.Destroy()
    }
}
