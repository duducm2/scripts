; =============================================================================
; Utils module: finance_budgets.ahk
; Monthly budgets per category, auto-copy, spent vs remaining
; =============================================================================

global g_FinanceBudLv := false
global g_FinanceBudRows := []
global g_FinanceBudHeader := false

Finance_ShowBudgets() {
    global g_FinanceGui, g_FinanceMonth, g_FinanceBudLv, g_FinanceBudHeader
    Finance_CloseGui()
    Finance_EnsureData()
    if (g_FinanceMonth = "")
        g_FinanceMonth := Finance_CurrentYearMonth()
    Finance_EnsureMonthBudgets(g_FinanceMonth)
    g_FinanceGui := Gui("+AlwaysOnTop +ToolWindow", "Budgets")
    g_FinanceGui.SetFont("s10", "Segoe UI")
    g_FinanceBudHeader := g_FinanceGui.Add("Text", "x12 y8 w860 h52")
    g_FinanceGui.Add("Text", "x12 y64 w860",
        "[,] prev  [.] next  [Shift+A]/Insert add  [Shift+E] edit planned  Delete  Backspace menu")
    g_FinanceBudLv := g_FinanceGui.Add("ListView", "x12 y90 w860 h430 Grid",
        ["Category", "Planned", "Spent", "Remaining", "Status"])
    g_FinanceBudLv.OnEvent("DoubleClick", (*) => Finance_BudEdit())
    g_FinanceGui.OnEvent("Close", (*) => Finance_CloseGui())
    g_FinanceGui.OnEvent("Escape", (*) => Finance_ShowMainMenu())
    Finance_BudRefresh()
    Finance_BindHotkeys([
        ["vkBC", (*) => Finance_BudShift(-1)],
        ["vkBE", (*) => Finance_BudShift(1)],
        ["+a", (*) => Finance_BudAdd()],
        ["Insert", (*) => Finance_BudAdd()],
        ["+e", (*) => Finance_BudEdit()],
        ["Delete", (*) => Finance_BudDelete()],
        ["Backspace", (*) => Finance_ShowMainMenu()],
        ["Escape", (*) => Finance_ShowMainMenu()]
    ])
    Finance_CenterGui(g_FinanceGui, 890, 560)
}

Finance_BudShift(d) {
    global g_FinanceMonth
    g_FinanceMonth := Finance_ShiftMonth(g_FinanceMonth, d)
    Finance_EnsureMonthBudgets(g_FinanceMonth)
    Finance_BudRefresh()
}

Finance_BudRefresh() {
    global g_FinanceBudLv, g_FinanceBudRows, g_FinanceBudHeader, g_FinanceMonth
    Finance_RecomputeBudgetSpent(g_FinanceMonth)
    cats := Finance_Load("categories")
    tot := Finance_MonthTotals(g_FinanceMonth)
    planned := 0.0
    spentTot := 0.0
    g_FinanceBudLv.Delete()
    g_FinanceBudRows := []
    for b in Finance_Load("budgets") {
        if (b["year_month"] != g_FinanceMonth)
            continue
        g_FinanceBudRows.Push(b)
        p := Finance_ParseDecimal(b["planned_amount"])
        s := Finance_ParseDecimal(b["spent_amount"])
        planned += p
        spentTot += s
        rem := p - s
        st := rem >= 0 ? "Remain " . Finance_FormatBrl(rem) : "Exceeded " . Finance_FormatBrl(-rem)
        g_FinanceBudLv.Add("", Finance_CatName(cats, b["category_id"]), Finance_FormatBrl(p),
        Finance_FormatBrl(s), Finance_FormatBrl(rem), st)
    }
    plannedBal := tot.income - planned
    savPct := tot.income > 0 ? Round(plannedBal / tot.income * 100, 2) : 0
    g_FinanceBudHeader.Value := Finance_MonthLabel(g_FinanceMonth)
    . "`nIncomes " . Finance_FormatBrl(tot.income)
    . "  ·  Planned expenses " . Finance_FormatBrl(planned)
    . "  ·  Planned balance " . Finance_FormatBrl(plannedBal)
    . "  ·  Planned savings " . savPct . "%"
    . "  ·  Spent " . Finance_FormatBrl(spentTot)
    loop 5
        g_FinanceBudLv.ModifyCol(A_Index, "AutoHdr")
}

Finance_BudSelected() {
    global g_FinanceBudLv, g_FinanceBudRows
    row := g_FinanceBudLv.GetNext()
    if (!row || row > g_FinanceBudRows.Length)
        return false
    return g_FinanceBudRows[row]
}

Finance_BudAdd(*) {
    global g_FinanceMonth
    cats := Finance_MainCategories(Finance_Load("categories"), "expense")
    names := []
    ids := []
    existing := Map()
    for b in Finance_Load("budgets") {
        if (b["year_month"] = g_FinanceMonth)
            existing[b["category_id"]] := true
    }
    for c in cats {
        if (existing.Has(c["id"]))
            continue
        names.Push(Finance_CatLabel(c))
        ids.Push(c["id"])
    }
    if (!names.Length) {
        Finance_Notify("All expense categories already have a budget", 1800, BANNER_ACCENT_INFO)
        return
    }
    idx := Finance_PickList("Category", names)
    if (idx < 1)
        return
    ib := Finance_InputBox("Planned amount (comma decimal)", "Budget", "0,00")
    if (ib.Result != "OK")
        return
    rows := Finance_Load("budgets")
    rows.Push(Map("year_month", g_FinanceMonth, "category_id", ids[idx],
        "planned_amount", Finance_FormatCsvDecimal(Finance_ParseDecimal(ib.Value)), "spent_amount", "0,00"))
    Finance_Save("budgets", rows)
    Finance_BudRefresh()
}

Finance_BudEdit(*) {
    b := Finance_BudSelected()
    if (!b)
        return
    ib := Finance_InputBox("Planned amount", "Edit budget", b["planned_amount"])
    if (ib.Result != "OK")
        return
    rows := Finance_Load("budgets")
    for r in rows {
        if (r["year_month"] = b["year_month"] && r["category_id"] = b["category_id"])
            r["planned_amount"] := Finance_FormatCsvDecimal(Finance_ParseDecimal(ib.Value))
    }
    Finance_Save("budgets", rows)
    Finance_BudRefresh()
}

Finance_BudDelete(*) {
    b := Finance_BudSelected()
    if (!b)
        return
    out := []
    for r in Finance_Load("budgets") {
        if (!(r["year_month"] = b["year_month"] && r["category_id"] = b["category_id"]))
            out.Push(r)
    }
    Finance_Save("budgets", out)
    Finance_BudRefresh()
}
