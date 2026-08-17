; =============================================================================
; Utils module: finance_recurring.ahk
; Recurring bills CRUD (tracking/visualization only; not posted as transactions)
; =============================================================================

global g_FinanceRecLv := false
global g_FinanceRecRows := []

Finance_ShowRecurring() {
    global g_FinanceGui, g_FinanceRecLv
    Finance_CloseGui()
    Finance_EnsureData()
    g_FinanceGui := Gui("+AlwaysOnTop +ToolWindow", "Recurring bills")
    g_FinanceGui.SetFont("s10", "Segoe UI")
    g_FinanceGui.Add("Text", "x12 y10 w860",
        "[A]/Insert add   [E] edit   Delete   Backspace menu   Tracking only (no auto-charge)")
    g_FinanceRecLv := g_FinanceGui.Add("ListView", "x12 y40 w860 h480 Grid",
        ["Icon", "Name", "Monthly"])
    g_FinanceRecLv.OnEvent("DoubleClick", (*) => Finance_RecEdit())
    g_FinanceGui.OnEvent("Close", (*) => Finance_CloseGui())
    g_FinanceGui.OnEvent("Escape", (*) => Finance_ShowMainMenu())
    Finance_RecRefresh()
    Finance_BindHotkeys([
        ["a", (*) => Finance_RecAdd()],
        ["Insert", (*) => Finance_RecAdd()],
        ["e", (*) => Finance_RecEdit()],
        ["Delete", (*) => Finance_RecDelete()],
        ["Backspace", (*) => Finance_ShowMainMenu()],
        ["Escape", (*) => Finance_ShowMainMenu()]
    ])
    Finance_CenterGui(g_FinanceGui, 890, 560)
}

Finance_RecRefresh() {
    global g_FinanceRecLv, g_FinanceRecRows
    if (!IsObject(g_FinanceRecLv))
        return
    rows := Finance_Load("recurring_bills")
    g_FinanceRecLv.Delete()
    g_FinanceRecRows := []
    for r in rows {
        g_FinanceRecRows.Push(r)
        amt := Finance_ParseDecimal(r["monthly_amount"])
        icon := r.Has("icon") ? r["icon"] : "🔁"
        g_FinanceRecLv.Add("", icon, r["name"], Finance_FormatBrl(amt))
    }
    loop 3
        g_FinanceRecLv.ModifyCol(A_Index, "AutoHdr")
}

Finance_RecSelected() {
    global g_FinanceRecLv, g_FinanceRecRows
    row := g_FinanceRecLv.GetNext()
    if (!row || row > g_FinanceRecRows.Length)
        return false
    return g_FinanceRecRows[row]
}

Finance_RecAdd(*) {
    Finance_RecForm(false)
}

Finance_RecEdit(*) {
    r := Finance_RecSelected()
    if (!r) {
        Finance_Notify("Select a bill", 1200, BANNER_ACCENT_ERROR)
        return
    }
    Finance_RecForm(r)
}

Finance_RecDelete(*) {
    sel := Finance_RecSelected()
    if (!sel)
        return
    if (!Finance_Confirm("Delete " . sel["name"] . "?", "Recurring bills"))
        return
    out := []
    for r in Finance_Load("recurring_bills") {
        if (r["id"] != sel["id"])
            out.Push(r)
    }
    Finance_Save("recurring_bills", out)
    Finance_RecRefresh()
}

Finance_RecForm(existing) {
    global g_FinanceGui
    isEdit := IsObject(existing)
    owner := ""
    try {
        if (IsObject(g_FinanceGui))
            owner := " +Owner" . g_FinanceGui.Hwnd
    } catch {
        owner := ""
    }
    Finance_DialogsBegin()
    g := Gui("+AlwaysOnTop +ToolWindow" . owner, isEdit ? "Edit recurring bill" : "Add recurring bill")
    g.SetFont("s10", "Segoe UI")
    g.Add("Text", , "Name")
    eName := g.Add("Edit", "w320", isEdit ? existing["name"] : "")
    g.Add("Text", "y+8", "Icon (emoji)")
    eIcon := g.Add("Edit", "w160", isEdit ? existing["icon"] : "🔁")
    g.Add("Text", "y+8", "Monthly amount")
    eAmt := g.Add("Edit", "w160", isEdit ? existing["monthly_amount"] : "0,00")
    g.SetFont("s9 c555555", "Segoe UI")
    g.Add("Text", "y+8 w320", "Tracking only. Does not create transactions or change balances.")
    g.SetFont("s10", "Segoe UI")
    saved := false
    g.Add("Button", "y+16 w100 Default", "Save").OnEvent("Click", SaveRec)
    g.Add("Button", "x+8 w100", "Cancel").OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())
    g.Show()
    try WinWaitClose("ahk_id " g.Hwnd)
    catch {
    }
    Finance_DialogsEnd()
    if (saved)
        Finance_RecRefresh()

    SaveRec(*) {
        name := Trim(eName.Value)
        if (name = "") {
            Finance_Alert("Name is required.", "Recurring bills")
            return
        }
        icon := Trim(eIcon.Value)
        if (icon = "")
            icon := "🔁"
        amt := Finance_FormatCsvDecimal(Finance_ParseDecimal(eAmt.Value))
        rows := Finance_Load("recurring_bills")
        row := Map(
            "id", isEdit ? existing["id"] : Finance_SlugId("BILL_", name, rows),
        "name", name,
        "icon", icon,
        "monthly_amount", amt)
        if (isEdit) {
            out := []
            for r in rows {
                if (r["id"] = existing["id"])
                    out.Push(row)
                else
                    out.Push(r)
            }
            rows := out
        } else {
            rows.Push(row)
        }
        Finance_Save("recurring_bills", rows)
        saved := true
        g.Destroy()
    }
}
