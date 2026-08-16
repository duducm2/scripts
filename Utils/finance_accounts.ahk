; =============================================================================
; Utils module: finance_accounts.ahk
; Account CRUD and balance adjustment
; =============================================================================

global g_FinanceAccLv := false
global g_FinanceAccRows := []
global g_FinanceAccHeader := false

Finance_ShowAccounts() {
    global g_FinanceGui, g_FinanceAccLv, g_FinanceAccHeader
    Finance_CloseGui()
    Finance_EnsureData()
    g_FinanceGui := Gui("+AlwaysOnTop +ToolWindow", "Accounts")
    g_FinanceGui.SetFont("s10", "Segoe UI")
    g_FinanceAccHeader := g_FinanceGui.Add("Text", "x12 y10 w860 h24")
    g_FinanceGui.Add("Text", "x12 y36 w860", "Insert add   F2 edit   Delete   J adjust balance   Backspace menu")
    g_FinanceAccLv := g_FinanceGui.Add("ListView", "x12 y64 w860 h460 Grid",
        ["Icon", "Name", "Id", "Current", "Initial"])
    g_FinanceAccLv.OnEvent("DoubleClick", (*) => Finance_AccEdit())
    g_FinanceGui.OnEvent("Close", (*) => Finance_CloseGui())
    g_FinanceGui.OnEvent("Escape", (*) => Finance_ShowMainMenu())
    Finance_AccRefresh()
    Finance_BindHotkeys([
        ["Insert", (*) => Finance_AccAdd()],
        ["F2", (*) => Finance_AccEdit()],
        ["Delete", (*) => Finance_AccDelete()],
        ["j", (*) => Finance_AccAdjust()],
        ["Backspace", (*) => Finance_ShowMainMenu()],
        ["Escape", (*) => Finance_ShowMainMenu()]
    ])
    Finance_CenterGui(g_FinanceGui, 890, 560)
}

Finance_AccRefresh() {
    global g_FinanceAccLv, g_FinanceAccRows, g_FinanceAccHeader
    accs := Finance_Load("accounts")
    g_FinanceAccLv.Delete()
    g_FinanceAccRows := []
    tot := 0.0
    for a in accs {
        g_FinanceAccRows.Push(a)
        cur := Finance_ParseDecimal(a["current_balance"])
        tot += cur
        g_FinanceAccLv.Add("", a["icon"], a["name"], a["id"], Finance_FormatBrl(cur),
        Finance_FormatBrl(Finance_ParseDecimal(a["initial_balance"])))
    }
    g_FinanceAccHeader.Value := "Current total  " . Finance_FormatBrl(tot)
    loop 5
        g_FinanceAccLv.ModifyCol(A_Index, "AutoHdr")
}

Finance_AccSelected() {
    global g_FinanceAccLv, g_FinanceAccRows
    row := g_FinanceAccLv.GetNext()
    if (!row || row > g_FinanceAccRows.Length)
        return false
    return g_FinanceAccRows[row]
}

Finance_AccAdd(*) {
    Finance_AccForm(false)
}

Finance_AccEdit(*) {
    a := Finance_AccSelected()
    if (!a) {
        Finance_Notify("Select an account", 1200, BANNER_ACCENT_ERROR)
        return
    }
    Finance_AccForm(a)
}

Finance_AccDelete(*) {
    a := Finance_AccSelected()
    if (!a)
        return
    if (MsgBox("Delete " . a["name"] . "?", "Accounts", "YesNo Icon?") != "Yes")
        return
    accs := Finance_Load("accounts")
    out := []
    for r in accs {
        if (r["id"] != a["id"])
            out.Push(r)
    }
    Finance_Save("accounts", out)
    Finance_AccRefresh()
}

Finance_AccForm(existing) {
    isEdit := IsObject(existing)
    g := Gui("+AlwaysOnTop +ToolWindow", isEdit ? "Edit account" : "Add account")
    g.SetFont("s10", "Segoe UI")
    g.Add("Text", , "Name")
    eName := g.Add("Edit", "w320", isEdit ? existing["name"] : "")
    g.Add("Text", "y+8", "Icon (emoji or file name)")
    eIcon := g.Add("Edit", "w160", isEdit ? existing["icon"] : "🏦")
    g.Add("Text", "y+8", "Initial balance")
    eInit := g.Add("Edit", "w160", isEdit ? existing["initial_balance"] : "0,00")
    saved := false
    g.Add("Button", "y+16 w100 Default", "Save").OnEvent("Click", SaveAcc)
    g.Add("Button", "x+8 w100", "Cancel").OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())
    g.Show()
    try WinWaitClose("ahk_id " g.Hwnd)
    catch {
    }
    if (saved)
        Finance_AccRefresh()

    SaveAcc(*) {
        name := Trim(eName.Value)
        if (name = "") {
            MsgBox("Name is required.", "Accounts", "Icon!")
            return
        }
        accs := Finance_Load("accounts")
        init := Finance_FormatCsvDecimal(Finance_ParseDecimal(eInit.Value))
        if (isEdit) {
            oldInit := Finance_ParseDecimal(existing["initial_balance"])
            newInit := Finance_ParseDecimal(init)
            delta := newInit - oldInit
            cur := Finance_ParseDecimal(existing["current_balance"]) + delta
            row := Map("id", existing["id"], "name", name, "icon", Trim(eIcon.Value),
            "initial_balance", init, "current_balance", Finance_FormatCsvDecimal(cur))
            out := []
            for r in accs {
                if (r["id"] = existing["id"])
                    out.Push(row)
                else
                    out.Push(r)
            }
            accs := out
        } else {
            accs.Push(Map("id", Finance_SlugId("ACC_", name, accs), "name", name, "icon", Trim(eIcon.Value),
            "initial_balance", init, "current_balance", init))
        }
        Finance_Save("accounts", accs)
        saved := true
        g.Destroy()
    }
}

Finance_AccAdjust(*) {
    a := Finance_AccSelected()
    if (!a) {
        Finance_Notify("Select an account", 1200, BANNER_ACCENT_ERROR)
        return
    }
    g := Gui("+AlwaysOnTop +ToolWindow", "Adjust balance — " . a["name"])
    g.SetFont("s10", "Segoe UI")
    g.Add("Text", "w400", "Current: " . Finance_FormatBrl(Finance_ParseDecimal(a["current_balance"])))
    g.Add("Text", "y+8", "New current balance")
    eNew := g.Add("Edit", "w180", a["current_balance"])
    g.Add("Text", "y+12 w400", "[1] Create an adjustment transaction`n[2] Modify the initial balance directly")
    g.Add("Button", "y+12 w200", "1 — New transaction").OnEvent("Click", (*) => DoAdj(1))
    g.Add("Button", "x+8 w200", "2 — Change initial").OnEvent("Click", (*) => DoAdj(2))
    g.OnEvent("Escape", (*) => g.Destroy())
    g.Show()
    try WinWaitClose("ahk_id " g.Hwnd)
    catch {
    }
    Finance_AccRefresh()

    DoAdj(mode) {
        newBal := Finance_ParseDecimal(eNew.Value)
        oldBal := Finance_ParseDecimal(a["current_balance"])
        delta := newBal - oldBal
        if (delta = 0) {
            g.Destroy()
            return
        }
        accs := Finance_Load("accounts")
        row := Finance_FindById(accs, a["id"])
        if (mode = 1) {
            txs := Finance_Load("transactions")
            tx := Map(
                "id", Finance_NextId("TX", txs),
                "date", Finance_Today(),
                "description", "Balance adjustment",
                "amount", Finance_FormatCsvDecimal(Abs(delta)),
                "type", "adjustment",
                "category_id", Finance_CatIdByName("Ajuste"),
                "subcategory", "",
                "account_id", a["id"],
                "card_id", "",
                "transfer_account_id", ""
            )
            if (delta < 0)
                tx["type"] := "expense"
            txs.Push(tx)
            Finance_Save("transactions", txs)
            row["current_balance"] := Finance_FormatCsvDecimal(newBal)
        } else {
            init := Finance_ParseDecimal(row["initial_balance"]) + delta
            row["initial_balance"] := Finance_FormatCsvDecimal(init)
            row["current_balance"] := Finance_FormatCsvDecimal(newBal)
        }
        Finance_Save("accounts", accs)
        g.Destroy()
    }
}
