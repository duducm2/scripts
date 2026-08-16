; =============================================================================
; Utils module: finance_launcher.ahk
; Finance app main menu (Utility Shortcuts [F])
; =============================================================================

global g_FinanceMenuLabels := []

Finance_LaunchApp() {
    Finance_EnsureData()
    Finance_ShowMainMenu()
}

Finance_ShowMainMenu() {
    global g_FinanceGui, g_FinanceMonth
    Finance_CloseGui()
    Finance_EnsureData()
    if (g_FinanceMonth = "")
        g_FinanceMonth := Finance_CurrentYearMonth()

    g_FinanceGui := Gui("+AlwaysOnTop +ToolWindow", "Finance")
    g_FinanceGui.SetFont("s10", "Segoe UI")
    g_FinanceGui.BackColor := "1E1E1E"
    g_FinanceGui.OnEvent("Close", (*) => Finance_CloseGui())
    g_FinanceGui.OnEvent("Escape", (*) => Finance_CloseGui())

    g_FinanceGui.SetFont("s16 cWhite Bold", "Segoe UI")
    g_FinanceGui.Add("Text", "x20 y16 w880", "Finance")
    g_FinanceGui.SetFont("s10 cC0C0C0 Norm", "Segoe UI")
    tot := Finance_MonthTotals(g_FinanceMonth)
    g_FinanceGui.Add("Text", "x20 y48 w880",
        Finance_FormatBrl(Finance_TotalBalance()) . "  ·  "
        . Finance_MonthLabel(g_FinanceMonth) . "  ·  In "
        . Finance_FormatBrl(tot.income) . "  ·  Out "
        . Finance_FormatBrl(tot.expense))

    notes := Finance_CollectNotifications()
    noteLine := notes.Length ? notes[1] : "No alerts"
    if (notes.Length > 1)
        noteLine .= "  (+" . (notes.Length - 1) . " more)"
    g_FinanceGui.SetFont("s9 cF1C40F", "Segoe UI")
    g_FinanceGui.Add("Text", "x20 y72 w880", noteLine)

    items := [
        ["D", "Dashboard", "Charts, widgets, notifications"],
        ["T", "Transactions", "List, filter, edit"],
        ["A", "Accounts", "Balances and CRUD"],
        ["C", "Credit card", "Limit, spent, mark paid"],
        ["B", "Budgets", "Monthly limits"],
        ["G", "Goals", "Funds and targets"],
        ["R", "Reports", "Category, cash flow"],
        ["K", "Categories", "CRUD, search, filter"],
        ["I", "AI import", "Desktop daily / monthly"],
        ["S", "Settings", "Dashboard widgets"]
    ]

    x0 := 20
    y0 := 108
    colW := 440
    rowH := 78
    idx := 0
    g_FinanceGui.SetFont("s11 cWhite", "Segoe UI")
    for it in items {
        col := Mod(idx, 2)
        row := idx // 2
        x := x0 + col * colW
        y := y0 + row * rowH
        box := g_FinanceGui.Add("Text", "x" . x . " y" . y . " w420 h70 Background2C2C2C cWhite",
            "")
        g_FinanceGui.SetFont("s14 cF1C40F Bold", "Segoe UI")
        g_FinanceGui.Add("Text", "x" . (x + 12) . " y" . (y + 10) . " w40 BackgroundTrans", "[" . it[1] . "]")
        g_FinanceGui.SetFont("s12 cWhite Bold", "Segoe UI")
        g_FinanceGui.Add("Text", "x" . (x + 58) . " y" . (y + 10) . " w340 BackgroundTrans", it[2])
        g_FinanceGui.SetFont("s9 cA0A0A0 Norm", "Segoe UI")
        g_FinanceGui.Add("Text", "x" . (x + 58) . " y" . (y + 36) . " w340 BackgroundTrans", it[3])
        idx += 1
    }

    g_FinanceGui.SetFont("s9 c808080", "Segoe UI")
    g_FinanceGui.Add("Text", "x20 y510 w880", "Esc close   letters open a module   S settings")

    Finance_BindHotkeys([
        ["d", Finance_OnDash], ["t", Finance_OnTx], ["a", Finance_OnAcc], ["c", Finance_OnCard],
        ["b", Finance_OnBud], ["g", Finance_OnGoals], ["r", Finance_OnRep], ["k", Finance_OnCat],
        ["i", Finance_OnImp], ["s", Finance_OnSet], ["Escape", (*) => Finance_CloseGui()]
    ])
    Finance_CenterGui(g_FinanceGui, 900, 560)
}

Finance_OnDash(*) {
    Finance_OpenDashboard(false)
}
Finance_OnRep(*) {
    Finance_OpenDashboard(true)
}
Finance_OnTx(*) {
    Finance_ShowTransactions()
}
Finance_OnAcc(*) {
    Finance_ShowAccounts()
}
Finance_OnCard(*) {
    Finance_ShowCreditCard()
}
Finance_OnBud(*) {
    Finance_ShowBudgets()
}
Finance_OnGoals(*) {
    Finance_ShowGoals()
}
Finance_OnCat(*) {
    Finance_ShowCategories()
}
Finance_OnImp(*) {
    Finance_ShowImportMenu()
}
Finance_OnSet(*) {
    Finance_ShowSettings()
}

Finance_OpenDashboard(reportsTab := false) {
    Finance_EnsureData()
    Finance_RecomputeBudgetSpent(Finance_CurrentYearMonth())
    py := Finance_PythonDir() . "\chart_generator.py"
    if (!FileExist(py)) {
        Finance_Notify("chart_generator.py not found", 2000, BANNER_ACCENT_ERROR)
        return
    }
    try StandardLoadingBar_Show("Building dashboard…", BANNER_ACCENT_INTERMEDIATE)
    catch {
    }
    cmd := 'python "' . py . '"'
    if (reportsTab)
        cmd .= " --reports"
    try {
        RunWait(cmd, A_ScriptDir, "Hide")
    } catch as e {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        Finance_Notify("Python failed: " . e.Message, 2500, BANNER_ACCENT_ERROR)
        return
    }
    html := Finance_OutputDir() . "\dashboard.html"
    try StandardLoadingBar_Hide(400)
    catch {
    }
    if (!FileExist(html)) {
        Finance_Notify("dashboard.html was not generated", 2200, BANNER_ACCENT_ERROR)
        return
    }
    Run('"' . html . '"')
}

Finance_ShowSettings() {
    global g_FinanceGui
    Finance_CloseGui()
    g_FinanceGui := Gui("+AlwaysOnTop +ToolWindow", "Finance settings")
    g_FinanceGui.SetFont("s10", "Segoe UI")
    g_FinanceGui.Add("Text", , "Dashboard widgets (1 = show)")
    keys := [
        ["ShowBalance", "Current balance / incomes / expenses / card"],
        ["ShowPies", "Pie charts"],
        ["ShowPerformance", "Performance metrics"],
        ["ShowGoals", "Goals overview"],
        ["ShowBudgets", "Budgets overview"],
        ["ShowNotifications", "Notification banners"]
    ]
    ctrls := Map()
    for k in keys {
        v := Finance_Setting("Dashboard", k[1], "1")
        ctrls[k[1]] := g_FinanceGui.Add("CheckBox", "Checked" . (v = "1" ? "1" : "0"), k[2])
    }
    g_FinanceGui.Add("Text", "y+16", "Alerts")
    n1 := g_FinanceGui.Add("CheckBox", "Checked" . (Finance_Setting("General", "NotifyBudgetExceeded", "1") = "1" ? "1" :
        "0"),
    "Budget exceeded")
    n2 := g_FinanceGui.Add("CheckBox", "Checked" . (Finance_Setting("General", "NotifyCardHighUsage", "1") = "1" ? "1" :
        "0"),
    "High credit-card usage")
    g_FinanceGui.Add("Button", "y+16 w100 Default", "Save").OnEvent("Click", SaveSettings)
    g_FinanceGui.Add("Button", "x+8 w100", "Back").OnEvent("Click", (*) => Finance_ShowMainMenu())
    g_FinanceGui.OnEvent("Escape", (*) => Finance_ShowMainMenu())
    g_FinanceGui.OnEvent("Close", (*) => Finance_CloseGui())
    Finance_CenterGui(g_FinanceGui, 520, 420)

    SaveSettings(*) {
        for k in keys
            Finance_SetSetting("Dashboard", k[1], ctrls[k[1]].Value ? "1" : "0")
        Finance_SetSetting("General", "NotifyBudgetExceeded", n1.Value ? "1" : "0")
        Finance_SetSetting("General", "NotifyCardHighUsage", n2.Value ? "1" : "0")
        Finance_Notify("Settings saved", 1200, BANNER_ACCENT_SUCCESS)
        Finance_ShowMainMenu()
    }
}

Finance_ShowImportMenu() {
    global g_FinanceGui
    Finance_CloseGui()
    g_FinanceGui := Gui("+AlwaysOnTop +ToolWindow", "AI import")
    g_FinanceGui.SetFont("s11", "Segoe UI")
    g_FinanceGui.Add("Text", "w520", "Import structured files from the Desktop (newest match).")
    g_FinanceGui.SetFont("s12 Bold", "Segoe UI")
    g_FinanceGui.Add("Text", "y+16", "[1]  Daily transactions   FINANCE_DAILY*.csv")
    g_FinanceGui.Add("Text", "y+8", "[2]  Monthly investments & goals   FINANCE_MONTHLY*.csv")
    g_FinanceGui.Add("Text", "y+8", "[3]  Copy daily prompt    [4]  Copy monthly prompt")
    g_FinanceGui.SetFont("s10 Norm", "Segoe UI")
    g_FinanceGui.Add("Text", "y+20 c555555", "Esc / Backspace = back")
    g_FinanceGui.OnEvent("Escape", (*) => Finance_ShowMainMenu())
    g_FinanceGui.OnEvent("Close", (*) => Finance_CloseGui())
    Finance_BindHotkeys([
        ["1", (*) => Finance_ImportDaily()],
        ["2", (*) => Finance_ImportMonthly()],
        ["3", (*) => Finance_CopyPrompt("daily_transactions.txt")],
        ["4", (*) => Finance_CopyPrompt("monthly_investments.txt")],
        ["Backspace", (*) => Finance_ShowMainMenu()],
        ["Escape", (*) => Finance_ShowMainMenu()]
    ])
    Finance_CenterGui(g_FinanceGui, 560, 240)
}
