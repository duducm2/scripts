; =============================================================================
; Utils module: finance_launcher.ahk
; Finance app main menu (Utility Shortcuts [F] / Win+Alt+Shift+D)
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
        ["D", "Dashboard", "Cockpit charts and widgets"],
        ["T", "Transactions", "List, filter, edit"],
        ["A", "Accounts", "Balances, CRUD, primary"],
        ["C", "Credit cards", "Limits, spent, pay, primary"],
        ["B", "Budgets", "Monthly limits"],
        ["G", "Goals", "Funds and targets"],
        ["L", "Recurring bills", "Tracking only, no auto-charge"],
        ["K", "Categories", "CRUD, search, filter"],
        ["I", "AI import", "Desktop daily / monthly"],
        ["P", "Push to cloud", "Commit and push scripts repo"],
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
    g_FinanceGui.Add("Text", "x20 y586 w880", "Esc close   letters open a module   P push   S settings")

    Finance_BindHotkeys([
        ["d", Finance_OnDash], ["t", Finance_OnTx], ["a", Finance_OnAcc], ["c", Finance_OnCard],
        ["b", Finance_OnBud], ["g", Finance_OnGoals], ["l", Finance_OnRec], ["k", Finance_OnCat],
        ["i", Finance_OnImp], ["p", Finance_OnGitPush], ["s", Finance_OnSet], ["Escape", (*) => Finance_CloseGui()]
    ])
    Finance_CenterGui(g_FinanceGui, 900, 620)
}

Finance_OnDash(*) {
    Finance_OpenDashboard()
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
Finance_OnRec(*) {
    Finance_ShowRecurring()
}
Finance_OnCat(*) {
    Finance_ShowCategories()
}
Finance_OnImp(*) {
    Finance_ShowImportMenu()
}
Finance_OnGitPush(*) {
    Finance_GitSyncPush()
}
Finance_OnSet(*) {
    Finance_ShowSettings()
}

Finance_OpenDashboard() {
    Finance_EnsureData()
    py := Finance_PythonDir() . "\chart_generator.py"
    if (!FileExist(py)) {
        Finance_Notify("chart_generator.py not found", 2000, BANNER_ACCENT_ERROR)
        return
    }
    dataDir := Finance_DataDir()
    outDir := Finance_OutputDir()
    pyCmd := Finance_FindPythonCmd()
    if (pyCmd = "") {
        Finance_Notify("Python not found. Install Python or enable the py launcher.", 3500, BANNER_ACCENT_ERROR)
        return
    }
    try StandardLoadingBar_Show("Building dashboard…", BANNER_ACCENT_INTERMEDIATE)
    catch {
    }
    cmd := pyCmd . ' "' . py . '" --data-dir "' . dataDir . '" --output-dir "' . outDir . '"'
    exitCode := 0
    try {
        exitCode := RunWait(A_ComSpec . ' /c ' . cmd, A_ScriptDir, "Hide")
    } catch as e {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        Finance_Notify("Python failed: " . e.Message, 2500, BANNER_ACCENT_ERROR)
        return
    }
    if (exitCode != 0) {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        Finance_Notify("Dashboard Python failed (exit " . exitCode . ")", 3500, BANNER_ACCENT_ERROR)
        return
    }
    html := outDir . "\dashboard.html"
    try StandardLoadingBar_Hide(400)
    catch {
    }
    if (!FileExist(html)) {
        Finance_Notify("dashboard.html was not generated", 2200, BANNER_ACCENT_ERROR)
        return
    }
    tmpHtml := A_Temp . "\finance_dashboard_" . A_TickCount . ".html"
    try FileCopy(html, tmpHtml, 1)
    catch {
        tmpHtml := html
    }
    fileUrl := "file:///" . StrReplace(StrReplace(tmpHtml, "\", "/"), " ", "%20") . "?t=" . A_TickCount
    try Run('chrome.exe --new-window "' . fileUrl . '"')
    catch as e {
        Finance_Notify("Chrome failed: " . e.Message, 2500, BANNER_ACCENT_ERROR)
        return
    }
    Finance_CloseGui()
}

Finance_ShowSettings() {
    global g_FinanceGui
    Finance_CloseGui()
    Finance_EnsureData()
    g_FinanceGui := Gui("+AlwaysOnTop +ToolWindow", "Finance settings")
    g_FinanceGui.SetFont("s10", "Segoe UI")
    g_FinanceGui.Add("Text", "x12 y10 w480", "Dashboard widgets")
    keys := [
        ["ShowBalance", "Current balance / incomes / expenses / card"],
        ["ShowAccounts", "Accounts comparison"],
        ["ShowPies", "Pie charts"],
        ["ShowPerformance", "Performance metrics"],
        ["ShowGoals", "Goals overview"],
        ["ShowBudgets", "Budgets overview"],
        ["ShowRecurring", "Recurring bills"],
        ["ShowNotifications", "Notification banners"]
    ]
    ctrls := Map()
    y := 34
    for k in keys {
        v := Finance_Setting("Dashboard", k[1], "1")
        ctrls[k[1]] := g_FinanceGui.Add("CheckBox",
            "x12 y" . y . " w480 Checked" . (v = "1" ? "1" : "0"), k[2])
        y += 24
    }
    y += 8
    g_FinanceGui.Add("Text", "x12 y" . y . " w480", "Alerts")
    y += 24
    n1 := g_FinanceGui.Add("CheckBox",
        "x12 y" . y . " w480 Checked" . (Finance_Setting("General", "NotifyBudgetExceeded", "1") = "1" ? "1" : "0"),
        "Budget exceeded")
    y += 24
    n2 := g_FinanceGui.Add("CheckBox",
        "x12 y" . y . " w480 Checked" . (Finance_Setting("General", "NotifyCardHighUsage", "1") = "1" ? "1" : "0"),
        "High credit-card usage")
    y += 32
    g_FinanceGui.Add("Button", "x12 y" . y . " w100 Default", "Save").OnEvent("Click", SaveSettings)
    g_FinanceGui.Add("Button", "x120 y" . y . " w100", "Back").OnEvent("Click", (*) => Finance_ShowMainMenu())
    g_FinanceGui.Add("Button", "x230 y" . y . " w200", "Rebuild balances [R]").OnEvent("Click", (*) =>
        Finance_OnRebuildBalances())
    y += 36
    g_FinanceGui.SetFont("s9 c555555", "Segoe UI")
    g_FinanceGui.Add("Text", "x12 y" . y . " w480",
        "Rebuild: reset accounts/cards from initial balances, replay all transactions. Esc / Backspace = back")
    g_FinanceGui.OnEvent("Escape", (*) => Finance_ShowMainMenu())
    g_FinanceGui.OnEvent("Close", (*) => Finance_CloseGui())
    Finance_BindHotkeys([
        ["r", (*) => Finance_OnRebuildBalances()],
        ["Backspace", (*) => Finance_ShowMainMenu()],
        ["Escape", (*) => Finance_ShowMainMenu()]
    ])
    Finance_CenterGui(g_FinanceGui, 510, y + 40)

    SaveSettings(*) {
        for k in keys
            Finance_SetSetting("Dashboard", k[1], ctrls[k[1]].Value ? "1" : "0")
        Finance_SetSetting("General", "NotifyBudgetExceeded", n1.Value ? "1" : "0")
        Finance_SetSetting("General", "NotifyCardHighUsage", n2.Value ? "1" : "0")
        Finance_Notify("Settings saved", 1200, BANNER_ACCENT_SUCCESS)
        Finance_ShowMainMenu()
    }
}

Finance_OnRebuildBalances(*) {
    if (!Finance_Confirm(
        "Reset account balances from initial_balance and card spent from initial_spent, then replay every transaction? This cannot be undone except by Restore from backups.",
        "Rebuild balances"))
        return
    n := Finance_RebuildBalancesFromTransactions()
    Finance_Notify("Rebuilt balances from " . n . " transaction(s)", 2200, BANNER_ACCENT_SUCCESS)
    Finance_ShowMainMenu()
}

Finance_ShowImportMenu() {
    global g_FinanceGui
    Finance_CloseGui()
    g_FinanceGui := Gui("+AlwaysOnTop +ToolWindow", "AI import")
    g_FinanceGui.SetFont("s11", "Segoe UI")
    g_FinanceGui.Add("Text", "w520",
        "Import FINANCE_*.txt packs from the Desktop (newest match). Extracts CSV from ===FILE=== sections. Also accepts legacy .csv and gemini-code*.txt dumps."
    )
    g_FinanceGui.SetFont("s12 Bold", "Segoe UI")
    g_FinanceGui.Add("Text", "y+16", "[1]  Daily transactions   FINANCE_DAILY*.txt")
    g_FinanceGui.Add("Text", "y+8", "[2]  Monthly investments & goals   FINANCE_MONTHLY*.txt")
    g_FinanceGui.SetFont("s10 Norm", "Segoe UI")
    g_FinanceGui.Add("Text", "y+16 c555555 w520",
        "Prompts: Utility Shortcuts (#!+U) → Prompts → [d] daily / [m] monthly")
    g_FinanceGui.Add("Text", "y+8 c555555", "Esc / Backspace = back")
    g_FinanceGui.OnEvent("Escape", (*) => Finance_ShowMainMenu())
    g_FinanceGui.OnEvent("Close", (*) => Finance_CloseGui())
    Finance_BindHotkeys([
        ["1", (*) => Finance_ImportDaily()],
        ["2", (*) => Finance_ImportMonthly()],
        ["Backspace", (*) => Finance_ShowMainMenu()],
        ["Escape", (*) => Finance_ShowMainMenu()]
    ])
    Finance_CenterGui(g_FinanceGui, 560, 220)
}
