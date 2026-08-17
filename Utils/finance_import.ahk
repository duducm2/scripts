; =============================================================================
; Utils module: finance_import.ahk
; Import AI-generated FINANCE_DAILY / FINANCE_MONTHLY files from Desktop
; =============================================================================

Finance_DesktopNewest(pattern) {
    newest := ""
    newestTime := 0
    loop files A_Desktop . "\" . pattern, "F" {
        ts := Number(A_LoopFileTimeModified)
        if (ts > newestTime) {
            newestTime := ts
            newest := A_LoopFileFullPath
        }
    }
    return newest
}

; Skip Gemini preambles (e.g. "FILE: FINANCE_DAILY.csv") so the header row is first.
Finance_ReadAiImportCsv(path) {
    text := Finance_ReadUtf8(path)
    if (text = "")
        return []
    text := StrReplace(text, "`r`n", "`n")
    text := StrReplace(text, "`r", "`n")
    lines := StrSplit(text, "`n")
    start := 0
    for idx, line in lines {
        t := Trim(line)
        if (t = "")
            continue
        lower := StrLower(t)
        if (InStr(lower, "date,") = 1 && InStr(lower, "description") && InStr(lower, "amount")) {
            start := idx
            break
        }
        if (InStr(lower, "file:") = 1 || SubStr(t, 1, 1) = "#")
            continue
    }
    if (!start)
        return Finance_ReadCsv(path)
    cleaned := ""
    loop lines.Length {
        if (A_Index < start)
            continue
        if (cleaned != "")
            cleaned .= "`n"
        cleaned .= lines[A_Index]
    }
    tmp := A_Temp . "\finance_ai_import_norm.csv"
    Finance_WriteUtf8(tmp, cleaned)
    rows := Finance_ReadCsv(tmp)
    try FileDelete(tmp)
    catch {
    }
    return rows
}

; Newest Desktop dump that looks like a daily finance CSV (Gemini code downloads).
Finance_DesktopNewestDailyCodeDump() {
    newest := ""
    newestTime := 0
    loop files A_Desktop . "\gemini-code*.txt", "F" {
        body := Finance_ReadUtf8(A_LoopFileFullPath)
        if (body = "")
            continue
        lower := StrLower(body)
        if (!InStr(lower, "date,description,amount"))
            continue
        ts := Number(A_LoopFileTimeModified)
        if (ts > newestTime) {
            newestTime := ts
            newest := A_LoopFileFullPath
        }
    }
    return newest
}

Finance_ArchiveImported(path) {
    destDir := Finance_DataDir() . "\imported"
    if (!DirExist(destDir))
        DirCreate(destDir)
    SplitPath(path, &name)
    dest := destDir . "\" . FormatTime(, "yyyyMMdd-HHmmss") . "_" . name
    try FileMove(path, dest, 1)
    catch {
        try FileCopy(path, dest, 1)
        catch {
        }
    }
}

Finance_ImportConfirm(title, lines) {
    global g_FinanceGui
    owner := ""
    try {
        if (IsObject(g_FinanceGui))
            owner := " +Owner" . g_FinanceGui.Hwnd
    } catch {
        owner := ""
    }
    Finance_DialogsBegin()
    g := Gui("+AlwaysOnTop +ToolWindow" . owner, title)
    g.SetFont("s10", "Segoe UI")
    g.Add("Text", "w640", "Confirm import of " . lines.Length . " row(s).")
    lv := g.Add("ListView", "w640 r14 Grid", ["Row"])
    for line in lines
        lv.Add("", line)
    lv.ModifyCol(1, 620)
    ok := false
    g.Add("Button", "y+10 w120 Default", "Import").OnEvent("Click", ConfirmYes)
    g.Add("Button", "x+8 w120", "Cancel").OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())
    g.Show()
    try WinWaitClose("ahk_id " g.Hwnd)
    catch {
    }
    Finance_DialogsEnd()
    return ok

    ConfirmYes(*) {
        ok := true
        g.Destroy()
    }
}

Finance_ImportDaily(*) {
    path := Finance_DesktopNewest("FINANCE_DAILY*.csv")
    if (path = "")
        path := Finance_DesktopNewest("FINANCE_DAILY*.txt")
    if (path = "")
        path := Finance_DesktopNewest("FINANCE_DAILY*.ini")
    if (path = "")
        path := Finance_DesktopNewestDailyCodeDump()
    if (path = "") {
        Finance_Notify("No FINANCE_DAILY file on Desktop", 2000, BANNER_ACCENT_ERROR)
        return
    }
    rows := Finance_ReadAiImportCsv(path)
    if (!rows.Length) {
        Finance_Notify("File has no data rows", 1800, BANNER_ACCENT_ERROR)
        return
    }
    lines := []
    parsed := []
    for r in rows {
        desc := r.Has("description") ? r["description"] : ""
        amt := r.Has("amount") ? r["amount"] : "0"
        t := r.Has("type") ? r["type"] : "expense"
        date := r.Has("date") && r["date"] != "" ? r["date"] : Finance_Today()
        parsed.Push(Map(
            "date", date,
            "description", desc,
            "amount", Finance_FormatCsvDecimal(Abs(Finance_ParseDecimal(amt))),
            "type", t,
            "category_id", r.Has("category_id") ? r["category_id"] : "",
            "subcategory", r.Has("subcategory") ? r["subcategory"] : "",
            "account_id", r.Has("account_id") ? r["account_id"] : "",
            "card_id", r.Has("card_id") ? r["card_id"] : "",
            "transfer_account_id", r.Has("transfer_account_id") ? r["transfer_account_id"] : ""
        ))
        lines.Push(date . "  " . t . "  " . desc . "  " . amt)
    }
    if (!Finance_ImportConfirm("Import daily transactions", lines))
        return
    txs := Finance_Load("transactions")
    for p in parsed {
        p["id"] := Finance_NextId("TX", txs)
        txs.Push(p)
        Finance_ApplyTransactionToBalances(p, false)
    }
    Finance_Save("transactions", txs)
    Finance_RecomputeBudgetSpent(Finance_CurrentYearMonth())
    Finance_ArchiveImported(path)
    Finance_Notify("Imported " . parsed.Length . " transactions", 1800, BANNER_ACCENT_SUCCESS)
    Finance_ShowTransactions()
}

Finance_ImportMonthly(*) {
    path := Finance_DesktopNewest("FINANCE_MONTHLY*.csv")
    if (path = "")
        path := Finance_DesktopNewest("FINANCE_MONTHLY*.ini")
    if (path = "") {
        Finance_Notify("No FINANCE_MONTHLY file on Desktop", 2000, BANNER_ACCENT_ERROR)
        return
    }
    rows := Finance_ReadCsv(path)
    if (!rows.Length) {
        Finance_Notify("File has no data rows", 1800, BANNER_ACCENT_ERROR)
        return
    }
    lines := []
    for r in rows {
        et := r.Has("entity_type") ? r["entity_type"] : ""
        eid := r.Has("entity_id") ? r["entity_id"] : ""
        adj := r.Has("adjustment_amount") ? r["adjustment_amount"] : "0"
        desc := r.Has("description") ? r["description"] : ""
        lines.Push(et . "  " . eid . "  " . adj . "  " . desc)
    }
    if (!Finance_ImportConfirm("Import monthly adjustments", lines))
        return
    accs := Finance_Load("accounts")
    goals := Finance_Load("goals")
    txs := Finance_Load("transactions")
    n := 0
    for r in rows {
        et := r.Has("entity_type") ? StrLower(Trim(r["entity_type"])) : ""
        eid := r.Has("entity_id") ? Trim(r["entity_id"]) : ""
        adj := Finance_ParseDecimal(r.Has("adjustment_amount") ? r["adjustment_amount"] : "0")
        desc := r.Has("description") ? r["description"] : "Monthly adjustment"
        if (et = "account") {
            acc := Finance_FindById(accs, eid)
            if (!acc)
                continue
            Finance_AdjustAccount(accs, eid, adj)
            tx := Map(
                "id", Finance_NextId("TX", txs),
                "date", Finance_Today(),
                "description", desc,
                "amount", Finance_FormatCsvDecimal(Abs(adj)),
                "type", adj >= 0 ? "income" : "expense",
                "category_id", adj >= 0 ? Finance_CatIdByName("Investments") : Finance_CatIdByName("Adjustment"),
            "subcategory", "",
            "account_id", eid,
            "card_id", "",
            "transfer_account_id", ""
            )
            txs.Push(tx)
            n += 1
        } else if (et = "goal") {
            g := Finance_FindById(goals, eid)
            if (!g)
                continue
            cur := Finance_ParseDecimal(g["current_amount"]) + adj
            g["current_amount"] := Finance_FormatCsvDecimal(cur)
            n += 1
        }
    }
    Finance_Save("accounts", accs)
    Finance_Save("goals", goals)
    Finance_Save("transactions", txs)
    Finance_ArchiveImported(path)
    Finance_Notify("Applied " . n . " monthly adjustments", 1800, BANNER_ACCENT_SUCCESS)
    Finance_ShowMainMenu()
}
