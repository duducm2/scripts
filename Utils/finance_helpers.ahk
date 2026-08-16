; =============================================================================
; Utils module: finance_helpers.ahk
; CSV storage, Brazilian decimals, primary keys, seed/migration
; =============================================================================

global g_FinanceGui := false
global g_FinanceHotkeys := []
global g_FinanceMonth := ""
global g_FinanceNotifyQueue := []

Finance_DataDir() {
    dir := A_ScriptDir . "\finances\data"
    if (!DirExist(dir))
        DirCreate(dir)
    imported := dir . "\imported"
    if (!DirExist(imported))
        DirCreate(imported)
    return dir
}

Finance_OutputDir() {
    dir := A_ScriptDir . "\finances\output"
    if (!DirExist(dir))
        DirCreate(dir)
    return dir
}

Finance_PythonDir() {
    return A_ScriptDir . "\finances\python"
}

Finance_SettingsPath() {
    return Finance_DataDir() . "\settings.ini"
}

Finance_EnsureData() {
    Finance_DataDir()
    Finance_EnsureSettings()
    if (!FileExist(Finance_DataDir() . "\categories.csv"))
        Finance_MigrateCategoriesFromIni()
    if (!FileExist(Finance_DataDir() . "\accounts.csv"))
        Finance_MigrateAccountsFromIni()
    if (!FileExist(Finance_DataDir() . "\credit_cards.csv"))
        Finance_SeedCreditCards()
    if (!FileExist(Finance_DataDir() . "\goals.csv"))
        Finance_SeedGoals()
    if (!FileExist(Finance_DataDir() . "\budgets.csv"))
        Finance_SeedBudgets()
    if (!FileExist(Finance_DataDir() . "\transactions.csv"))
        Finance_SeedTransactions()
    Finance_FixDefaultIds()
    Finance_EnsureMonthBudgets(Finance_CurrentYearMonth())
}

Finance_FixDefaultIds() {
    acc := Finance_AccIdByNameContains("Mercado Pago main")
    if (acc = "")
        acc := Finance_AccIdByNameContains("Mercado Pago")
    if (acc != "")
        Finance_SetSetting("General", "DefaultAccountId", acc)
    cards := Finance_Load("credit_cards")
    if (cards.Length)
        Finance_SetSetting("General", "PrimaryCardId", cards[1]["id"])
}

Finance_EnsureSettings() {
    path := Finance_SettingsPath()
    if (FileExist(path))
        return
    content := "[Dashboard]`n"
        . "ShowBalance=1`n"
        . "ShowPies=1`n"
        . "ShowPerformance=1`n"
        . "ShowGoals=1`n"
        . "ShowBudgets=1`n"
        . "ShowNotifications=1`n"
        . "`n[General]`n"
        . "DefaultAccountId=ACC_MP_MAIN`n"
        . "PrimaryCardId=CARD_MP`n"
        . "NotifyBudgetExceeded=1`n"
        . "NotifyCardHighUsage=1`n"
        . "CardUsageWarnPct=80`n"
    Finance_WriteUtf8(path, content)
}

Finance_Setting(section, key, default := "") {
    val := IniRead(Finance_SettingsPath(), section, key, default)
    if (val = "ERROR")
        return default
    return val
}

Finance_SetSetting(section, key, value) {
    IniWrite(value, Finance_SettingsPath(), section, key)
}

Finance_WriteUtf8(path, content) {
    f := FileOpen(path, "w", "UTF-8")
    if (!f)
        throw Error("Could not write " . path)
    f.Write(content)
    f.Close()
}

Finance_ReadUtf8(path) {
    if (!FileExist(path))
        return ""
    f := FileOpen(path, "r", "UTF-8")
    if (!f)
        return ""
    text := f.Read()
    f.Close()
    if (SubStr(text, 1, 1) = Chr(0xFEFF))
        text := SubStr(text, 2)
    return text
}

; --- decimals (Brazilian comma) ---

Finance_NormalizeDot(str) {
    s := Trim(String(str))
    s := StrReplace(s, "R$", "")
    s := StrReplace(s, " ", "")
    s := StrReplace(s, "`t", "")
    if (s = "")
        return "0,00"
    neg := false
    if (SubStr(s, 1, 1) = "-") {
        neg := true
        s := SubStr(s, 2)
    }
    if (SubStr(s, 1, 1) = "+")
        s := SubStr(s, 2)
    lastComma := InStr(s, ",", , -1)
    lastDot := InStr(s, ".", , -1)
    if (lastComma && lastDot) {
        if (lastComma > lastDot)
            s := StrReplace(StrReplace(s, ".", ""), ",", ".")
        else
            s := StrReplace(s, ",", "")
    } else if (lastComma) {
        parts := StrSplit(s, ",")
        if (parts.Length = 2 && StrLen(parts[2]) <= 2) {
            s := StrReplace(s, ".", "")
            s := StrReplace(s, ",", ".")
        } else {
            s := StrReplace(s, ",", "")
        }
    } else if (lastDot) {
        parts := StrSplit(s, ".")
        if (parts.Length = 2 && StrLen(parts[2]) <= 2) {
            s := StrReplace(s, ",", "")
        } else {
            s := StrReplace(s, ".", "")
        }
    }
    if (neg)
        s := "-" . s
    return s
}

Finance_ParseDecimal(str) {
    n := Finance_NormalizeDot(str)
    if (n = "" || n = "-")
        return 0.0
    return Number(n)
}

Finance_FormatDecimal(num, decimals := 2) {
    n := Number(num)
    neg := n < 0
    if (neg)
        n := -n
    s := Format("{:0." . decimals . "f}", n)
    parts := StrSplit(s, ".")
    intPart := parts[1]
    frac := (parts.Length > 1) ? parts[2] : ""
    grouped := ""
    while (StrLen(intPart) > 3) {
        grouped := "." . SubStr(intPart, -2) . grouped
        intPart := SubStr(intPart, 1, StrLen(intPart) - 3)
    }
    grouped := intPart . grouped
    out := grouped . "," . frac
    return neg ? "-" . out : out
}

Finance_FormatCsvDecimal(num, decimals := 2) {
    n := Number(num)
    s := Format("{:0." . decimals . "f}", n)
    return StrReplace(s, ".", ",")
}

Finance_FormatBrl(num) {
    return "R$ " . Finance_FormatDecimal(num)
}

; --- CSV ---

Finance_SplitCsvLine(line) {
    fields := []
    i := 1
    len := StrLen(line)
    if (len && SubStr(line, len, 1) = "`r") {
        line := SubStr(line, 1, len - 1)
        len := StrLen(line)
    }
    while (i <= len) {
        if (SubStr(line, i, 1) = '"') {
            i += 1
            val := ""
            while (i <= len) {
                c := SubStr(line, i, 1)
                if (c = '"') {
                    if (i < len && SubStr(line, i + 1, 1) = '"') {
                        val .= '"'
                        i += 2
                        continue
                    }
                    i += 1
                    break
                }
                val .= c
                i += 1
            }
            fields.Push(val)
            if (i <= len && SubStr(line, i, 1) = ",")
                i += 1
        } else {
            next := InStr(line, ",", false, i)
            if (!next) {
                fields.Push(SubStr(line, i))
                break
            }
            fields.Push(SubStr(line, i, next - i))
            i := next + 1
            if (i > len)
                fields.Push("")
        }
    }
    return fields
}

Finance_CsvEscape(val) {
    s := String(val)
    if (InStr(s, ",") || InStr(s, '"') || InStr(s, "`n") || InStr(s, "`r"))
        return '"' . StrReplace(s, '"', '""') . '"'
    return s
}

Finance_ReadCsv(fileName) {
    path := InStr(fileName, "\") || InStr(fileName, "/") ? fileName : Finance_DataDir() . "\" . fileName
    rows := []
    text := Finance_ReadUtf8(path)
    if (text = "")
        return rows
    text := StrReplace(text, "`r`n", "`n")
    text := StrReplace(text, "`r", "`n")
    lines := StrSplit(text, "`n")
    headers := []
    for idx, line in lines {
        if (Trim(line) = "")
            continue
        fields := Finance_SplitCsvLine(line)
        if (headers.Length = 0) {
            for h in fields
                headers.Push(Trim(h))
            continue
        }
        row := Map()
        loop headers.Length {
            key := headers[A_Index]
            row[key] := (A_Index <= fields.Length) ? fields[A_Index] : ""
        }
        rows.Push(row)
    }
    return rows
}

Finance_WriteCsv(fileName, rows, headers) {
    path := InStr(fileName, "\") || InStr(fileName, "/") ? fileName : Finance_DataDir() . "\" . fileName
    out := ""
    loop headers.Length {
        if (A_Index > 1)
            out .= ","
        out .= Finance_CsvEscape(headers[A_Index])
    }
    out .= "`n"
    for row in rows {
        loop headers.Length {
            if (A_Index > 1)
                out .= ","
            key := headers[A_Index]
            val := row.Has(key) ? row[key] : ""
            out .= Finance_CsvEscape(val)
        }
        out .= "`n"
    }
    Finance_WriteUtf8(path, out)
}

Finance_Headers(kind) {
    switch kind {
        case "transactions":
            return ["id", "date", "description", "amount", "type", "category_id", "subcategory", "account_id",
                "card_id", "transfer_account_id"]
        case "accounts":
            return ["id", "name", "icon", "initial_balance", "current_balance"]
        case "categories":
            return ["id", "name", "type", "parent_id", "color", "icon"]
        case "credit_cards":
            return ["id", "name", "limit", "current_spent", "linked_account_id", "closing_day"]
        case "goals":
            return ["id", "name", "current_amount", "target_amount", "target_date", "status"]
        case "budgets":
            return ["year_month", "category_id", "planned_amount", "spent_amount"]
        default:
            return []
    }
}

Finance_Save(kind, rows) {
    Finance_WriteCsv(kind . ".csv", rows, Finance_Headers(kind))
}

Finance_Load(kind) {
    return Finance_ReadCsv(kind . ".csv")
}

Finance_FindById(rows, id) {
    for row in rows {
        if (row["id"] = id)
            return row
    }
    return false
}

Finance_NextId(prefix, rows, pad := 3) {
    maxN := 0
    for row in rows {
        id := row.Has("id") ? row["id"] : ""
        if (SubStr(id, 1, StrLen(prefix)) != prefix)
            continue
        rest := SubStr(id, StrLen(prefix) + 1)
        if (rest != "" && IsDigit(rest)) {
            n := Integer(rest)
            if (n > maxN)
                maxN := n
        }
    }
    return prefix . Format("{:0" . pad . "d}", maxN + 1)
}

Finance_SlugId(prefix, name, existing) {
    base := Finance_Slug(name)
    id := prefix . base
    if (!Finance_IdExists(existing, id))
        return id
    n := 2
    loop {
        cand := id . n
        if (!Finance_IdExists(existing, cand))
            return cand
        n += 1
    }
}

Finance_IdExists(rows, id) {
    for row in rows {
        if (row.Has("id") && row["id"] = id)
            return true
    }
    return false
}

Finance_Slug(name) {
    s := Finance_Unaccent(Trim(name))
    s := StrUpper(s)
    out := ""
    loop parse s {
        ch := A_LoopField
        if ((ch >= "A" && ch <= "Z") || (ch >= "0" && ch <= "9"))
            out .= ch
    }
    if (StrLen(out) > 8)
        out := SubStr(out, 1, 8)
    if (out = "")
        out := "X"
    return out
}

Finance_Unaccent(s) {
    pairs := [["á", "a"], ["à", "a"], ["â", "a"], ["ã", "a"], ["ä", "a"], ["é", "e"], ["ê", "e"], ["è", "e"],
    ["í", "i"], ["ó", "o"], ["ô", "o"], ["õ", "o"], ["ö", "o"], ["ú", "u"], ["ü", "u"], ["ç", "c"],
    ["Á", "A"], ["À", "A"], ["Â", "A"], ["Ã", "A"], ["É", "E"], ["Ê", "E"], ["Í", "I"], ["Ó", "O"],
    ["Ô", "O"], ["Õ", "O"], ["Ú", "U"], ["Ç", "C"]]
    for p in pairs
        s := StrReplace(s, p[1], p[2])
    return s
}

Finance_CurrentYearMonth() {
    return FormatTime(, "yyyy-MM")
}

Finance_ShiftMonth(yearMonth, delta) {
    parts := StrSplit(yearMonth, "-")
    y := Integer(parts[1])
    m := Integer(parts[2]) + delta
    while (m > 12) {
        m -= 12
        y += 1
    }
    while (m < 1) {
        m += 12
        y -= 1
    }
    return Format("{:04}-{:02}", y, m)
}

Finance_MonthLabel(yearMonth) {
    parts := StrSplit(yearMonth, "-")
    names := ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
        "November", "December"]
    m := Integer(parts[2])
    return names[m] . " " . parts[1]
}

Finance_Today() {
    return FormatTime(, "yyyy-MM-dd")
}

Finance_TxInMonth(row, yearMonth) {
    d := row["date"]
    return (SubStr(d, 1, 7) = yearMonth)
}

Finance_TxSign(type, amount) {
    a := Finance_ParseDecimal(amount)
    if (type = "income")
        return a
    if (type = "expense" || type = "card_expense")
        return -a
    return 0.0
}

Finance_TypeLabel(type) {
    switch type {
        case "expense":
            return "Expense"
        case "income":
            return "Income"
        case "transfer":
            return "Transfer"
        case "card_expense":
            return "Credit card"
        case "adjustment":
            return "Adjustment"
        default:
            return type
    }
}

Finance_CatName(cats, id) {
    row := Finance_FindById(cats, id)
    return row ? row["name"] : id
}

Finance_AccName(accs, id) {
    row := Finance_FindById(accs, id)
    return row ? row["name"] : id
}

Finance_MainCategories(cats, type := "") {
    out := []
    for c in cats {
        if (c["parent_id"] != "")
            continue
        if (type != "" && c["type"] != type)
            continue
        out.Push(c)
    }
    return out
}

Finance_Subcategories(cats, parentId) {
    out := []
    for c in cats {
        if (c["parent_id"] = parentId)
            out.Push(c)
    }
    return out
}

Finance_CanAddMainCategory(cats) {
    n := 0
    for c in cats {
        if (c["parent_id"] = "")
            n += 1
    }
    return n < 50
}

Finance_CanAddSubcategory(cats, parentId) {
    return Finance_Subcategories(cats, parentId).Length < 10
}

Finance_ApplyTransactionToBalances(tx, reverse := false) {
    accs := Finance_Load("accounts")
    cards := Finance_Load("credit_cards")
    amt := Finance_ParseDecimal(tx["amount"])
    if (reverse)
        amt := -amt
    type := tx["type"]
    accId := tx["account_id"]
    if (type = "income" || type = "adjustment") {
        Finance_AdjustAccount(accs, accId, amt)
    } else if (type = "expense") {
        Finance_AdjustAccount(accs, accId, -amt)
    } else if (type = "card_expense") {
        Finance_AdjustCard(cards, tx["card_id"], amt)
    } else if (type = "transfer") {
        Finance_AdjustAccount(accs, accId, -amt)
        dest := tx.Has("transfer_account_id") ? tx["transfer_account_id"] : ""
        if (dest != "")
            Finance_AdjustAccount(accs, dest, amt)
    }
    Finance_Save("accounts", accs)
    Finance_Save("credit_cards", cards)
}

Finance_AdjustAccount(accs, id, delta) {
    if (id = "")
        return
    row := Finance_FindById(accs, id)
    if (!row)
        return
    cur := Finance_ParseDecimal(row["current_balance"])
    row["current_balance"] := Finance_FormatCsvDecimal(cur + delta)
}

Finance_AdjustCard(cards, id, delta) {
    if (id = "")
        return
    row := Finance_FindById(cards, id)
    if (!row)
        return
    cur := Finance_ParseDecimal(row["current_spent"])
    row["current_spent"] := Finance_FormatCsvDecimal(cur + delta)
}

Finance_ReplaceTransaction(oldTx, newTx) {
    if (IsObject(oldTx))
        Finance_ApplyTransactionToBalances(oldTx, true)
    Finance_ApplyTransactionToBalances(newTx, false)
}

Finance_MonthTotals(yearMonth) {
    txs := Finance_Load("transactions")
    inc := 0.0
    exp := 0.0
    for tx in txs {
        if (!Finance_TxInMonth(tx, yearMonth))
            continue
        t := tx["type"]
        a := Finance_ParseDecimal(tx["amount"])
        if (t = "income")
            inc += a
        else if (t = "expense" || t = "card_expense")
            exp += a
    }
    return { income: inc, expense: exp, balance: inc - exp }
}

Finance_TotalBalance() {
    accs := Finance_Load("accounts")
    tot := 0.0
    for a in accs
        tot += Finance_ParseDecimal(a["current_balance"])
    return tot
}

Finance_RecomputeBudgetSpent(yearMonth) {
    budgets := Finance_Load("budgets")
    txs := Finance_Load("transactions")
    cats := Finance_Load("categories")
    spentByCat := Map()
    for tx in txs {
        if (!Finance_TxInMonth(tx, yearMonth))
            continue
        if (tx["type"] != "expense" && tx["type"] != "card_expense")
            continue
        cid := tx["category_id"]
        cat := Finance_FindById(cats, cid)
        mainId := cid
        if (cat && cat["parent_id"] != "")
            mainId := cat["parent_id"]
        a := Finance_ParseDecimal(tx["amount"])
        spentByCat[mainId] := (spentByCat.Has(mainId) ? spentByCat[mainId] : 0.0) + a
    }
    for b in budgets {
        if (b["year_month"] != yearMonth)
            continue
        s := spentByCat.Has(b["category_id"]) ? spentByCat[b["category_id"]] : 0.0
        b["spent_amount"] := Finance_FormatCsvDecimal(s)
    }
    Finance_Save("budgets", budgets)
    return budgets
}

Finance_EnsureMonthBudgets(yearMonth) {
    budgets := Finance_Load("budgets")
    has := false
    for b in budgets {
        if (b["year_month"] = yearMonth) {
            has := true
            break
        }
    }
    if (has) {
        Finance_RecomputeBudgetSpent(yearMonth)
        return
    }
    prev := Finance_ShiftMonth(yearMonth, -1)
    copied := false
    for b in budgets {
        if (b["year_month"] = prev) {
            budgets.Push(Map(
                "year_month", yearMonth,
                "category_id", b["category_id"],
                "planned_amount", b["planned_amount"],
                "spent_amount", "0,00"
            ))
            copied := true
        }
    }
    if (copied)
        Finance_Save("budgets", budgets)
    Finance_RecomputeBudgetSpent(yearMonth)
}

Finance_Palette() {
    return ["#2ECC71", "#3498DB", "#9B59B6", "#E67E22", "#E74C3C", "#1ABC9C", "#F1C40F", "#34495E",
        "#16A085", "#27AE60", "#2980B9", "#8E44AD", "#D35400", "#C0392B", "#7F8C8D", "#2C3E50"]
}

Finance_ColorForIndex(i) {
    pal := Finance_Palette()
    return pal[Mod(i - 1, pal.Length) + 1]
}

; --- GUI shared ---

Finance_CloseGui() {
    global g_FinanceGui, g_FinanceHotkeys
    Finance_UnbindHotkeys()
    if (IsObject(g_FinanceGui)) {
        try g_FinanceGui.Destroy()
        catch {
        }
    }
    g_FinanceGui := false
}

Finance_UnbindHotkeys() {
    global g_FinanceHotkeys, g_FinanceGui
    hwnd := 0
    try {
        if (IsObject(g_FinanceGui))
            hwnd := g_FinanceGui.Hwnd
    } catch {
        hwnd := 0
    }
    if (hwnd) {
        try HotIfWinActive("ahk_id " hwnd)
        catch {
        }
    }
    for item in g_FinanceHotkeys {
        try Hotkey(item, "Off")
        catch {
        }
    }
    g_FinanceHotkeys := []
    try HotIf()
    catch {
    }
}

Finance_HotkeyNoop(*) {
}

Finance_BindHotkeys(pairs) {
    global g_FinanceGui, g_FinanceHotkeys
    Finance_UnbindHotkeys()
    if (!IsObject(g_FinanceGui))
        return
    try HotIfWinActive("ahk_id " g_FinanceGui.Hwnd)
    catch {
        return
    }
    for p in pairs {
        try {
            Hotkey(p[1], p[2], "On")
            g_FinanceHotkeys.Push(p[1])
        } catch {
        }
    }
    try HotIf()
    catch {
    }
}

Finance_CenterGui(guiObj, w := 920, h := 620) {
    MonitorGetWorkArea(MonitorGetPrimary(), &L, &T, &R, &B)
    x := L + ((R - L) - w) // 2
    y := T + ((B - T) - h) // 2
    guiObj.Show("x" . x . " y" . y . " w" . w . " h" . h)
}

Finance_Notify(msg, ms := 1800, accent := "") {
    if (accent = "")
        accent := BANNER_ACCENT_INFO
    try ShowCenteredOverlay_Utils(msg, ms, accent)
    catch {
        TrayTip("Finance", msg)
    }
}

Finance_CollectNotifications() {
    notes := []
    ym := Finance_CurrentYearMonth()
    Finance_RecomputeBudgetSpent(ym)
    if (Finance_Setting("General", "NotifyBudgetExceeded", "1") = "1") {
        cats := Finance_Load("categories")
        for b in Finance_Load("budgets") {
            if (b["year_month"] != ym)
                continue
            planned := Finance_ParseDecimal(b["planned_amount"])
            spent := Finance_ParseDecimal(b["spent_amount"])
            if (planned > 0 && spent > planned) {
                name := Finance_CatName(cats, b["category_id"])
                notes.Push("Budget exceeded: " . name . " (" . Finance_FormatBrl(spent) . " / " . Finance_FormatBrl(
                    planned) . ")")
            }
        }
    }
    if (Finance_Setting("General", "NotifyCardHighUsage", "1") = "1") {
        warn := Finance_ParseDecimal(Finance_Setting("General", "CardUsageWarnPct", "80"))
        for c in Finance_Load("credit_cards") {
            lim := Finance_ParseDecimal(c["limit"])
            spent := Finance_ParseDecimal(c["current_spent"])
            if (lim > 0 && (spent / lim) * 100 >= warn)
                notes.Push("Card " . c["name"] . " at " . Format("{:.0f}", (spent / lim) * 100) . "% of limit")
        }
    }
    today := Finance_Today()
    for g in Finance_Load("goals") {
        if (g["status"] = "in_progress" && g["target_date"] != "" && StrCompare(g["target_date"], today) < 0)
            notes.Push("Goal expired: " . g["name"])
    }
    return notes
}

Finance_ComboFromRows(rows, idKey := "id", nameKey := "name", includeEmpty := false) {
    names := []
    ids := []
    if (includeEmpty) {
        names.Push("(none)")
        ids.Push("")
    }
    for r in rows {
        names.Push(r[nameKey])
        ids.Push(r[idKey])
    }
    return { names: names, ids: ids }
}

Finance_ComboIndex(ids, id) {
    loop ids.Length {
        if (ids[A_Index] = id)
            return A_Index
    }
    return 1
}

; --- seed / migration ---

Finance_IniSections(path) {
    sections := []
    text := Finance_ReadUtf8(path)
    if (text = "")
        return sections
    loop parse text, "`n", "`r" {
        line := Trim(A_LoopField)
        if (SubStr(line, 1, 1) = "[" && SubStr(line, -1) = "]")
            sections.Push(SubStr(line, 2, StrLen(line) - 2))
    }
    return sections
}

Finance_MigrateCategoriesFromIni() {
    rows := []
    palI := 1
    expensePath := A_ScriptDir . "\categories-expenses.ini"
    incomePath := A_ScriptDir . "\categories-income.ini"
    if (FileExist(expensePath)) {
        for section in Finance_IniSections(expensePath) {
            mainId := Finance_SlugId("CAT_", section, rows)
            rows.Push(Map("id", mainId, "name", section, "type", "expense", "parent_id", "",
                "color", Finance_ColorForIndex(palI), "icon", "tag"))
            palI += 1
            raw := IniRead(expensePath, section)
            if (raw = "ERROR")
                raw := ""
            keys := StrSplit(raw, "`n")
            for k in keys {
                key := Trim(StrSplit(k, "=")[1])
                if (key = "" || key = "geral")
                    continue
                subId := Finance_SlugId("CAT_", key, rows)
                rows.Push(Map("id", subId, "name", key, "type", "expense", "parent_id", mainId,
                    "color", Finance_ColorForIndex(palI), "icon", "dot"))
                palI += 1
            }
        }
    }
    if (FileExist(incomePath)) {
        for section in Finance_IniSections(incomePath) {
            mainId := Finance_SlugId("CAT_", section, rows)
            rows.Push(Map("id", mainId, "name", section, "type", "income", "parent_id", "",
                "color", Finance_ColorForIndex(palI), "icon", "arrow-down"))
            palI += 1
        }
    }
    if (rows.Length = 0) {
        rows.Push(Map("id", "CAT_OUTROS", "name", "Outros", "type", "expense", "parent_id", "",
            "color", "#7F8C8D", "icon", "tag"))
        rows.Push(Map("id", "CAT_SALARIO", "name", "Salário", "type", "income", "parent_id", "",
            "color", "#3498DB", "icon", "briefcase"))
    }
    Finance_Save("categories", rows)
}

Finance_MigrateAccountsFromIni() {
    rows := []
    path := A_ScriptDir . "\accounts.ini"
    icons := Map(
        "BoschLife", "🏦",
        "FGTS", "🏛️",
        "Meal voucher", "🍽️",
        "Meli dólares", "💵",
        "Mercado Pago long-term", "📈",
        "Mercado Pago main account", "💳",
        "Mercado Pago short-term", "💰",
        "Nubank Main", "🟣",
        "Transition money", "🔄"
    )
    balances := Map(
        "BoschLife", "23481,62",
        "FGTS", "54104,01",
        "Meal voucher", "0,00",
        "Meli dólares", "0,00",
        "Mercado Pago long-term", "43265,02",
        "Mercado Pago main account", "48004,95",
        "Mercado Pago short-term", "5000,00",
        "Nubank Main", "5000,00",
        "Transition money", "0,00"
    )
    names := FileExist(path) ? Finance_IniSections(path) : ["Mercado Pago main account", "Nubank Main"]
    for name in names {
        id := Finance_SlugId("ACC_", name, rows)
        icon := icons.Has(name) ? icons[name] : "🏦"
        bal := balances.Has(name) ? balances[name] : "0,00"
        ; long-term name in INI is without "s..."
        if (InStr(name, "long-term")) {
            icon := "📈"
            if (!balances.Has(name))
                bal := "43265,02"
        }
        rows.Push(Map("id", id, "name", name, "icon", icon, "initial_balance", bal, "current_balance", bal))
    }
    Finance_Save("accounts", rows)
}

Finance_SeedCreditCards() {
    accs := Finance_Load("accounts")
    linked := "ACC_MPMAIN"
    for a in accs {
        if (InStr(a["name"], "Mercado Pago main")) {
            linked := a["id"]
            break
        }
    }
    rows := []
    rows.Push(Map("id", "CARD_MP", "name", "Mercado Pago", "limit", "12000,00",
        "current_spent", "2010,22", "linked_account_id", linked, "closing_day", "9"))
    Finance_Save("credit_cards", rows)
}

Finance_SeedGoals() {
    rows := []
    rows.Push(Map("id", "GOAL_PREV", "name", "Previdência privada", "current_amount", "16733,13",
        "target_amount", "926400,00", "target_date", "2062-01-01", "status", "in_progress"))
    rows.Push(Map("id", "GOAL_EMERG", "name", "Fundo de emergência", "current_amount", "2141,14",
        "target_amount", "18000,00", "target_date", "2099-11-01", "status", "in_progress"))
    rows.Push(Map("id", "GOAL_FATHER", "name", "Father's money", "current_amount", "10000,00",
        "target_amount", "10000,00", "target_date", "2026-09-15", "status", "in_progress"))
    rows.Push(Map("id", "GOAL_ALEM", "name", "Alemanha", "current_amount", "22314,00",
        "target_amount", "50000,00", "target_date", "", "status", "paused"))
    rows.Push(Map("id", "GOAL_ALEM2", "name", "Alemanha (parte Leonardo)", "current_amount", "500,00",
        "target_amount", "15000,00", "target_date", "", "status", "paused"))
    rows.Push(Map("id", "GOAL_CARRO", "name", "Novo carro", "current_amount", "600,00",
        "target_amount", "40000,00", "target_date", "2026-01-01", "status", "expired"))
    rows.Push(Map("id", "GOAL_TERR", "name", "Entrada terreno", "current_amount", "7218,73",
        "target_amount", "40000,00", "target_date", "2026-01-01", "status", "expired"))
    rows.Push(Map("id", "GOAL_NCAR", "name", "New car", "current_amount", "32000,00",
        "target_amount", "70000,00", "target_date", "2026-01-01", "status", "expired"))
    Finance_Save("goals", rows)
}

Finance_CatIdByName(name) {
    for c in Finance_Load("categories") {
        if (c["name"] = name && c["parent_id"] = "")
            return c["id"]
    }
    return ""
}

Finance_AccIdByNameContains(needle) {
    for a in Finance_Load("accounts") {
        if (InStr(a["name"], needle))
            return a["id"]
    }
    return ""
}

Finance_SeedBudgets() {
    ym := "2026-08"
    pairs := [["Mercado", "3000,00", "1039,26"], ["Alimentação", "300,00", "603,43"], ["Beleza", "150,00", "100,00"],
    ["Lazer", "600,00", "167,28"], ["Saúde", "1000,00", "670,27"], ["Carro", "400,00", "696,59"],
    ["Contas de casa", "2300,00", "0,00"], ["Educação", "400,00", "139,77"], ["Eletrônicos", "50,00", "733,44"],
    ["Humanitário", "150,00", "350,90"]]
    rows := []
    for p in pairs {
        cid := Finance_CatIdByName(p[1])
        if (cid = "")
            continue
        rows.Push(Map("year_month", ym, "category_id", cid, "planned_amount", p[2], "spent_amount", p[3]))
    }
    Finance_Save("budgets", rows)
}

Finance_SeedTransactions() {
    accBl := Finance_AccIdByNameContains("BoschLife")
    accMp := Finance_AccIdByNameContains("Mercado Pago main")
    if (accMp = "")
        accMp := Finance_AccIdByNameContains("Mercado Pago")
    catMer := Finance_CatIdByName("Mercado")
    catAli := Finance_CatIdByName("Alimentação")
    catSal := Finance_CatIdByName("Salário")
    catBon := Finance_CatIdByName("Bonificação")
    catEle := Finance_CatIdByName("Eletrônicos")
    rows := []
    rows.Push(Map("id", "TX001", "date", "2026-08-16", "description", "Banana", "amount", "3,00",
        "type", "expense", "category_id", catMer, "subcategory", "Hortifruti", "account_id", accBl, "card_id", "",
        "transfer_account_id", ""))
    rows.Push(Map("id", "TX002", "date", "2026-08-16", "description", "Gift", "amount", "10,00",
        "type", "income", "category_id", catBon, "subcategory", "", "account_id", accBl, "card_id", "",
        "transfer_account_id", ""))
    rows.Push(Map("id", "TX003", "date", "2026-08-15", "description", "Lunch", "amount", "31,24",
        "type", "expense", "category_id", catAli, "subcategory", "", "account_id", accMp, "card_id", "",
        "transfer_account_id", ""))
    rows.Push(Map("id", "TX004", "date", "2026-08-15", "description", "Groceries", "amount", "317,04",
        "type", "expense", "category_id", catMer, "subcategory", "", "account_id", accMp, "card_id", "",
        "transfer_account_id", ""))
    rows.Push(Map("id", "TX005", "date", "2026-08-15", "description", "Mic", "amount", "529,99",
        "type", "card_expense", "category_id", catEle, "subcategory", "", "account_id", accMp, "card_id", "CARD_MP",
        "transfer_account_id", ""))
    rows.Push(Map("id", "TX006", "date", "2026-08-03", "description", "Salary and PLR", "amount", "4876,76",
        "type", "income", "category_id", catSal, "subcategory", "", "account_id", accMp, "card_id", "",
        "transfer_account_id", ""))
    Finance_Save("transactions", rows)
}

Finance_PickList(title, labels) {
    result := { index: 0 }
    g := Gui("+AlwaysOnTop +ToolWindow", title)
    g.SetFont("s10", "Segoe UI")
    lv := g.Add("ListView", "w420 r12", ["#", "Name"])
    loop labels.Length
        lv.Add("", A_Index, labels[A_Index])
    lv.ModifyCol(1, 40)
    lv.ModifyCol(2, 360)
    g.Add("Button", "w90 Default", "OK").OnEvent("Click", (*) => (result.index := lv.GetNext() ? lv.GetNext() : 1, g.Destroy()))
    g.Add("Button", "x+8 w90", "Cancel").OnEvent("Click", (*) => (result.index := 0, g.Destroy()))
    g.OnEvent("Escape", (*) => (result.index := 0, g.Destroy()))
    lv.OnEvent("DoubleClick", (*) => (result.index := lv.GetNext() ? lv.GetNext() : 1, g.Destroy()))
    g.Show()
    try WinWaitClose("ahk_id " g.Hwnd)
    catch {
    }
    return result.index
}
