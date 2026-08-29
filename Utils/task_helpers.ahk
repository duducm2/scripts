; =============================================================================
; Utils module: task_helpers.ahk
; Tasks — CSV I/O, paths, GUI chrome, settings, recurrence helpers
; =============================================================================

global g_TaskGui := false
global g_TaskHotkeys := []
global g_TaskFilter := "all"
global g_TaskBrowseProjectId := ""
global g_TaskBrowseTaskId := ""
global g_TaskHabitsView := false
global g_TaskLetterJump := ""
global g_TaskLv := false
global g_TaskRows := []

Task_DataDir() {
    dir := A_ScriptDir . "\tasks\data"
    if (!DirExist(dir))
        DirCreate(dir)
    imported := dir . "\imported"
    if (!DirExist(imported))
        DirCreate(imported)
    attach := dir . "\attachments"
    if (!DirExist(attach))
        DirCreate(attach)
    return dir
}

Task_AttachmentsDir() {
    return Task_DataDir() . "\attachments"
}

; Delete a managed file under tasks/data/attachments (image/text sidecars only).
; Skips urls and absolute paths outside the attachments folder.
Task_DeleteManagedAttachmentFile(att) {
    if (!IsObject(att))
        return
    kind := att.Has("kind") ? att["kind"] : ""
    ref := Trim(att.Has("ref") ? att["ref"] : "")
    if (ref = "")
        return
    if (kind != "image" && kind != "text")
        return
    path := ""
    norm := StrReplace(ref, "/", "\")
    if (InStr(norm, "attachments\") = 1)
        path := Task_DataDir() . "\" . norm
    else {
        attachRoot := Task_AttachmentsDir()
        if (InStr(norm, attachRoot . "\") = 1 || norm = attachRoot)
            path := norm
        else
            return
    }
    if (path != "" && FileExist(path)) {
        try FileDelete(path)
        catch {
        }
    }
}

; Drop matching attachment rows and delete their managed files from disk.
Task_PurgeAttachments(shouldDrop) {
    attOut := []
    for a in Task_Load("attachments") {
        drop := false
        try drop := shouldDrop(a)
        catch {
            drop := false
        }
        if (drop)
            Task_DeleteManagedAttachmentFile(a)
        else
            attOut.Push(a)
    }
    Task_Save("attachments", attOut)
}

Task_OutputDir() {
    dir := A_ScriptDir . "\tasks\output"
    if (!DirExist(dir))
        DirCreate(dir)
    return dir
}

Task_PythonDir() {
    return A_ScriptDir . "\tasks\python"
}

Task_SettingsPath() {
    return Task_DataDir() . "\settings.ini"
}

Task_EnsureSettings() {
    path := Task_SettingsPath()
    if (FileExist(path))
        return
    content := "[General]`n"
        . "Filter=work`n"
        . "LastProjectId=`n"
        . "DashboardChromeHwnd=`n"
        . "`n[Dashboard]`n"
        . "ShowByFilter=1`n"
        . "ShowByEmoji=1`n"
        . "ShowHabitsUpcoming=1`n"
        . "ShowCompletedRecent=1`n"
    Task_WriteUtf8(path, content)
}

Task_Setting(section, key, default := "") {
    val := IniRead(Task_SettingsPath(), section, key, default)
    if (val = "ERROR")
        return default
    return val
}

Task_SetSetting(section, key, value) {
    IniWrite(value, Task_SettingsPath(), section, key)
}

Task_FindPythonCmd() {
    candidates := ["py -3", "py", "python3", "python"]
    for c in candidates {
        try {
            ec := RunWait(A_ComSpec . ' /c ' . c . ' -c "print(1)" >nul 2>&1', , "Hide")
            if (ec = 0)
                return c
        } catch {
        }
    }
    localApps := EnvGet("LOCALAPPDATA")
    pathGlobs := [
        localApps . "\Programs\Python\Python3*\python.exe",
        EnvGet("ProgramFiles") . "\Python3*\python.exe",
        "C:\Python3*\python.exe"
    ]
    for g in pathGlobs {
        loop files g, "F" {
            try {
                ec := RunWait('"' . A_LoopFileFullPath . '" -c "print(1)"', , "Hide")
                if (ec = 0)
                    return '"' . A_LoopFileFullPath . '"'
            } catch {
            }
        }
    }
    return ""
}

Task_WriteUtf8(path, content) {
    f := FileOpen(path, "w", "UTF-8")
    if (!f)
        throw Error("Could not write " . path)
    f.Write(content)
    f.Close()
}

Task_ReadUtf8(path) {
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

Task_SplitCsvLine(line) {
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

Task_CsvEscape(val) {
    s := String(val)
    if (InStr(s, ",") || InStr(s, '"') || InStr(s, "`n") || InStr(s, "`r"))
        return '"' . StrReplace(s, '"', '""') . '"'
    return s
}

Task_ReadCsv(fileName) {
    path := InStr(fileName, "\") || InStr(fileName, "/") ? fileName : Task_DataDir() . "\" . fileName
    rows := []
    text := Task_ReadUtf8(path)
    if (text = "")
        return rows
    text := StrReplace(text, "`r`n", "`n")
    text := StrReplace(text, "`r", "`n")
    lines := StrSplit(text, "`n")
    headers := []
    for idx, line in lines {
        if (Trim(line) = "")
            continue
        fields := Task_SplitCsvLine(line)
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

Task_WriteCsv(fileName, rows, headers) {
    path := InStr(fileName, "\") || InStr(fileName, "/") ? fileName : Task_DataDir() . "\" . fileName
    out := ""
    loop headers.Length {
        if (A_Index > 1)
            out .= ","
        out .= Task_CsvEscape(headers[A_Index])
    }
    out .= "`n"
    for row in rows {
        loop headers.Length {
            if (A_Index > 1)
                out .= ","
            key := headers[A_Index]
            val := row.Has(key) ? row[key] : ""
            out .= Task_CsvEscape(val)
        }
        out .= "`n"
    }
    Task_WriteUtf8(path, out)
}

Task_Headers(kind) {
    switch kind {
        case "projects":
            return ["id", "title", "filter", "section_path", "sort_order", "active", "created_at"]
        case "tasks":
            return ["id", "project_id", "title", "emoji", "kind", "recurrence", "due_date", "next_due",
                "section_path", "filter", "sort_order", "completed_at", "created_at", "active"]
        case "info_points":
            return ["id", "parent_type", "parent_id", "title", "body", "emoji", "section_path", "sort_order",
                "created_at"]
        case "attachments":
            return ["id", "parent_type", "parent_id", "kind", "ref", "description", "sort_order"]
        default:
            return []
    }
}

Task_Save(kind, rows) {
    Task_WriteCsv(kind . ".csv", rows, Task_Headers(kind))
}

Task_Load(kind) {
    return Task_ReadCsv(kind . ".csv")
}

Task_EnsureData() {
    Task_DataDir()
    Task_EnsureSettings()
    for kind in ["projects", "tasks", "info_points", "attachments"] {
        path := Task_DataDir() . "\" . kind . ".csv"
        if (!FileExist(path)) {
            hdrs := Task_Headers(kind)
            line := ""
            loop hdrs.Length {
                if (A_Index > 1)
                    line .= ","
                line .= hdrs[A_Index]
            }
            Task_WriteUtf8(path, line . "`n")
        }
    }
    global g_TaskFilter
    g_TaskFilter := Task_Setting("General", "Filter", "work")
    if (g_TaskFilter = "" || g_TaskFilter = "all")
        g_TaskFilter := "work"
}

Task_FindById(rows, id) {
    for row in rows {
        if (row["id"] = id)
            return row
    }
    return false
}

Task_NextId(prefix, rows, pad := 4) {
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

Task_NextSortOrder(rows) {
    maxN := 0
    for row in rows {
        so := row.Has("sort_order") ? Integer(row["sort_order"] || 0) : 0
        if (so > maxN)
            maxN := so
    }
    return String(maxN + 10)
}

Task_Today() {
    return FormatTime(, "yyyy-MM-dd")
}

Task_NowStamp() {
    return FormatTime(, "yyyy-MM-dd HH:mm:ss")
}

Task_FilterLabels() {
    return ["work", "personal", "habits"]
}

Task_FilterLabel(f) {
    switch f {
        case "work":
            return "Work"
        case "personal":
            return "Personal"
        case "habits":
            return "Habits"
        default:
            return "Work"
    }
}

Task_SetFilter(filt) {
    global g_TaskFilter
    filt := StrLower(Trim(filt))
    if (filt != "work" && filt != "personal" && filt != "habits")
        filt := "work"
    g_TaskFilter := filt
    Task_SetSetting("General", "Filter", filt)
    return filt
}

Task_MatchesFilter(row) {
    global g_TaskFilter
    fWant := g_TaskFilter
    if (fWant = "" || fWant = "all")
        fWant := "work"
    f := row.Has("filter") ? row["filter"] : ""
    return (f = fWant)
}

Task_DefaultEmoji() {
    return "🔲"
}

Task_InfoEmoji() {
    return "ℹ️"
}

Task_StatusEmojis() {
    return Map(
        "general", "🔲",
        "waiting", "⏳",
        "important", "⚡",
        "done", "✅",
        "doubt", "❓",
        "info", "ℹ️"
    )
}

Task_IsOpenEmoji(emoji) {
    return (emoji != "✅")
}

Task_AdvanceNextDue(recurrence, fromDate := "") {
    base := fromDate != "" ? fromDate : Task_Today()
    if (!RegExMatch(base, "^(\d{4})-(\d{2})-(\d{2})", &m))
        base := Task_Today()
    y := Integer(SubStr(base, 1, 4))
    mo := Integer(SubStr(base, 6, 2))
    d := Integer(SubStr(base, 9, 2))
    days := 0
    addMonths := 0
    addYears := 0
    switch recurrence {
        case "daily":
            days := 1
        case "weekly":
            days := 7
        case "monthly":
            addMonths := 1
        case "quarterly":
            addMonths := 3
        case "biannual":
            addMonths := 6
        case "yearly":
            addYears := 1
        case "every_2y":
            addYears := 2
        case "every_3y":
            addYears := 3
        case "every_5y":
            addYears := 5
        case "every_10y":
            addYears := 10
        default:
            return base
    }
    if (days) {
        stamp := Format("{:04}{:02}{:02}000000", y, mo, d)
        stamp := DateAdd(stamp, days, "Days")
        return FormatTime(stamp, "yyyy-MM-dd")
    }
    y += addYears
    mo += addMonths
    while (mo > 12) {
        mo -= 12
        y += 1
    }
    while (mo < 1) {
        mo += 12
        y -= 1
    }
    leap := (Mod(y, 4) = 0 && (Mod(y, 100) != 0 || Mod(y, 400) = 0))
    daysInMonth := [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    maxD := daysInMonth[mo]
    if (d > maxD)
        d := maxD
    return Format("{:04}-{:02}-{:02}", y, mo, d)
}

Task_CompleteTask(row) {
    emojiDone := Task_StatusEmojis()["done"]
    if (row["kind"] = "habitual") {
        from := Trim(row["next_due"]) != "" ? row["next_due"] : Task_Today()
        row["next_due"] := Task_AdvanceNextDue(row["recurrence"], from)
        row["completed_at"] := Task_NowStamp()
        row["emoji"] := Task_DefaultEmoji()
    } else {
        row["emoji"] := emojiDone
        row["completed_at"] := Task_NowStamp()
    }
    return row
}

Task_CloseGui() {
    global g_TaskGui, g_TaskHotkeys, g_TaskLetterJump
    Task_LetterJumpStop()
    Task_UnbindHotkeys()
    if (IsObject(g_TaskGui)) {
        try g_TaskGui.Destroy()
        catch {
        }
    }
    g_TaskGui := false
}

Task_UnbindHotkeys() {
    global g_TaskHotkeys
    try HotIf(Task_HotIfTaskKeys)
    catch {
    }
    for item in g_TaskHotkeys {
        try Hotkey(item, "Off")
        catch {
        }
        ; Also clear undecorated form if we bound $Backspace
        if (SubStr(item, 1, 1) = "$") {
            try Hotkey(SubStr(item, 2), "Off")
            catch {
            }
        }
    }
    g_TaskHotkeys := []
    try HotIf()
    catch {
    }
}

Task_GuiFocusIsEdit() {
    global g_TaskGui
    if (!IsObject(g_TaskGui))
        return false
    try {
        focused := ControlGetFocus("ahk_id " g_TaskGui.Hwnd)
        if (focused = "")
            return false
        ; Edit / RichEdit only — ListView must keep Backspace as "up"
        if (InStr(focused, "Edit") = 1)
            return true
        if (InStr(focused, "RichEdit") = 1)
            return true
        return false
    } catch {
        return false
    }
}

Task_HotIfTaskKeys(*) {
    global g_TaskGui
    if (!IsObject(g_TaskGui))
        return false
    try {
        if (!WinActive("ahk_id " g_TaskGui.Hwnd))
            return false
        return !Task_GuiFocusIsEdit()
    } catch {
        return false
    }
}

; Run navigation after the hotkey thread ends so CloseGui can rebind Backspace safely.
global g_TaskNavPending := false
global g_TaskNavFn := ""

Task_DeferNav(fn) {
    global g_TaskNavPending, g_TaskNavFn
    g_TaskNavFn := fn
    if (g_TaskNavPending)
        return
    g_TaskNavPending := true
    SetTimer(Task_RunDeferredNav, -15)
}

Task_RunDeferredNav(*) {
    global g_TaskNavPending, g_TaskNavFn
    g_TaskNavPending := false
    fn := g_TaskNavFn
    g_TaskNavFn := ""
    try {
        if (fn)
            fn.Call()
    } catch {
    }
}

Task_BindHotkeys(pairs) {
    global g_TaskGui, g_TaskHotkeys
    Task_UnbindHotkeys()
    if (!IsObject(g_TaskGui))
        return
    try HotIf(Task_HotIfTaskKeys)
    catch {
        return
    }
    for p in pairs {
        key := p[1]
        fn := p[2]
        ; Only Backspace needs deferral (Escape often also arrives via Gui.OnEvent).
        if (key = "Backspace")
            fn := Task_WrapNav(fn)
        bindKey := key
        if (key = "Backspace")
            bindKey := "$Backspace"
        try {
            Hotkey(bindKey, fn, "On")
            g_TaskHotkeys.Push(bindKey)
        } catch {
            try {
                Hotkey(key, fn, "On")
                g_TaskHotkeys.Push(key)
            } catch {
            }
        }
    }
    try HotIf()
    catch {
    }
}

Task_WrapNav(fn) {
    return (*) => Task_DeferNav(fn)
}

Task_ActivateListGui() {
    global g_TaskGui, g_TaskLv
    try {
        if (IsObject(g_TaskGui))
            WinActivate("ahk_id " g_TaskGui.Hwnd)
    } catch {
    }
    try {
        if (IsObject(g_TaskLv))
            g_TaskLv.Focus()
    } catch {
    }
}

Task_CenterGui(guiObj, w := 920, h := 620) {
    MonitorGetWorkArea(MonitorGetPrimary(), &L, &T, &R, &B)
    x := L + ((R - L) - w) // 2
    y := T + ((B - T) - h) // 2
    guiObj.Show("x" . x . " y" . y . " w" . w . " h" . h)
    Task_ActivateListGui()
}

Task_Notify(msg, ms := 1800, accent := "") {
    if (accent = "")
        accent := BANNER_ACCENT_INFO
    try ShowCenteredOverlay_Utils(msg, ms, accent)
    catch {
        TrayTip("Tasks", msg)
    }
}

Task_DialogsBegin() {
    global g_TaskGui
    try {
        if (IsObject(g_TaskGui))
            g_TaskGui.Opt("-AlwaysOnTop")
    } catch {
    }
}

Task_DialogsEnd() {
    global g_TaskGui
    try {
        if (IsObject(g_TaskGui))
            g_TaskGui.Opt("+AlwaysOnTop")
    } catch {
    }
}

Task_OwnerOpt() {
    global g_TaskGui
    hwnd := 0
    try {
        if (IsObject(g_TaskGui))
            hwnd := g_TaskGui.Hwnd
    } catch {
        hwnd := 0
    }
    return hwnd ? " Owner" . hwnd : ""
}

Task_Confirm(msg, title := "Tasks") {
    Task_DialogsBegin()
    result := MsgBox(msg, title, "YesNo Icon?" . Task_OwnerOpt())
    Task_DialogsEnd()
    return result = "Yes"
}

Task_Alert(msg, title := "Tasks") {
    Task_DialogsBegin()
    MsgBox(msg, title, "Icon!" . Task_OwnerOpt())
    Task_DialogsEnd()
}

Task_InputBox(prompt, title := "Tasks", default := "") {
    Task_DialogsBegin()
    result := InputBox(prompt, title, "w360" . Task_OwnerOpt(), default)
    Task_DialogsEnd()
    return result
}

Task_StyleDarkListView(lv) {
    hwnd := 0
    try hwnd := lv.Hwnd
    catch {
        return
    }
    if (!hwnd)
        return
    try DllCall("uxtheme\SetWindowTheme", "ptr", hwnd, "wstr", "", "wstr", "")
    catch {
    }
    bg := 0x302D2D
    fg := 0xF2F2F2
    try SendMessage(0x1001, 0, bg, hwnd)
    catch {
    }
    try SendMessage(0x1024, 0, fg, hwnd)
    catch {
    }
    try SendMessage(0x1026, 0, bg, hwnd)
    catch {
    }
    hdr := 0
    try hdr := SendMessage(0x101F, 0, 0, hwnd)
    catch {
        hdr := 0
    }
    if (hdr) {
        try DllCall("uxtheme\SetWindowTheme", "ptr", hdr, "wstr", "DarkMode_ItemsView", "wstr", "")
        catch {
        }
    }
}

Task_AddBrowseChrome(guiObj, levelNoun) {
    global g_TaskFilter, g_TaskBrowseProjectId
    crumb := levelNoun
    if (g_TaskBrowseProjectId != "") {
        p := Task_FindById(Task_Load("projects"), g_TaskBrowseProjectId)
        if (IsObject(p))
            crumb := p["title"] . " › " . levelNoun
    }
    filt := Task_FilterLabel(g_TaskFilter)
    try guiObj.BackColor := "1E1E1E"
    catch {
    }
    try DllCall("dwmapi\DwmSetWindowAttribute", "ptr", guiObj.Hwnd, "uint", 20, "int*", 1, "int", 4)
    catch {
    }
    guiObj.SetFont("s11 cF1C40F Bold", "Segoe UI")
    guiObj.Add("Text", "x12 y8 w860 cF1C40F BackgroundTrans", crumb . "  ·  filter: " . filt)
    guiObj.SetFont("s9 cA0A0A0 Norm", "Segoe UI")
    guiObj.Add("Text", "x12 y32 w860 cA0A0A0 BackgroundTrans",
        "1 Work  2 Personal  3 Habits   A add  E edit  Del  Enter drill  Backspace up  C/I/U/G/W emoji  V paste  N info"
    )
    guiObj.SetFont("s10 cWhite Norm", "Segoe UI")
    return 56
}

Task_LetterJumpStop() {
    global g_TaskLetterJump
    if (!IsObject(g_TaskLetterJump) || !g_TaskLetterJump.HasProp("chars")) {
        g_TaskLetterJump := ""
        return
    }
    try HotIf(Task_HotIfTaskKeys)
    catch {
    }
    for ch in g_TaskLetterJump.chars {
        try Hotkey(ch, "Off")
        catch {
        }
        try Hotkey(StrUpper(ch), "Off")
        catch {
        }
    }
    try HotIf()
    catch {
    }
    g_TaskLetterJump := ""
}

Task_LetterJumpStart(getNameFn) {
    global g_TaskLetterJump, g_TaskLv, g_TaskRows, g_TaskHotkeys
    Task_LetterJumpStop()
    ; Bind under same HotIf as action keys so inactive Tasks windows do not swallow a-z.
    chars := []
    try HotIf(Task_HotIfTaskKeys)
    catch {
        return
    }
    loop 26 {
        ch := Chr(96 + A_Index)
        ; Skip keys already bound as actions (a/e/c/w/i/u/g/x/n/v/f/…)
        skip := false
        for hk in g_TaskHotkeys {
            if (StrLower(hk) = ch) {
                skip := true
                break
            }
        }
        if (skip)
            continue
        handler := Task_LetterJumpMakeHandler(ch, getNameFn)
        try {
            Hotkey(ch, handler, "On")
            Hotkey(StrUpper(ch), handler, "On")
            chars.Push(ch)
        } catch {
        }
    }
    try HotIf()
    catch {
    }
    g_TaskLetterJump := { chars: chars }
}

Task_LetterJumpMakeHandler(char, getNameFn) {
    return (*) => Task_LetterJumpHandle(char, getNameFn)
}

Task_LetterJumpHandle(char, getNameFn) {
    global g_TaskLv, g_TaskRows
    if (!IsObject(g_TaskLv) || !IsObject(g_TaskRows))
        return
    rowNum := ModalList_FindFirstByStartingLetter(g_TaskRows, char, getNameFn)
    if (rowNum > 0)
        ListView_SelectRowFocused(g_TaskLv, rowNum)
}

Task_ClipboardHasImage() {
    try {
        if (DllCall("IsClipboardFormatAvailable", "UInt", 8, "Int")) ; CF_DIB
            return true
        if (DllCall("IsClipboardFormatAvailable", "UInt", 2, "Int")) ; CF_BITMAP
            return true
    } catch {
    }
    return false
}

Task_SaveClipboardImage(destPath) {
    ps := "Add-Type -AssemblyName System.Windows.Forms; "
        . "Add-Type -AssemblyName System.Drawing; "
        . "$img = [System.Windows.Forms.Clipboard]::GetImage(); "
        . "if ($null -eq $img) { exit 2 }; "
        . "$img.Save('" . StrReplace(destPath, "'", "''") . "', [System.Drawing.Imaging.ImageFormat]::Png); "
        . "exit 0"
    try {
        ec := RunWait(A_ComSpec . ' /c powershell -NoProfile -ExecutionPolicy Bypass -Command "' . ps . '"', , "Hide")
        return (ec = 0 && FileExist(destPath))
    } catch {
        return false
    }
}

Task_CountOpenTasks() {
    n := 0
    for t in Task_Load("tasks") {
        if (!Task_MatchesFilter(t))
            continue
        if (t.Has("active") && t["active"] = "0")
            continue
        if (t.Has("kind") && t["kind"] = "habitual")
            continue
        if (Task_IsOpenEmoji(t["emoji"]))
            n += 1
    }
    return n
}

Task_HabitIsDue(t) {
    if (!IsObject(t) || t["kind"] != "habitual")
        return false
    due := Trim(t["next_due"]) != "" ? Trim(t["next_due"]) : Trim(t["due_date"])
    if (due = "")
        return true
    return due <= Task_Today()
}

Task_FindProjectForFilter(filt) {
    projects := Task_Load("projects")
    inboxId := ""
    firstId := ""
    for p in projects {
        if (p.Has("active") && p["active"] = "0")
            continue
        if (p["filter"] != filt)
            continue
        if (firstId = "")
            firstId := p["id"]
        title := StrLower(p["title"])
        if (InStr(title, "inbox") || InStr(title, "personal inbox") || InStr(title, "work inbox"))
            return p["id"]
    }
    if (firstId != "")
        return firstId
    ; Create a seed inbox project for that filter
    row := Map(
        "id", Task_NextId("PROJ_", projects),
        "title", (filt = "work" ? "Work inbox" : (filt = "personal" ? "Personal inbox" : "Habits inbox")),
        "filter", filt,
        "section_path", "",
        "sort_order", Task_NextSortOrder(projects),
        "active", "1",
        "created_at", Task_NowStamp()
    )
    projects.Push(row)
    Task_Save("projects", projects)
    return row["id"]
}

; Spawn a punctual copy under Work/Personal; keep habit; advance next_due.
Task_SendHabitToFilter(habit, filt) {
    if (!IsObject(habit) || habit["kind"] != "habitual") {
        Task_Notify("Select a habit", 1200, BANNER_ACCENT_ERROR)
        return false
    }
    if (filt != "work" && filt != "personal") {
        Task_Notify("Target must be work or personal", 1500, BANNER_ACCENT_ERROR)
        return false
    }
    projId := Task_FindProjectForFilter(filt)
    tasks := Task_Load("tasks")
    copyId := Task_NextId("TASK_", tasks)
    copy := Map(
        "id", copyId,
        "project_id", projId,
        "title", habit["title"],
        "emoji", Task_DefaultEmoji(),
        "kind", "punctual",
        "recurrence", "",
        "due_date", Task_Today(),
        "next_due", "",
        "section_path", habit["section_path"],
        "filter", filt,
        "sort_order", Task_NextSortOrder(tasks),
        "completed_at", "",
        "created_at", Task_NowStamp(),
        "active", "1"
    )
    from := Trim(habit["next_due"]) != "" ? habit["next_due"] : Task_Today()
    habitNext := Task_AdvanceNextDue(habit["recurrence"], from)
    out := []
    for t in tasks {
        if (t["id"] = habit["id"]) {
            t["next_due"] := habitNext
            t["completed_at"] := Task_NowStamp()
            out.Push(t)
        } else {
            out.Push(t)
        }
    }
    out.Push(copy)
    Task_Save("tasks", out)

    infos := Task_Load("info_points")
    infos.Push(Map(
        "id", Task_NextId("INFO_", infos),
        "parent_type", "task",
        "parent_id", copyId,
        "title", "From habit",
        "body", "Spawned from habit " . habit["id"] . " (" . habit["recurrence"] . "). Next habit due: "
        . habitNext,
        "emoji", Task_InfoEmoji(),
        "section_path", "",
        "sort_order", Task_NextSortOrder(infos),
        "created_at", Task_NowStamp()
    ))
    Task_Save("info_points", infos)
    Task_Notify("Sent to " . filt . " · habit next " . habitNext, 2200, BANNER_ACCENT_SUCCESS)
    return true
}

Task_Terms() {
    return [
        ["Project", "Container for tasks. Flat list; MD headings become section_path on children."],
        ["Task", "Actionable punctual item. Habits are separate (kind=habitual) and do not count as open."],
        ["Habit", "Recurring item (kind=habitual). Track due items on the Tasks dashboard."],
        ["Filter", "Browse starts on Work. While browsing: 1=Work  2=Personal  3=Habits."],
        ["Emoji", "🔲 general  ⏳ waiting  ⚡ important  ✅ done  ❓ doubt  ℹ️ info  or any custom."],
        ["Info point", "Non-actionable note attached to a project or task (large text)."],
        ["Attachment", "image | url | file | text refs under tasks/data/attachments or absolute paths."]
    ]
}
