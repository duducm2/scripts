; =============================================================================
; Utils module: hotstring_selector_handlers_02.ahk
; Utility selector view switching, ListView populate, HotIfWinActive binds, CRUD
; =============================================================================

HandleUtilitySelectorBack(*) {
    global g_HotstringSelectorActive, g_UtilitySelectorMode
    if (!g_HotstringSelectorActive)
        return
    if (g_UtilitySelectorMode = "category")
        UtilitySelector_SwitchToTop()
}

UtilitySelector_SwitchToTop() {
    global g_UtilitySelectorMode, g_UtilitySelectorCategory
    g_UtilitySelectorMode := "top"
    g_UtilitySelectorCategory := ""
    UtilitySelector_RebuildGui()
}

UtilitySelector_SwitchToCategory(category) {
    global g_UtilitySelectorMode, g_UtilitySelectorCategory
    g_UtilitySelectorMode := "category"
    g_UtilitySelectorCategory := category
    UtilitySelector_RebuildGui()
}

UtilitySelector_HotkeyName(char) {
    if (char = ",")
        return "vkBC"
    if (char = ".")
        return "vkBE"
    return char
}

UtilitySelector_SelectorHwnd() {
    global g_HotstringSelectorGui
    hwnd := 0
    try {
        if (IsObject(g_HotstringSelectorGui))
            hwnd := g_HotstringSelectorGui.Hwnd
    } catch {
        hwnd := 0
    }
    return hwnd
}

UtilitySelector_UnbindModalHotkeys() {
    global g_HotstringSelectorGui, g_HotstringHotkeyHandlers, g_UtilitySelectorHotkeysBound
    hwnd := UtilitySelector_SelectorHwnd()
    if (hwnd) {
        try HotIfWinActive("ahk_id " hwnd)
        catch {
        }
    }
    for handler in g_HotstringHotkeyHandlers {
        try Hotkey(handler.key, "Off")
        catch {
        }
    }
    if (hwnd) {
        try HotIf()
        catch {
        }
    }
    g_HotstringHotkeyHandlers := []
    g_UtilitySelectorHotkeysBound := false
}

UtilitySelector_BindOneChar(char, handler) {
    global g_HotstringHotkeyHandlers
    key := UtilitySelector_HotkeyName(char)
    try {
        Hotkey(key, handler, "On")
        g_HotstringHotkeyHandlers.Push({ char: char, key: key, handler: handler })
    } catch {
    }
    if (RegExMatch(char, "^[a-z]$")) {
        upperKey := StrUpper(char)
        try {
            Hotkey(upperKey, handler, "On")
            g_HotstringHotkeyHandlers.Push({ char: char, key: upperKey, handler: handler })
        } catch {
        }
    }
}

UtilitySelector_BindModalHotkeys() {
    global g_HotstringSelectorGui, g_UtilitySelectorHotkeysBound, g_HotstringHotkeyHandlers
    global g_UtilitySelectorMode, g_UtilitySelectorCategory, g_UtilityTopCategoryById, g_UtilitySelectorRows
    UtilitySelector_UnbindModalHotkeys()
    hwnd := UtilitySelector_SelectorHwnd()
    if (!hwnd)
        return
    try HotIfWinActive("ahk_id " hwnd)
    catch {
        return
    }

    if (g_UtilitySelectorMode = "top") {
        for id, category in g_UtilityTopCategoryById {
            handler := CreateHotstringCharHandler(id)
            UtilitySelector_BindOneChar(id, handler)
        }
    } else {
        seen := Map()
        for row in g_UtilitySelectorRows {
            ch := row.HasProp("char") ? row.char : ""
            if (ch = "" || seen.Has(ch))
                continue
            seen[ch] := true
            UtilitySelector_BindOneChar(ch, CreateHotstringCharHandler(ch))
        }
        if (g_UtilitySelectorCategory = "Prompts") {
            UtilitySelector_BindOneChar("l", CreateHotstringCharHandler("l"))
        }
        try {
            Hotkey("Backspace", HandleUtilitySelectorBack, "On")
            g_HotstringHotkeyHandlers.Push({ char: "Backspace", key: "Backspace", handler: HandleUtilitySelectorBack })
        } catch {
        }
        if (g_UtilitySelectorCategory = "Prompts" || g_UtilitySelectorCategory = "Hotstrings") {
            try {
                Hotkey("Insert", UtilitySelector_OnAdd, "On")
                g_HotstringHotkeyHandlers.Push({ char: "Insert", key: "Insert", handler: UtilitySelector_OnAdd })
            } catch {
            }
            try {
                Hotkey("F2", UtilitySelector_OnEdit, "On")
                g_HotstringHotkeyHandlers.Push({ char: "F2", key: "F2", handler: UtilitySelector_OnEdit })
            } catch {
            }
            try {
                Hotkey("Delete", UtilitySelector_OnDelete, "On")
                g_HotstringHotkeyHandlers.Push({ char: "Delete", key: "Delete", handler: UtilitySelector_OnDelete })
            } catch {
            }
        }
    }

    try {
        Hotkey("Enter", UtilitySelector_OnListActivate, "On")
        g_HotstringHotkeyHandlers.Push({ char: "Enter", key: "Enter", handler: UtilitySelector_OnListActivate })
    } catch {
    }
    try {
        Hotkey("Escape", HandleHotstringEscape, "On")
        g_HotstringHotkeyHandlers.Push({ char: "Escape", key: "Escape", handler: HandleHotstringEscape })
    } catch {
    }

    try HotIf()
    catch {
    }
    g_UtilitySelectorHotkeysBound := true
}

UtilitySelector_HintText() {
    global g_UtilitySelectorMode, g_UtilitySelectorCategory
    if (g_UtilitySelectorMode = "top")
        return "Char = open category   Enter/double-click = open   Esc = close"
    if (g_UtilitySelectorCategory = "Prompts")
        return "Char = paste   Enter/double-click = paste   Insert = add   F2 = edit   Delete = remove   L = Gemini arm   Backspace = back   Esc = close"
    if (g_UtilitySelectorCategory = "Hotstrings")
        return "Char = paste   Enter/double-click = paste   Insert = add   F2 = edit   Delete = remove   Backspace = back   Esc = close"
    if (g_UtilitySelectorCategory = "Projects")
        return "Char = paste name   Enter/double-click = paste name   Backspace = back   Esc = close"
    return "Char = run   Enter/double-click = run   Backspace = back   Esc = close"
}

UtilitySelector_PopulateLv() {
    global g_HotstringSelectorLv, g_UtilitySelectorMode, g_UtilitySelectorCategory, g_UtilitySelectorRows
    global g_UtilityTopCategories, g_UtilityTopCategoryById
    if (!IsObject(g_HotstringSelectorLv))
        return
    g_HotstringSelectorLv.Delete()
    g_UtilitySelectorRows := []

    if (g_UtilitySelectorMode = "top") {
        counts := Map("Prompts", PromptData_Load().Length, "Projects", UtilitySelector_ProjectRows().Length,
        "Macros", UtilitySelector_MacroRows().Length, "Hotstrings", HotstringData_Load().Length)
        idByCat := Map()
        for id, cat in g_UtilityTopCategoryById
            idByCat[cat] := id
        for cat in g_UtilityTopCategories {
            ch := idByCat.Has(cat) ? idByCat[cat] : ""
            n := counts.Has(cat) ? counts[cat] : 0
            g_UtilitySelectorRows.Push({ char: ch, category: cat, count: n })
            g_HotstringSelectorLv.Add("", ch, cat, n)
        }
        try g_HotstringSelectorLv.ModifyCol(1, 50)
        try g_HotstringSelectorLv.ModifyCol(2, 220)
        try g_HotstringSelectorLv.ModifyCol(3, 80)
        return
    }

    if (g_UtilitySelectorCategory = "Prompts") {
        for prompt in PromptData_Sorted() {
            g_UtilitySelectorRows.Push(prompt)
            g_HotstringSelectorLv.Add("", prompt.char, prompt.category, prompt.name, prompt.filePath)
        }
        try g_HotstringSelectorLv.ModifyCol(1, 50)
        try g_HotstringSelectorLv.ModifyCol(2, 100)
        try g_HotstringSelectorLv.ModifyCol(3, 340)
        try g_HotstringSelectorLv.ModifyCol(4, 330)
        return
    }

    if (g_UtilitySelectorCategory = "Projects") {
        for project in UtilitySelector_ProjectRows() {
            g_UtilitySelectorRows.Push(project)
            g_HotstringSelectorLv.Add("", project.HasProp("char") ? project.char : "", project.name)
        }
        try g_HotstringSelectorLv.ModifyCol(1, 50)
        try g_HotstringSelectorLv.ModifyCol(2, 400)
        return
    }

    if (g_UtilitySelectorCategory = "Hotstrings") {
        HotstringData_Load()
        global g_HotstringEntries
        for item in g_HotstringEntries {
            g_UtilitySelectorRows.Push(item)
            g_HotstringSelectorLv.Add("", item.char, item.name, GetPreviewText(item.text, 80))
        }
        try g_HotstringSelectorLv.ModifyCol(1, 50)
        try g_HotstringSelectorLv.ModifyCol(2, 260)
        try g_HotstringSelectorLv.ModifyCol(3, 420)
        return
    }

    if (g_UtilitySelectorCategory = "Macros") {
        for row in UtilitySelector_MacroRows() {
            g_UtilitySelectorRows.Push(row)
            g_HotstringSelectorLv.Add("", row.char, row.title)
        }
        try g_HotstringSelectorLv.ModifyCol(1, 50)
        try g_HotstringSelectorLv.ModifyCol(2, 500)
    }
}

UtilitySelector_SelectedIndex() {
    global g_HotstringSelectorLv
    if (!IsObject(g_HotstringSelectorLv))
        return 0
    row := 0
    try row := g_HotstringSelectorLv.GetNext(0)
    catch {
        return 0
    }
    return row ? Integer(row) : 0
}

UtilitySelector_OnListActivate(*) {
    global g_UtilitySelectorMode, g_UtilitySelectorRows, g_UtilitySelectorCategory
    idx := UtilitySelector_SelectedIndex()
    if (idx < 1 || idx > g_UtilitySelectorRows.Length)
        return
    row := g_UtilitySelectorRows[idx]
    if (g_UtilitySelectorMode = "top") {
        if (row.HasProp("category") && row.category != "")
            UtilitySelector_SwitchToCategory(row.category)
        return
    }
    ch := row.HasProp("char") ? row.char : ""
    if (ch = "")
        return
    HandleHotstringChar(ch)
}

UtilitySelector_DialogsBegin() {
    global g_HotstringSelectorGui
    try {
        if (IsObject(g_HotstringSelectorGui))
            g_HotstringSelectorGui.Opt("-AlwaysOnTop")
    } catch {
    }
}

UtilitySelector_DialogsEnd() {
    global g_HotstringSelectorGui
    try {
        if (IsObject(g_HotstringSelectorGui))
            g_HotstringSelectorGui.Opt("+AlwaysOnTop")
    } catch {
    }
}

UtilitySelector_InputBox(prompt, title, width := 420, defaultVal := "") {
    return InputBox(prompt, title, "w" . width, defaultVal)
}

UtilitySelector_RefocusGui() {
    global g_HotstringSelectorGui, g_HotstringSelectorLv
    try {
        if (IsObject(g_HotstringSelectorGui))
            WinActivate("ahk_id " g_HotstringSelectorGui.Hwnd)
    } catch {
    }
    try {
        if (IsObject(g_HotstringSelectorLv))
            g_HotstringSelectorLv.Focus()
    } catch {
    }
}

UtilitySelector_Notify(msg) {
    ShowCenteredOverlay_Utils(msg, 1800, BANNER_ACCENT_ERROR)
}

UtilitySelector_AvailableChars(seqFn, isValidFn, entries, charProp := "char", excludeIndex := 0) {
    taken := Map()
    loop entries.Length {
        if (A_Index = excludeIndex)
            continue
        item := entries[A_Index]
        ch := item.HasProp("char") ? item.char : ""
        if (ch != "")
            taken[ch] := true
    }
    avail := []
    for c in seqFn() {
        if (!taken.Has(c) && isValidFn(c))
            avail.Push(c)
    }
    return avail
}

UtilitySelector_PromptChar(seqFn, isValidFn, entries, currentChar := "", excludeIndex := 0) {
    avail := UtilitySelector_AvailableChars(seqFn, isValidFn, entries, "char", excludeIndex)
    if (avail.Length = 0 && (currentChar = "" || !isValidFn(currentChar))) {
        UtilitySelector_Notify("No free characters left. Delete an item first.")
        return ""
    }
    hint := ""
    for c in avail
        hint .= (hint = "" ? "" : " ") . c
    prompt := "Unique character from the assignment pool."
    if (hint != "")
        prompt .= "`nAvailable: " . hint
    result := UtilitySelector_InputBox(prompt, "Character", 520, currentChar)
    if (result.Result != "OK")
        return ""
    ch := StrLower(Trim(result.Value))
    if (StrLen(ch) != 1 || !isValidFn(ch)) {
        UtilitySelector_Notify("Character must be one of the assignment pool keys.")
        return ""
    }
    for c in avail {
        if (c = ch)
            return ch
    }
    if (ch = currentChar)
        return ch
    UtilitySelector_Notify("Character '" . ch . "' is already assigned.")
    return ""
}

UtilitySelector_OnAdd(*) {
    global g_UtilitySelectorCategory, g_PromptEntries, g_HotstringEntries
    if (g_UtilitySelectorCategory = "Prompts")
        UtilitySelector_PromptsAdd()
    else if (g_UtilitySelectorCategory = "Hotstrings")
        UtilitySelector_HotstringsAdd()
}

UtilitySelector_OnEdit(*) {
    global g_UtilitySelectorCategory
    if (g_UtilitySelectorCategory = "Prompts")
        UtilitySelector_PromptsEdit()
    else if (g_UtilitySelectorCategory = "Hotstrings")
        UtilitySelector_HotstringsEdit()
}

UtilitySelector_OnDelete(*) {
    global g_UtilitySelectorCategory
    if (g_UtilitySelectorCategory = "Prompts")
        UtilitySelector_PromptsDelete()
    else if (g_UtilitySelectorCategory = "Hotstrings")
        UtilitySelector_HotstringsDelete()
}

UtilitySelector_PromptsAdd() {
    global g_PromptEntries
    PromptData_Load()
    UtilitySelector_DialogsBegin()
    nameBox := UtilitySelector_InputBox("Prompt name:", "Add prompt")
    if (nameBox.Result != "OK") {
        UtilitySelector_DialogsEnd()
        UtilitySelector_RefocusGui()
        return
    }
    name := Trim(nameBox.Value)
    if (name = "") {
        UtilitySelector_DialogsEnd()
        UtilitySelector_Notify("Name is required.")
        UtilitySelector_RefocusGui()
        return
    }
    catBox := UtilitySelector_InputBox("Category (e.g. General or Mnemonic):", "Prompt category", 420, "General")
    if (catBox.Result != "OK") {
        UtilitySelector_DialogsEnd()
        UtilitySelector_RefocusGui()
        return
    }
    category := Trim(catBox.Value)
    if (category = "")
        category := "General"
    ch := UtilitySelector_PromptChar(PromptData_CharSequence, PromptData_IsValidChar, g_PromptEntries)
    if (ch = "") {
        UtilitySelector_DialogsEnd()
        UtilitySelector_RefocusGui()
        return
    }
    startDir := A_ScriptDir "\assets\prompt\"
    selected := FileSelect(1, startDir, "Select prompt file", "Text (*.txt)")
    UtilitySelector_DialogsEnd()
    if (selected = "") {
        UtilitySelector_RefocusGui()
        return
    }
    list := []
    for p in g_PromptEntries
        list.Push(p)
    list.Push({ name: name, char: ch, category: category, author: "", filePath: PromptData_ToStoredPath(selected),
        source: "file" })
    if (!PromptData_Save(list)) {
        UtilitySelector_Notify("Failed to save prompt.")
        UtilitySelector_RefocusGui()
        return
    }
    UtilitySelector_RebuildPromptCharMap()
    UtilitySelector_PopulateLv()
    UtilitySelector_BindModalHotkeys()
    UtilitySelector_RefocusGui()
}

UtilitySelector_PromptsEdit() {
    global g_PromptEntries, g_UtilitySelectorRows
    idx := UtilitySelector_SelectedIndex()
    if (idx < 1 || idx > g_UtilitySelectorRows.Length) {
        UtilitySelector_Notify("Select a prompt to edit.")
        return
    }
    row := g_UtilitySelectorRows[idx]
    listIndex := row.HasProp("listIndex") ? row.listIndex : 0
    PromptData_Load()
    if (listIndex < 1 || listIndex > g_PromptEntries.Length)
        return
    prompt := g_PromptEntries[listIndex]
    UtilitySelector_DialogsBegin()
    nameBox := UtilitySelector_InputBox("Prompt name:", "Edit prompt", 420, prompt.name)
    if (nameBox.Result != "OK") {
        UtilitySelector_DialogsEnd()
        UtilitySelector_RefocusGui()
        return
    }
    name := Trim(nameBox.Value)
    if (name = "") {
        UtilitySelector_DialogsEnd()
        UtilitySelector_Notify("Name is required.")
        UtilitySelector_RefocusGui()
        return
    }
    catBox := UtilitySelector_InputBox("Category (e.g. General or Mnemonic):", "Prompt category", 420, prompt.category)
    if (catBox.Result != "OK") {
        UtilitySelector_DialogsEnd()
        UtilitySelector_RefocusGui()
        return
    }
    category := Trim(catBox.Value)
    if (category = "")
        category := "General"
    currentChar := prompt.HasProp("char") ? prompt.char : ""
    ch := UtilitySelector_PromptChar(PromptData_CharSequence, PromptData_IsValidChar, g_PromptEntries, currentChar,
        listIndex)
    if (ch = "") {
        UtilitySelector_DialogsEnd()
        UtilitySelector_RefocusGui()
        return
    }
    start := PromptData_ResolvePath(prompt)
    selected := FileSelect(1, start, "Select prompt file", "Text (*.txt)")
    UtilitySelector_DialogsEnd()
    filePath := prompt.filePath
    source := prompt.source
    if (selected != "") {
        filePath := PromptData_ToStoredPath(selected)
        source := "file"
    }
    list := []
    loop g_PromptEntries.Length {
        if (A_Index = listIndex)
            list.Push({ name: name, char: ch, category: category,
                author: prompt.HasProp("author") ? prompt.author : "", filePath: filePath, source: source })
        else
            list.Push(g_PromptEntries[A_Index])
    }
    if (!PromptData_Save(list)) {
        UtilitySelector_Notify("Failed to save prompt.")
        UtilitySelector_RefocusGui()
        return
    }
    UtilitySelector_RebuildPromptCharMap()
    UtilitySelector_PopulateLv()
    UtilitySelector_BindModalHotkeys()
    try g_HotstringSelectorLv.Modify(idx, "Select Focus Vis")
    catch {
    }
    UtilitySelector_RefocusGui()
}

UtilitySelector_PromptsDelete() {
    global g_PromptEntries, g_UtilitySelectorRows, g_HotstringSelectorLv
    idx := UtilitySelector_SelectedIndex()
    if (idx < 1 || idx > g_UtilitySelectorRows.Length) {
        UtilitySelector_Notify("Select a prompt to delete.")
        return
    }
    row := g_UtilitySelectorRows[idx]
    listIndex := row.HasProp("listIndex") ? row.listIndex : 0
    PromptData_Load()
    if (listIndex < 1 || listIndex > g_PromptEntries.Length)
        return
    prompt := g_PromptEntries[listIndex]
    label := prompt.name != "" ? prompt.name : "(unnamed)"
    hwnd := UtilitySelector_SelectorHwnd()
    msgOpts := "YesNo Icon! Default2"
    if (hwnd)
        msgOpts .= " Owner" . hwnd
    UtilitySelector_DialogsBegin()
    confirmed := (MsgBox("Delete prompt '" . label . "'?`nThe prompt file is not deleted.", "Delete prompt", msgOpts) =
    "Yes")
    UtilitySelector_DialogsEnd()
    if (!confirmed) {
        UtilitySelector_RefocusGui()
        return
    }
    list := []
    loop g_PromptEntries.Length {
        if (A_Index != listIndex)
            list.Push(g_PromptEntries[A_Index])
    }
    if (!PromptData_Save(list)) {
        UtilitySelector_Notify("Failed to save prompt list.")
        UtilitySelector_RefocusGui()
        return
    }
    UtilitySelector_RebuildPromptCharMap()
    UtilitySelector_PopulateLv()
    UtilitySelector_BindModalHotkeys()
    UtilitySelector_RefocusGui()
}

UtilitySelector_HotstringsAdd() {
    global g_HotstringEntries
    HotstringData_Load()
    UtilitySelector_DialogsBegin()
    nameBox := UtilitySelector_InputBox("Name:", "Add hotstring")
    if (nameBox.Result != "OK") {
        UtilitySelector_DialogsEnd()
        UtilitySelector_RefocusGui()
        return
    }
    name := Trim(nameBox.Value)
    if (name = "") {
        UtilitySelector_DialogsEnd()
        UtilitySelector_Notify("Name is required.")
        UtilitySelector_RefocusGui()
        return
    }
    textBox := UtilitySelector_InputBox("Text to paste:", "Hotstring text", 520)
    if (textBox.Result != "OK") {
        UtilitySelector_DialogsEnd()
        UtilitySelector_RefocusGui()
        return
    }
    textVal := textBox.Value
    ch := UtilitySelector_PromptChar(HotstringData_CharSequence, HotstringData_IsValidChar, g_HotstringEntries)
    UtilitySelector_DialogsEnd()
    if (ch = "") {
        UtilitySelector_RefocusGui()
        return
    }
    list := []
    for item in g_HotstringEntries
        list.Push(item)
    list.Push({ name: name, char: ch, text: textVal })
    if (!HotstringData_Save(list)) {
        UtilitySelector_Notify("Failed to save hotstring.")
        UtilitySelector_RefocusGui()
        return
    }
    UtilitySelector_PopulateLv()
    UtilitySelector_BindModalHotkeys()
    UtilitySelector_RefocusGui()
}

UtilitySelector_HotstringsEdit() {
    global g_HotstringEntries, g_UtilitySelectorRows, g_HotstringSelectorLv
    idx := UtilitySelector_SelectedIndex()
    if (idx < 1 || idx > g_HotstringEntries.Length) {
        UtilitySelector_Notify("Select a hotstring to edit.")
        return
    }
    HotstringData_Load()
    if (idx > g_HotstringEntries.Length)
        return
    item := g_HotstringEntries[idx]
    UtilitySelector_DialogsBegin()
    nameBox := UtilitySelector_InputBox("Name:", "Edit hotstring", 420, item.name)
    if (nameBox.Result != "OK") {
        UtilitySelector_DialogsEnd()
        UtilitySelector_RefocusGui()
        return
    }
    name := Trim(nameBox.Value)
    if (name = "") {
        UtilitySelector_DialogsEnd()
        UtilitySelector_Notify("Name is required.")
        UtilitySelector_RefocusGui()
        return
    }
    textBox := UtilitySelector_InputBox("Text to paste:", "Hotstring text", 520, item.text)
    if (textBox.Result != "OK") {
        UtilitySelector_DialogsEnd()
        UtilitySelector_RefocusGui()
        return
    }
    textVal := textBox.Value
    currentChar := item.HasProp("char") ? item.char : ""
    ch := UtilitySelector_PromptChar(HotstringData_CharSequence, HotstringData_IsValidChar, g_HotstringEntries,
        currentChar, idx)
    UtilitySelector_DialogsEnd()
    if (ch = "") {
        UtilitySelector_RefocusGui()
        return
    }
    list := []
    loop g_HotstringEntries.Length {
        if (A_Index = idx)
            list.Push({ name: name, char: ch, text: textVal })
        else
            list.Push(g_HotstringEntries[A_Index])
    }
    if (!HotstringData_Save(list)) {
        UtilitySelector_Notify("Failed to save hotstring.")
        UtilitySelector_RefocusGui()
        return
    }
    UtilitySelector_PopulateLv()
    UtilitySelector_BindModalHotkeys()
    try g_HotstringSelectorLv.Modify(idx, "Select Focus Vis")
    catch {
    }
    UtilitySelector_RefocusGui()
}

UtilitySelector_HotstringsDelete() {
    global g_HotstringEntries
    idx := UtilitySelector_SelectedIndex()
    if (idx < 1) {
        UtilitySelector_Notify("Select a hotstring to delete.")
        return
    }
    HotstringData_Load()
    if (idx > g_HotstringEntries.Length)
        return
    item := g_HotstringEntries[idx]
    label := item.name != "" ? item.name : "(unnamed)"
    hwnd := UtilitySelector_SelectorHwnd()
    msgOpts := "YesNo Icon! Default2"
    if (hwnd)
        msgOpts .= " Owner" . hwnd
    UtilitySelector_DialogsBegin()
    confirmed := (MsgBox("Delete hotstring '" . label . "'?", "Delete hotstring", msgOpts) = "Yes")
    UtilitySelector_DialogsEnd()
    if (!confirmed) {
        UtilitySelector_RefocusGui()
        return
    }
    list := []
    loop g_HotstringEntries.Length {
        if (A_Index != idx)
            list.Push(g_HotstringEntries[A_Index])
    }
    if (!HotstringData_Save(list)) {
        UtilitySelector_Notify("Failed to save hotstring list.")
        UtilitySelector_RefocusGui()
        return
    }
    UtilitySelector_PopulateLv()
    UtilitySelector_BindModalHotkeys()
    UtilitySelector_RefocusGui()
}
