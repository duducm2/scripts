; =============================================================================
; Utils module: paste_field_mapping.ahk
; Learn-and-persist window → main text field mappings for Win+Alt+Shift+L.
; Before paste: focus saved field via UIA when a mapping matches.
; Match gates: (1) exe + UrlNeedle (browsers) or title needle, (2) same exe +
;   strict UIA field identity when title/url changed (non-browser; browsers only
;   when UrlNeedle matches current page).
; After paste (no mapping): Interactive Input Y/N to save the focused field.
; Persistent store: assets/data/paste_field_mappings.ini
; Manage UI: #!+L [M] ListView to delete saved mappings.
; =============================================================================

global g_PasteFieldMappingsCache := []
global g_PasteFieldMappingsCacheReady := false
global g_PasteFieldLearnPromptBusy := false
global g_PasteFieldManageResult := ""
global g_PasteFieldManageGui := false
global g_PasteFieldManageLv := false
global g_PasteFieldManageList := []
global g_PasteFieldManageHotkeysBound := false

PasteField_MappingsIniPath() {
    return A_ScriptDir "\assets\data\paste_field_mappings.ini"
}

PasteField_InvalidateCache() {
    global g_PasteFieldMappingsCache, g_PasteFieldMappingsCacheReady
    g_PasteFieldMappingsCache := []
    g_PasteFieldMappingsCacheReady := false
}

; Read all [Mapping_N] sections into an array of objects.
PasteField_LoadMappings() {
    global g_PasteFieldMappingsCache, g_PasteFieldMappingsCacheReady
    if (g_PasteFieldMappingsCacheReady)
        return g_PasteFieldMappingsCache
    list := []
    path := PasteField_MappingsIniPath()
    if (!FileExist(path)) {
        g_PasteFieldMappingsCache := list
        g_PasteFieldMappingsCacheReady := true
        return list
    }
    idx := 1
    loop 200 {
        section := "Mapping_" . idx
        exe := ""
        try exe := IniRead(path, section, "Exe", "")
        catch {
            break
        }
        if (exe = "" || exe = "ERROR")
            break
        titleNeedle := ""
        urlNeedle := ""
        name := ""
        automationId := ""
        className := ""
        typeVal := ""
        uiaAttach := "element"
        try titleNeedle := IniRead(path, section, "TitleNeedle", "")
        try urlNeedle := IniRead(path, section, "UrlNeedle", "")
        try name := IniRead(path, section, "Name", "")
        try automationId := IniRead(path, section, "AutomationId", "")
        try className := IniRead(path, section, "ClassName", "")
        try typeVal := IniRead(path, section, "Type", "")
        try uiaAttach := IniRead(path, section, "UiaAttach", "element")
        if (uiaAttach = "" || uiaAttach = "ERROR")
            uiaAttach := "element"
        list.Push({
            index: idx,
            exe: exe,
            titleNeedle: titleNeedle = "ERROR" ? "" : titleNeedle,
            urlNeedle: urlNeedle = "ERROR" ? "" : urlNeedle,
            name: name = "ERROR" ? "" : name,
            automationId: automationId = "ERROR" ? "" : automationId,
            className: className = "ERROR" ? "" : className,
            type: typeVal = "ERROR" ? "" : typeVal,
            uiaAttach: uiaAttach
        })
        idx += 1
    }
    g_PasteFieldMappingsCache := list
    g_PasteFieldMappingsCacheReady := true
    return list
}

PasteField_NextMappingIndex() {
    return PasteField_LoadMappings().Length + 1
}

; host + path, lowercase; strip scheme, query, hash; trim trailing /.
PasteField_NormalizeUrlNeedle(url) {
    u := Trim(String(url))
    if (u = "")
        return ""
    u := RegExReplace(u, "i)^https?://", "")
    hashPos := InStr(u, "#")
    if (hashPos > 0)
        u := SubStr(u, 1, hashPos - 1)
    qPos := InStr(u, "?")
    if (qPos > 0)
        u := SubStr(u, 1, qPos - 1)
    u := Trim(u)
    while (StrLen(u) > 1 && SubStr(u, -1) = "/")
        u := SubStr(u, 1, StrLen(u) - 1)
    return StrLower(u)
}

; Current page URL needle for Chromium hwnds; "" if unavailable or non-browser.
PasteField_CaptureUrlNeedle(hwnd) {
    if (!hwnd)
        return ""
    cls := ""
    try cls := WinGetClass("ahk_id " hwnd)
    if (!PasteField_IsChromiumClass(cls))
        return ""
    try {
        uia := UIA_Browser("ahk_id " hwnd)
        url := ""
        try url := uia.GetCurrentURL(false)
        catch {
            url := ""
        }
        needle := PasteField_NormalizeUrlNeedle(url)
        if (needle != "")
            return needle
        try url := uia.GetCurrentURL(true)
        catch {
            url := ""
        }
        return PasteField_NormalizeUrlNeedle(url)
    } catch {
        return ""
    }
}

; True when currentUrlNeedle equals or is under mapping UrlNeedle (prefix path match).
PasteField_UrlNeedleMatches(mappingUrlNeedle, currentUrlNeedle) {
    m := Trim(StrLower(mappingUrlNeedle))
    c := Trim(StrLower(currentUrlNeedle))
    if (m = "" || c = "")
        return false
    if (c = m)
        return true
    if (InStr(c, m) = 1) {
        next := SubStr(c, StrLen(m) + 1, 1)
        return (next = "" || next = "/")
    }
    return false
}

PasteField_WindowLabel(hwnd) {
    urlNeedle := PasteField_CaptureUrlNeedle(hwnd)
    if (urlNeedle != "")
        return urlNeedle
    title := ""
    exe := ""
    try title := WinGetTitle("ahk_id " hwnd)
    try exe := WinGetProcessName("ahk_id " hwnd)
    needle := PasteField_DeriveTitleNeedle(title, exe)
    if (needle != "")
        return needle
    if (exe != "")
        return exe
    return "this window"
}

; Stable title needle for matching: strip browser suffixes; prefer segment before " - "/" | ".
PasteField_DeriveTitleNeedle(title, exe := "") {
    t := Trim(title)
    if (t = "")
        return ""
    suffixes := [
        " - Google Chrome",
        " - Chromium",
        " - Microsoft Edge",
        " - Personal",
        " - Work"
    ]
    for suf in suffixes {
        if (StrLen(t) > StrLen(suf) && SubStr(t, -StrLen(suf)) = suf)
            t := Trim(SubStr(t, 1, StrLen(t) - StrLen(suf)))
    }
    sepPos := InStr(t, " - ")
    pipePos := InStr(t, " | ")
    cut := 0
    if (sepPos > 0)
        cut := sepPos
    if (pipePos > 0 && (cut = 0 || pipePos < cut))
        cut := pipePos
    if (cut > 1) {
        head := Trim(SubStr(t, 1, cut - 1))
        if (head != "" && StrLen(head) <= 80)
            t := head
    }
    t := Trim(t)
    if (t = "")
        return ""
    return t
}

PasteField_IsChromiumClass(className) {
    if (!className)
        return false
    if (className = "Chrome_WidgetWin_1" || className = "TeamsWebView")
        return true
    return InStr(className, "Chrome_WidgetWin")
}

PasteField_DetectUiaAttach(hwnd) {
    cls := ""
    try cls := WinGetClass("ahk_id " hwnd)
    return PasteField_IsChromiumClass(cls) ? "browser" : "element"
}

PasteField_IsBrowserHwnd(hwnd) {
    if (!hwnd)
        return false
    cls := ""
    try cls := WinGetClass("ahk_id " hwnd)
    return PasteField_IsChromiumClass(cls)
}

; Mapping for hwnd: url/title gate first, then strict UIA field identity (same exe).
PasteField_FindMapping(hwnd) {
    if (!hwnd)
        return ""
    exe := ""
    title := ""
    try exe := WinGetProcessName("ahk_id " hwnd)
    try title := WinGetTitle("ahk_id " hwnd)
    if (!exe)
        return ""
    exeLower := StrLower(exe)
    sameExe := []
    for mapping in PasteField_LoadMappings() {
        if (StrLower(mapping.exe) = exeLower)
            sameExe.Push(mapping)
    }
    if (sameExe.Length = 0)
        return ""

    isBrowser := PasteField_IsBrowserHwnd(hwnd)
    currentUrlNeedle := isBrowser ? PasteField_CaptureUrlNeedle(hwnd) : ""

    ; Pass 1 — url (browsers) or title gate (no UIA).
    if (isBrowser && currentUrlNeedle != "") {
        best := ""
        bestLen := -1
        for mapping in sameExe {
            urlNeedle := Trim(mapping.HasProp("urlNeedle") ? mapping.urlNeedle : "")
            if (urlNeedle = "")
                continue
            if (!PasteField_UrlNeedleMatches(urlNeedle, currentUrlNeedle))
                continue
            if (StrLen(urlNeedle) > bestLen) {
                best := mapping
                bestLen := StrLen(urlNeedle)
            }
        }
        if (best)
            return best
        ; No UrlNeedle hit — fall through to legacy TitleNeedle for rows without UrlNeedle.
        for mapping in sameExe {
            urlNeedle := Trim(mapping.HasProp("urlNeedle") ? mapping.urlNeedle : "")
            if (urlNeedle != "")
                continue
            needle := Trim(mapping.titleNeedle)
            if (needle = "" || InStr(title, needle, false))
                return mapping
        }
    } else {
        ; Non-browser, or browser with URL capture failure: TitleNeedle gate.
        for mapping in sameExe {
            needle := Trim(mapping.titleNeedle)
            if (needle = "" || InStr(title, needle, false))
                return mapping
        }
    }

    ; Pass 2 — strict UIA gate (AutomationId or ClassName+Type). Prefer AutomationId rows.
    ; Browser: only mappings whose UrlNeedle matches current page (legacy empty UrlNeedle excluded).
    aidCandidates := []
    cnCandidates := []
    for mapping in sameExe {
        if (isBrowser) {
            urlNeedle := Trim(mapping.HasProp("urlNeedle") ? mapping.urlNeedle : "")
            if (urlNeedle = "" || currentUrlNeedle = ""
                || !PasteField_UrlNeedleMatches(urlNeedle, currentUrlNeedle))
                continue
        }
        aid := Trim(mapping.automationId)
        cn := Trim(mapping.className)
        typeVal := Trim(String(mapping.HasProp("type") ? mapping.type : ""))
        if (aid != "")
            aidCandidates.Push(mapping)
        else if (cn != "" && typeVal != "")
            cnCandidates.Push(mapping)
    }
    candidates := []
    for mapping in aidCandidates
        candidates.Push(mapping)
    for mapping in cnCandidates
        candidates.Push(mapping)
    if (candidates.Length = 0)
        return ""

    rootsByAttach := Map()
    for mapping in candidates {
        attach := mapping.uiaAttach = "browser" ? "browser" : "element"
        root := 0
        if (rootsByAttach.Has(attach))
            root := rootsByAttach[attach]
        else {
            root := PasteField_AttachRoot(hwnd, attach)
            rootsByAttach[attach] := root
        }
        if (!root)
            continue
        if (PasteField_FindElementFromMapping(root, mapping, true))
            return mapping
    }
    return ""
}

PasteField_AttachRoot(hwnd, uiaAttach) {
    if (!hwnd)
        return 0
    attach := uiaAttach = "browser" ? "browser" : "element"
    try {
        if (attach = "browser") {
            uia := UIA_Browser("ahk_id " hwnd)
            Sleep 120
            return uia
        }
        return UIA.ElementFromHandle(hwnd)
    } catch {
        return 0
    }
}

; Find UIA element for a mapping under root.
; strictUiq=false: AutomationId → Name(+Type) → ClassName(+Type) (focus path).
; strictUiq=true: AutomationId(+Type) only, else ClassName AND Type; never Name.
PasteField_FindElementFromMapping(root, mapping, strictUiq := false) {
    if (!root || !IsObject(mapping))
        return 0
    aid := Trim(mapping.HasProp("automationId") ? mapping.automationId : "")
    name := Trim(mapping.HasProp("name") ? mapping.name : "")
    cn := Trim(mapping.HasProp("className") ? mapping.className : "")
    typeVal := Trim(String(mapping.HasProp("type") ? mapping.type : ""))
    typeArg := ""
    if (typeVal != "")
        typeArg := IsInteger(typeVal) ? Integer(typeVal) : typeVal

    if (strictUiq) {
        el := 0
        if (aid != "") {
            try {
                if (typeArg != "")
                    el := root.FindFirst({ AutomationId: aid, Type: typeArg })
                else
                    el := root.FindFirst({ AutomationId: aid })
            } catch {
                el := 0
            }
            return el ? el : 0
        }
        if (cn != "" && typeArg != "") {
            try el := root.FindFirst({ ClassName: cn, Type: typeArg })
            catch {
                el := 0
            }
        }
        return el ? el : 0
    }

    if (aid = "" && name = "" && cn = "")
        return 0

    el := 0
    try {
        if (aid != "") {
            if (typeArg != "")
                el := root.FindFirst({ AutomationId: aid, Type: typeArg })
            else
                el := root.FindFirst({ AutomationId: aid })
        }
    } catch {
        el := 0
    }
    if (!el && name != "") {
        try {
            if (typeArg != "")
                el := root.FindFirst({ Name: name, Type: typeArg })
            else
                el := root.FindFirst({ Name: name })
        } catch {
            el := 0
        }
    }
    if (!el && cn != "") {
        try {
            if (typeArg != "")
                el := root.FindFirst({ ClassName: cn, Type: typeArg })
            else
                el := root.FindFirst({ ClassName: cn })
        } catch {
            el := 0
        }
    }
    return el ? el : 0
}

PasteField_FocusFromMapping(hwnd, mapping) {
    if (!hwnd || !IsObject(mapping))
        return false
    aid := Trim(mapping.HasProp("automationId") ? mapping.automationId : "")
    name := Trim(mapping.HasProp("name") ? mapping.name : "")
    cn := Trim(mapping.HasProp("className") ? mapping.className : "")
    if (aid = "" && name = "" && cn = "")
        return false
    root := PasteField_AttachRoot(hwnd, mapping.uiaAttach)
    if (!root)
        return false
    el := PasteField_FindElementFromMapping(root, mapping, false)
    if (!el)
        return false
    return PasteField_SetFocusWithFallback(el)
}

; Lookup + focus. Returns true if a mapping entry exists (even if focus fails).
PasteField_TryFocusMappedField(hwnd) {
    mapping := PasteField_FindMapping(hwnd)
    if (!mapping)
        return false
    PasteField_FocusFromMapping(hwnd, mapping)
    return true
}

; Lookup + focus. Returns { hasMapping: bool, focused: bool }.
PasteField_FocusMappedField(hwnd) {
    mapping := PasteField_FindMapping(hwnd)
    if (!mapping)
        return { hasMapping: false, focused: false }
    focused := PasteField_FocusFromMapping(hwnd, mapping)
    return { hasMapping: true, focused: !!focused }
}

PasteField_CaptureFocusedFieldSignature() {
    if (!IsSet(UIA))
        return ""
    try {
        el := UIA.GetFocusedElement()
        if (!el)
            return ""
        name := ""
        automationId := ""
        className := ""
        typeVal := ""
        try name := el.Name
        try automationId := el.AutomationId
        try className := el.ClassName
        try typeVal := el.Type
        if (typeVal = "" || typeVal = 0) {
            try typeVal := el.ControlType
        }
        if (name = "" && automationId = "" && className = "")
            return ""
        return {
            name: name,
            automationId: automationId,
            className: className,
            type: typeVal
        }
    } catch {
        return ""
    }
}

PasteField_SaveMapping(hwnd, signature) {
    if (!hwnd || !IsObject(signature))
        return false
    exe := ""
    title := ""
    try exe := WinGetProcessName("ahk_id " hwnd)
    try title := WinGetTitle("ahk_id " hwnd)
    if (!exe)
        return false
    titleNeedle := PasteField_DeriveTitleNeedle(title, exe)
    uiaAttach := PasteField_DetectUiaAttach(hwnd)
    urlNeedle := ""
    if (uiaAttach = "browser")
        urlNeedle := PasteField_CaptureUrlNeedle(hwnd)
    path := PasteField_MappingsIniPath()
    try DirCreate(A_ScriptDir "\assets\data")
    idx := PasteField_NextMappingIndex()
    section := "Mapping_" . idx
    try {
        IniWrite(exe, path, section, "Exe")
        IniWrite(titleNeedle, path, section, "TitleNeedle")
        IniWrite(urlNeedle, path, section, "UrlNeedle")
        IniWrite(signature.HasProp("name") ? signature.name : "", path, section, "Name")
        IniWrite(signature.HasProp("automationId") ? signature.automationId : "", path, section, "AutomationId")
        IniWrite(signature.HasProp("className") ? signature.className : "", path, section, "ClassName")
        typeStr := ""
        if (signature.HasProp("type"))
            typeStr := String(signature.type)
        IniWrite(typeStr, path, section, "Type")
        IniWrite(uiaAttach, path, section, "UiaAttach")
    } catch {
        return false
    }
    PasteField_InvalidateCache()
    return true
}

; Rewrite INI as contiguous [Mapping_1]…[Mapping_N] (required after mid-list deletes).
PasteField_WriteAllMappings(list) {
    path := PasteField_MappingsIniPath()
    try DirCreate(A_ScriptDir "\assets\data")
    catch {
    }
    try FileDelete(path)
    catch {
    }
    if (!IsObject(list) || list.Length = 0) {
        PasteField_InvalidateCache()
        return true
    }
    try {
        idx := 1
        for mapping in list {
            section := "Mapping_" . idx
            IniWrite(mapping.HasProp("exe") ? mapping.exe : "", path, section, "Exe")
            IniWrite(mapping.HasProp("titleNeedle") ? mapping.titleNeedle : "", path, section, "TitleNeedle")
            urlNeedle := ""
            if (mapping.HasProp("urlNeedle"))
                urlNeedle := mapping.urlNeedle
            IniWrite(urlNeedle, path, section, "UrlNeedle")
            IniWrite(mapping.HasProp("name") ? mapping.name : "", path, section, "Name")
            IniWrite(mapping.HasProp("automationId") ? mapping.automationId : "", path, section, "AutomationId")
            IniWrite(mapping.HasProp("className") ? mapping.className : "", path, section, "ClassName")
            typeStr := ""
            if (mapping.HasProp("type"))
                typeStr := String(mapping.type)
            IniWrite(typeStr, path, section, "Type")
            uiaAttach := "element"
            if (mapping.HasProp("uiaAttach") && mapping.uiaAttach != "")
                uiaAttach := mapping.uiaAttach
            IniWrite(uiaAttach, path, section, "UiaAttach")
            idx += 1
        }
    } catch {
        PasteField_InvalidateCache()
        return false
    }
    PasteField_InvalidateCache()
    return true
}

; Remove 1-based list index; rewrite remaining as Mapping_1…N.
PasteField_RemoveMappingAt(index) {
    list := PasteField_LoadMappings()
    index := Integer(index)
    if (index < 1 || index > list.Length)
        return false
    newList := []
    for i, mapping in list {
        if (i = index)
            continue
        newList.Push(mapping)
    }
    return PasteField_WriteAllMappings(newList)
}

; Remove several 1-based indices in one rewrite (any order).
PasteField_RemoveMappingsAt(indices) {
    if (!IsObject(indices) || indices.Length = 0)
        return false
    remove := Map()
    for idx in indices {
        n := Integer(idx)
        if (n >= 1)
            remove[n] := true
    }
    if (remove.Count = 0)
        return false
    list := PasteField_LoadMappings()
    newList := []
    for i, mapping in list {
        if (remove.Has(i))
            continue
        newList.Push(mapping)
    }
    if (newList.Length = list.Length)
        return false
    return PasteField_WriteAllMappings(newList)
}

PasteField_ManageFieldLabel(mapping) {
    if (!IsObject(mapping))
        return ""
    name := Trim(mapping.HasProp("name") ? mapping.name : "")
    if (name != "")
        return name
    aid := Trim(mapping.HasProp("automationId") ? mapping.automationId : "")
    if (aid != "")
        return aid
    return Trim(mapping.HasProp("className") ? mapping.className : "")
}

PasteField_ManageTruncate(text, maxLen := 72) {
    t := String(text)
    if (StrLen(t) <= maxLen)
        return t
    return SubStr(t, 1, maxLen - 1) . "…"
}

PasteField_ManagePopulateLv() {
    global g_PasteFieldManageLv, g_PasteFieldManageList
    if (!IsObject(g_PasteFieldManageLv))
        return
    g_PasteFieldManageLv.Delete()
    for mapping in g_PasteFieldManageList {
        exe := mapping.HasProp("exe") ? mapping.exe : ""
        urlNeedle := Trim(mapping.HasProp("urlNeedle") ? mapping.urlNeedle : "")
        title := urlNeedle != "" ? urlNeedle : (mapping.HasProp("titleNeedle") ? mapping.titleNeedle : "")
        field := PasteField_ManageFieldLabel(mapping)
        g_PasteFieldManageLv.Add("", exe, PasteField_ManageTruncate(title, 80),
        PasteField_ManageTruncate(field, 80))
    }
    try g_PasteFieldManageLv.ModifyCol(1, 140)
    try g_PasteFieldManageLv.ModifyCol(2, 280)
    try g_PasteFieldManageLv.ModifyCol(3, 260)
}

PasteField_ManageUnbindHotkeys() {
    global g_PasteFieldManageHotkeysBound, g_PasteFieldManageGui
    if (!g_PasteFieldManageHotkeysBound)
        return
    hwnd := 0
    if (IsObject(g_PasteFieldManageGui)) {
        try hwnd := g_PasteFieldManageGui.Hwnd
        catch {
            hwnd := 0
        }
    }
    if (hwnd) {
        try HotIfWinActive("ahk_id " hwnd)
        catch {
        }
        try Hotkey("Delete", PasteField_ManageOnDelete, "Off")
        catch {
        }
        try Hotkey("Escape", PasteField_ManageOnDone, "Off")
        catch {
        }
        try HotIf()
        catch {
        }
    }
    g_PasteFieldManageHotkeysBound := false
}

PasteField_ManageCloseGui() {
    global g_PasteFieldManageGui, g_PasteFieldManageLv, g_PasteFieldManageList
    PasteField_ManageUnbindHotkeys()
    if (IsObject(g_PasteFieldManageGui)) {
        try g_PasteFieldManageGui.Destroy()
        catch {
        }
    }
    g_PasteFieldManageGui := false
    g_PasteFieldManageLv := false
    g_PasteFieldManageList := []
}

PasteField_ManageOnDone(*) {
    global g_PasteFieldManageResult, g_PasteFieldManageGui
    if (g_PasteFieldManageResult != "")
        return
    g_PasteFieldManageResult := "done"
    if (IsObject(g_PasteFieldManageGui)) {
        try g_PasteFieldManageGui.Hide()
        catch {
        }
    }
}

PasteField_ManageOnDelete(*) {
    global g_PasteFieldManageLv, g_PasteFieldManageList
    if (!IsObject(g_PasteFieldManageLv))
        return
    indices := []
    row := 0
    while (row := g_PasteFieldManageLv.GetNext(row))
        indices.Push(row)
    if (indices.Length = 0)
        return
    labels := []
    for idx in indices {
        if (idx >= 1 && idx <= g_PasteFieldManageList.Length) {
            m := g_PasteFieldManageList[idx]
            lab := (m.HasProp("exe") ? m.exe : "")
            urlNeedle := Trim(m.HasProp("urlNeedle") ? m.urlNeedle : "")
            if (urlNeedle != "")
                lab .= " — " . urlNeedle
            else if (m.HasProp("titleNeedle") && m.titleNeedle != "")
                lab .= " — " . m.titleNeedle
            labels.Push(PasteField_ManageTruncate(lab, 50))
        }
    }
    if (!PasteField_RemoveMappingsAt(indices)) {
        ShowCenteredOverlay_Utils("❌ Failed to remove mapping", 2000, BANNER_ACCENT_ERROR)
        return
    }
    g_PasteFieldManageList := PasteField_LoadMappings()
    if (labels.Length = 1)
        ShowCenteredOverlay_Utils("✅ Removed: " . labels[1], 1800, BANNER_ACCENT_SUCCESS)
    else
        ShowCenteredOverlay_Utils("✅ Removed " . labels.Length . " mappings", 1800, BANNER_ACCENT_SUCCESS)
    if (g_PasteFieldManageList.Length = 0) {
        PasteField_ManageOnDone()
        return
    }
    PasteField_ManagePopulateLv()
}

; Blocking ListView: select rows, Delete removes, Esc/Close dismisses.
PasteField_ShowMappingsManageUI() {
    global g_PasteFieldManageResult, g_PasteFieldManageGui, g_PasteFieldManageLv,
        g_PasteFieldManageList, g_PasteFieldManageHotkeysBound
    PasteField_InvalidateCache()
    list := PasteField_LoadMappings()
    if (list.Length = 0) {
        ShowCenteredOverlay_Utils("ℹ️ No main field mappings", 2200, BANNER_ACCENT_INFO)
        return
    }

    PasteField_ManageCloseGui()
    g_PasteFieldManageResult := ""
    g_PasteFieldManageList := list

    ; Avoid local name "gui" — conflicts with a global `gui` in the loaded script set (AHK v2).
    g_PasteFieldManageGui := Gui("+AlwaysOnTop +ToolWindow", "Main text field mappings")
    g_PasteFieldManageGui.SetFont("s10", "Segoe UI")
    g_PasteFieldManageGui.Add("Text", "w700", "Select a mapping and press Delete (or the button). Esc closes.")
    g_PasteFieldManageLv := g_PasteFieldManageGui.Add("ListView", "w700 h420 Multi", ["Exe", "Title / URL", "Field"])
    g_PasteFieldManageGui.Add("Button", "w100 Section", "Delete").OnEvent("Click", PasteField_ManageOnDelete)
    g_PasteFieldManageGui.Add("Button", "w100 ys", "Close").OnEvent("Click", PasteField_ManageOnDone)
    g_PasteFieldManageGui.OnEvent("Close", PasteField_ManageOnDone)
    g_PasteFieldManageGui.OnEvent("Escape", PasteField_ManageOnDone)

    PasteField_ManagePopulateLv()

    try {
        HotIfWinActive("ahk_id " g_PasteFieldManageGui.Hwnd)
        Hotkey("Delete", PasteField_ManageOnDelete, "On")
        Hotkey("Escape", PasteField_ManageOnDone, "On")
        HotIf()
        g_PasteFieldManageHotkeysBound := true
    } catch {
        g_PasteFieldManageHotkeysBound := false
    }

    g_PasteFieldManageGui.Show()
    try g_PasteFieldManageLv.Focus()
    catch {
    }

    start := A_TickCount
    while (g_PasteFieldManageResult = "") {
        if ((A_TickCount - start) >= 300000) {
            g_PasteFieldManageResult := "done"
            break
        }
        Sleep 50
    }
    g_PasteFieldManageResult := ""
    PasteField_ManageCloseGui()
}

PasteField_CloseLearnPrompt() {
    global g_PasteFieldLearnPromptBusy
    StandardLoadingBar_CloseKeysOverlay()
    StandardLoadingBar_Hide(0)
    g_PasteFieldLearnPromptBusy := false
}

PasteField_OnLearnYes(hwnd, signature, *) {
    PasteField_CloseLearnPrompt()
    label := PasteField_WindowLabel(hwnd)
    if (!IsObject(signature)) {
        ShowCenteredOverlay_Utils("❌ No focused text field to save", 2000, BANNER_ACCENT_ERROR)
        return
    }
    if (PasteField_SaveMapping(hwnd, signature))
        ShowCenteredOverlay_Utils("✅ Main text field saved for " . label, 1800, BANNER_ACCENT_SUCCESS)
    else
        ShowCenteredOverlay_Utils("❌ Failed to save main text field", 2000, BANNER_ACCENT_ERROR)
}

PasteField_OnLearnNo(*) {
    PasteField_CloseLearnPrompt()
}

; After paste: if this window has no mapping, ask to save the focused field.
; Captures signature before the banner so focus steal does not lose the field identity.
PasteField_PromptSaveMainField(hwnd) {
    global g_PasteFieldLearnPromptBusy
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return
    if (PasteField_FindMapping(hwnd))
        return
    if (g_PasteFieldLearnPromptBusy)
        return
    g_PasteFieldLearnPromptBusy := true

    sig := PasteField_CaptureFocusedFieldSignature()
    label := PasteField_WindowLabel(hwnd)
    fieldHint := ""
    if (IsObject(sig) && Trim(sig.name) != "") {
        n := Trim(sig.name)
        if (StrLen(n) > 60)
            n := SubStr(n, 1, 57) . "..."
        fieldHint := "`n" . n
    }

    msg := "❓ Set this text field as the main field for " . label . "?" . fieldHint
    keyCallbacks := Map(
        "Y", PasteField_OnLearnYes.Bind(hwnd, sig),
        "N", PasteField_OnLearnNo
    )
    StandardLoadingBar_ShowWithKeys(
        msg,
        keyCallbacks,
        5000,
        hwnd,
        PasteField_OnLearnNo,
        BANNER_ACCENT_INTERMEDIATE,
        520,
        17,
        "",
        true,
        "[Y] Yes  [N] No",
        true
    )
}

; Shared UIA focus helper: SetFocus -> poll -> Click fallback
PasteField_SetFocusWithFallback(el) {
    if (!el)
        return false
    try {
        el.SetFocus()
    } catch {
    }
    loop 3 {
        try {
            if (el.HasKeyboardFocus)
                return true
        } catch {
        }
        Sleep 40
    }
    try {
        el.Click()
    } catch {
    }
    Sleep 80
    try {
        return el.HasKeyboardFocus
    } catch {
        return false
    }
}
