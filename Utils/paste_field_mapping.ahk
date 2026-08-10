; =============================================================================
; Utils module: paste_field_mapping.ahk
; Learn-and-persist window → main text field mappings for Win+Alt+Shift+L.
; Before paste: focus saved field via UIA when exe+title match an INI entry.
; After paste (no mapping): Interactive Input Y/N to save the focused field.
; Persistent store: assets/data/paste_field_mappings.ini
; =============================================================================

global g_PasteFieldMappingsCache := []
global g_PasteFieldMappingsCacheReady := false
global g_PasteFieldLearnPromptBusy := false

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
        name := ""
        automationId := ""
        className := ""
        typeVal := ""
        uiaAttach := "element"
        try titleNeedle := IniRead(path, section, "TitleNeedle", "")
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

PasteField_WindowLabel(hwnd) {
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

; First INI mapping matching this window's exe + title needle.
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
    for mapping in PasteField_LoadMappings() {
        if (StrLower(mapping.exe) != exeLower)
            continue
        needle := Trim(mapping.titleNeedle)
        if (needle != "" && !InStr(title, needle, false))
            continue
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

PasteField_FocusFromMapping(hwnd, mapping) {
    if (!hwnd || !IsObject(mapping))
        return false
    aid := Trim(mapping.automationId)
    name := Trim(mapping.name)
    cn := Trim(mapping.className)
    typeVal := Trim(String(mapping.type))
    typeArg := ""
    if (typeVal != "")
        typeArg := IsInteger(typeVal) ? Integer(typeVal) : typeVal
    if (aid = "" && name = "" && cn = "")
        return false

    root := PasteField_AttachRoot(hwnd, mapping.uiaAttach)
    if (!root)
        return false

    el := 0
    ; Prefer AutomationId, then Name+Type, then ClassName+Type.
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
    path := PasteField_MappingsIniPath()
    try DirCreate(A_ScriptDir "\assets\data")
    idx := PasteField_NextMappingIndex()
    section := "Mapping_" . idx
    try {
        IniWrite(exe, path, section, "Exe")
        IniWrite(titleNeedle, path, section, "TitleNeedle")
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
