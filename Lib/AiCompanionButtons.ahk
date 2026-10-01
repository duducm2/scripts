; =============================================================================
; Lib: AiCompanionButtons.ahk
; Per-companion saved UIA targets for New Chat, Menu, Search, and Send.
; Shift+L captures them; Shift+N / Shift+D / Shift+S / Shift+G try the saved element
; before the built-in finders. Included from Utils.ahk after AiCompanionModels.
; =============================================================================

; actionId → { name, automationId, className, controlType } per companion.
global g_AiCompanionButtonsCache := Map()

AiCompanionButtons_GetIniPath() {
    try {
        UiElements_EnsureImported()
        return UiElements_IniPath()
    } catch {
        return A_ScriptDir "\assets\data\ai_companion_buttons.ini"
    }
}

; Fixed catalog. Another action is one entry here, one ini key, and one call site.
AiCompanionButtons_Actions() {
    return [{ id: "NewChat", label: "New Chat", chord: "n" }, { id: "Menu", label: "Menu", chord: "d" }, { id: "Search",
        label: "Search", chord: "s" }, { id: "Send", label: "Send", chord: "g" }
    ]
}

AiCompanionButtons_ActionById(actionId) {
    actionId := Trim(actionId)
    for action in AiCompanionButtons_Actions() {
        if (action.id = actionId)
            return action
    }
    return 0
}

AiCompanionButtons_ActionByChord(chord) {
    chord := StrLower(Trim(chord))
    for action in AiCompanionButtons_Actions() {
        if (action.chord = chord)
            return action
    }
    return 0
}

AiCompanionButtons_EmptySpec() {
    return { name: "", automationId: "", className: "", controlType: 0 }
}

AiCompanionButtons_HasSpec(spec) {
    if (!IsObject(spec))
        return false
    return Trim(spec.name) != "" || Trim(spec.automationId) != "" || Trim(spec.className) != ""
}

AiCompanionButtons_Summary(spec) {
    if (!AiCompanionButtons_HasSpec(spec))
        return "(built-in)"
    if (Trim(spec.name) != "")
        return spec.name
    if (Trim(spec.automationId) != "")
        return spec.automationId
    return spec.className
}

AiCompanionButtons_Escape(value) {
    out := ""
    loop parse String(value) {
        if (A_LoopField = "\")
            out .= "\\"
        else if (A_LoopField = "|")
            out .= "\p"
        else
            out .= A_LoopField
    }
    return out
}

AiCompanionButtons_Unescape(value) {
    out := ""
    i := 1
    len := StrLen(value)
    while (i <= len) {
        ch := SubStr(value, i, 1)
        if (ch = "\" && i < len) {
            n := SubStr(value, i + 1, 1)
            if (n = "p")
                out .= "|"
            else if (n = "\")
                out .= "\"
            else
                out .= n
            i += 2
        } else {
            out .= ch
            i++
        }
    }
    return out
}

AiCompanionButtons_Serialize(spec) {
    if (!AiCompanionButtons_HasSpec(spec))
        return ""
    ct := 0
    try ct := Integer(spec.controlType)
    catch
        ct := 0
    return "name=" . AiCompanionButtons_Escape(spec.name)
    . "|automationId=" . AiCompanionButtons_Escape(spec.automationId)
    . "|className=" . AiCompanionButtons_Escape(spec.className)
    . "|controlType=" . ct
}

AiCompanionButtons_Parse(raw) {
    spec := AiCompanionButtons_EmptySpec()
    raw := Trim(raw)
    if (raw = "" || raw = "ERROR")
        return spec
    for part in StrSplit(raw, "|") {
        part := Trim(part)
        if (part = "")
            continue
        eq := InStr(part, "=")
        if (eq < 2)
            continue
        key := StrLower(Trim(SubStr(part, 1, eq - 1)))
        value := AiCompanionButtons_Unescape(SubStr(part, eq + 1))
        if (key = "name")
            spec.name := value
        else if (key = "automationid")
            spec.automationId := value
        else if (key = "classname")
            spec.className := value
        else if (key = "controltype") {
            n := 0
            try n := Integer(Trim(value))
            catch
                n := 0
            spec.controlType := n
        }
    }
    return spec
}

AiCompanionButtons_Invalidate(companion := "") {
    global g_AiCompanionButtonsCache
    if (companion = "") {
        g_AiCompanionButtonsCache := Map()
        return
    }
    if g_AiCompanionButtonsCache.Has(companion)
        g_AiCompanionButtonsCache.Delete(companion)
}

AiCompanionButtons_Load(companion) {
    global g_AiCompanionButtonsCache
    if !AiCompanionModels_IsValidCompanion(companion)
        return Map()
    if g_AiCompanionButtonsCache.Has(companion)
        return g_AiCompanionButtonsCache[companion]

    store := Map()
    iniPath := AiCompanionButtons_GetIniPath()
    for action in AiCompanionButtons_Actions() {
        raw := ""
        try raw := IniRead(iniPath, companion, action.id, "")
        catch
            raw := ""
        if (raw = "ERROR")
            raw := ""
        store[action.id] := AiCompanionButtons_Parse(raw)
    }
    g_AiCompanionButtonsCache[companion] := store
    return store
}

AiCompanionButtons_Get(companion, actionId) {
    if (!AiCompanionButtons_ActionById(actionId))
        return AiCompanionButtons_EmptySpec()
    store := AiCompanionButtons_Load(companion)
    if (store.Has(actionId))
        return store[actionId]
    return AiCompanionButtons_EmptySpec()
}

AiCompanionButtons_Save(companion, actionId, spec) {
    global g_AiCompanionButtonsCache
    if !AiCompanionModels_IsValidCompanion(companion)
        return false
    if (!AiCompanionButtons_ActionById(actionId))
        return false
    if (!AiCompanionButtons_HasSpec(spec))
        return false
    iniPath := AiCompanionButtons_GetIniPath()
    try DirCreate(A_ScriptDir "\assets\data")
    catch {
    }
    try IniWrite(AiCompanionButtons_Serialize(spec), iniPath, companion, actionId)
    catch
        return false
    store := AiCompanionButtons_Load(companion)
    store[actionId] := {
        name: Trim(spec.name),
        automationId: Trim(spec.automationId),
        className: Trim(spec.className),
        controlType: spec.controlType
    }
    g_AiCompanionButtonsCache[companion] := store
    return true
}

AiCompanionButtons_Clear(companion, actionId) {
    global g_AiCompanionButtonsCache
    if !AiCompanionModels_IsValidCompanion(companion)
        return false
    if (!AiCompanionButtons_ActionById(actionId))
        return false
    iniPath := AiCompanionButtons_GetIniPath()
    try IniDelete(iniPath, companion, actionId)
    catch {
    }
    store := AiCompanionButtons_Load(companion)
    store[actionId] := AiCompanionButtons_EmptySpec()
    g_AiCompanionButtonsCache[companion] := store
    return true
}

AiCompanionButtons_SpecFromElement(el) {
    spec := AiCompanionButtons_EmptySpec()
    if (!IsObject(el))
        return spec
    try spec.name := Trim(el.Name)
    catch {
    }
    try spec.automationId := Trim(el.AutomationId)
    catch {
    }
    try spec.className := Trim(el.ClassName)
    catch {
    }
    try spec.controlType := el.Type
    catch
        spec.controlType := 0
    return spec
}

; True when the element belongs to hwnd. A missing WinId stays allowed:
; the caller already checked the cursor window, and web nodes often have no hwnd.
; A different WinId means the walk left the companion window.
AiCompanionButtons_ElementInWindow(el, hwnd) {
    if (!IsObject(el) || !hwnd)
        return false
    try {
        wid := el.WinId
        if (wid)
            return wid = hwnd
    }
    return true
}

; Deepest element at the cursor, walked up to the nearest Button or Hyperlink.
AiCompanionButtons_ButtonFromPoint(hwnd) {
    el := 0
    try el := UIA.SmallestElementFromPoint()
    catch
        return 0
    loop 16 {
        if (!IsObject(el))
            return 0
        if (!AiCompanionButtons_ElementInWindow(el, hwnd))
            return 0
        t := 0
        try t := el.Type
        catch
            t := 0
        ; Button 50000, Hyperlink 50005. Window 50032 is the walk ceiling.
        if (t = 50000 || t = 50005)
            return el
        if (t = 50032)
            return 0
        parent := 0
        try parent := el.WalkTree("p")
        catch
            return 0
        if (!IsObject(parent))
            return 0
        el := parent
    }
    return 0
}

AiCompanionButtons_TryFind(uia, criteria) {
    if (!IsObject(uia) || !IsObject(criteria))
        return 0
    try {
        el := uia.FindFirst(criteria)
        if (el)
            return el
    }
    return 0
}

; Automation id, then name + control type, then class substring + control type.
; 0 lets the caller run its built-in finder.
AiCompanionButtons_FindSaved(uia, companion, actionId) {
    if (!AiCompanionButtons_ActionById(actionId))
        return 0
    if !AiCompanionModels_IsValidCompanion(companion)
        return 0
    return UiElements_TrySaved(uia, companion, actionId)
}

AiCompanionButtons_Click(el) {
    if (!IsObject(el))
        return false
    try {
        if (el.GetPropertyValue(UIA.Property.IsInvokePatternAvailable)) {
            el.InvokePattern.Invoke()
            return true
        }
    } catch {
    }
    try {
        el.Click()
        return true
    } catch {
    }
    return false
}
