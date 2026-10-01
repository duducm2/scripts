; =============================================================================
; Lib: UiElements.ahk
; Saved UIA identity for cataloged controls. A saved spec is tried first;
; the caller's built-in finder stays in place when nothing is saved.
; Store: assets/data/ui_elements.ini
; =============================================================================

global g_UiElementsCache := Map()
global g_UiElementsGui := 0
global g_UiElementsLv := 0
global g_UiElementsFilter := 0
global g_UiElementsActive := false
global g_UiElementsAppKey := ""
global g_UiElementsHwnd := 0
global g_UiElementsDisplay := ""
global g_UiElementsRowMeta := Map()
global g_UiElementsCaptureActive := false
global g_UiElementsCaptureArmed := false
global g_UiElementsCaptureEntry := 0
global g_UiElementsCaptureEscPrev := false
global g_UiElementsCaptureEscBound := false
global g_UiElementsCaptureEscSwallow := false

UiElements_IniPath() {
    return A_ScriptDir "\assets\data\ui_elements.ini"
}

UiElements_EnsureImported() {
    static done := false
    if (done)
        return
    done := true
    newPath := UiElements_IniPath()
    flag := ""
    try flag := IniRead(newPath, "UiElements", "imported", "")
    catch
        flag := ""
    if (flag = "ERROR")
        flag := ""
    if (Trim(flag) = "1")
        return
    oldPath := A_ScriptDir "\assets\data\ai_companion_buttons.ini"
    try DirCreate(A_ScriptDir "\assets\data")
    catch {
    }
    for section in ["Gemini", "GeminiEnterprise", "CopilotWeb"] {
        for key in ["NewChat", "Menu", "Search", "Send"] {
            current := ""
            try current := IniRead(newPath, section, key, "")
            catch
                current := ""
            if (current = "ERROR")
                current := ""
            if (Trim(current) != "")
                continue
            old := ""
            try old := IniRead(oldPath, section, key, "")
            catch
                old := ""
            if (old = "ERROR" || Trim(old) = "")
                continue
            try IniWrite(old, newPath, section, key)
            catch {
            }
        }
    }
    try IniWrite("1", newPath, "UiElements", "imported")
    catch {
    }
}

UiElements_IsCompanionAction(section, elementId) {
    return AiCompanionModels_IsValidCompanion(section) && AiCompanionButtons_ActionById(elementId)
}

UiElements_Get(section, elementId) {
    global g_UiElementsCache
    UiElements_EnsureImported()
    section := Trim(section)
    elementId := Trim(elementId)
    if (section = "" || elementId = "")
        return AiCompanionButtons_EmptySpec()
    if (UiElements_IsCompanionAction(section, elementId))
        return AiCompanionButtons_Get(section, elementId)
    cacheKey := section
    if (!g_UiElementsCache.Has(cacheKey))
        g_UiElementsCache[cacheKey] := Map()
    bucket := g_UiElementsCache[cacheKey]
    if (bucket.Has(elementId))
        return bucket[elementId]
    raw := ""
    try raw := IniRead(UiElements_IniPath(), section, elementId, "")
    catch
        raw := ""
    if (raw = "ERROR")
        raw := ""
    spec := AiCompanionButtons_Parse(raw)
    bucket[elementId] := spec
    return spec
}

UiElements_Save(section, elementId, spec) {
    global g_UiElementsCache
    if (!AiCompanionButtons_HasSpec(spec))
        return false
    if (UiElements_IsCompanionAction(section, elementId))
        return AiCompanionButtons_Save(section, elementId, spec)
    UiElements_EnsureImported()
    try DirCreate(A_ScriptDir "\assets\data")
    catch {
    }
    try IniWrite(AiCompanionButtons_Serialize(spec), UiElements_IniPath(), section, elementId)
    catch
        return false
    if (!g_UiElementsCache.Has(section))
        g_UiElementsCache[section] := Map()
    g_UiElementsCache[section][elementId] := {
        name: Trim(spec.name),
        automationId: Trim(spec.automationId),
        className: Trim(spec.className),
        controlType: spec.controlType
    }
    return true
}

; Saved element, or 0 so the caller runs its built-in finder.
UiElements_TrySaved(root, section, elementId) {
    if (!IsObject(root))
        return 0
    spec := UiElements_Get(section, elementId)
    if (!AiCompanionButtons_HasSpec(spec))
        return 0
    ct := 0
    try ct := Integer(spec.controlType)
    catch
        ct := 0

    if (Trim(spec.automationId) != "") {
        criteria := { AutomationId: spec.automationId }
        if (ct > 0)
            criteria.Type := ct
        el := AiCompanionButtons_TryFind(root, criteria)
        if (el)
            return el
        if (ct > 0) {
            el := AiCompanionButtons_TryFind(root, { AutomationId: spec.automationId })
            if (el)
                return el
        }
    }
    if (Trim(spec.name) != "") {
        if (ct > 0) {
            el := AiCompanionButtons_TryFind(root, { Name: spec.name, Type: ct })
            if (el)
                return el
        }
        el := AiCompanionButtons_TryFind(root, { Name: spec.name })
        if (el)
            return el
    }
    if (Trim(spec.className) != "") {
        if (ct > 0) {
            el := AiCompanionButtons_TryFind(root, { ClassName: spec.className, matchmode: "Substring", Type: ct })
            if (el)
                return el
        }
        el := AiCompanionButtons_TryFind(root, { ClassName: spec.className, matchmode: "Substring" })
        if (el)
            return el
    }
    return 0
}

; Saved spec, then an optional finder that receives the same root.
UiElements_Find(root, section, elementId, finder?) {
    hit := UiElements_TrySaved(root, section, elementId)
    if (IsObject(hit))
        return hit
    if (IsSet(finder) && HasMethod(finder, "Call")) {
        try return finder.Call(root)
        catch
            return 0
    }
    return 0
}

UiElements_Summary(section, elementId) {
    spec := UiElements_Get(section, elementId)
    return AiCompanionButtons_Summary(spec)
}

; ---------- Manager ----------------------------------------------------------

UiElements_OpenManager(hwnd, appKey) {
    global g_UiElementsActive, g_UiElementsAppKey, g_UiElementsHwnd, g_UiElementsDisplay
    UiElements_CaptureStop()
    try StandardLoadingBar_Hide(0)
    g_UiElementsAppKey := Trim(appKey)
    g_UiElementsHwnd := hwnd
    g_UiElementsDisplay := g_UiElementsAppKey
    if (g_UiElementsDisplay = "") {
        try g_UiElementsDisplay := WinGetProcessName("ahk_id " hwnd)
        catch
            g_UiElementsDisplay := "this app"
    }
    g_UiElementsActive := true
    UiElements_Rebuild()
}

UiElements_Close(*) {
    global g_UiElementsActive, g_UiElementsGui, g_UiElementsLv, g_UiElementsFilter
    if (g_UiElementsCaptureActive)
        UiElements_CaptureStop()
    UiElements_UnbindEnter()
    UiElements_SafeDestroyGui(g_UiElementsGui)
    g_UiElementsGui := 0
    g_UiElementsLv := 0
    g_UiElementsFilter := 0
    g_UiElementsActive := false
}

UiElements_SafeDestroyGui(guiObj) {
    if (!IsObject(guiObj))
        return
    try guiObj.Destroy()
    catch {
    }
}

UiElements_Rebuild() {
    global g_UiElementsGui, g_UiElementsLv, g_UiElementsFilter, g_UiElementsActive, g_UiElementsDisplay,
        g_UiElementsAppKey, g_UiElementsRowMeta

    if (!g_UiElementsActive)
        return
    query := ""
    if (IsObject(g_UiElementsFilter)) {
        try query := g_UiElementsFilter.Value
        catch
            query := ""
    }
    UiElements_UnbindEnter()
    UiElements_SafeDestroyGui(g_UiElementsGui)
    g_UiElementsGui := 0
    g_UiElementsLv := 0
    g_UiElementsFilter := 0

    title := "UI elements — " . g_UiElementsDisplay
    g_UiElementsGui := Gui("+AlwaysOnTop +ToolWindow", title)
    g_UiElementsGui.SetFont("s10", "Segoe UI")
    g_UiElementsGui.Add("Text", "w760 Wrap",
        "Enter recaptures the selected control. Esc closes. Saved identity is tried before the built-in finder.")
    g_UiElementsFilter := g_UiElementsGui.Add("Edit", "w760")
    g_UiElementsFilter.OnEvent("Change", UiElements_OnFilter)
    g_UiElementsLv := g_UiElementsGui.Add("ListView", "w760 h420 -Multi", ["Element", "Saved", "Source"])
    g_UiElementsLv.OnEvent("DoubleClick", UiElements_BeginSelected)
    g_UiElementsGui.Add("Button", "w120 Default", "Recapture").OnEvent("Click", UiElements_BeginSelected)
    g_UiElementsGui.Add("Button", "w100 x+8", "Close").OnEvent("Click", UiElements_Close)
    g_UiElementsGui.OnEvent("Close", UiElements_Close)
    g_UiElementsGui.OnEvent("Escape", UiElements_Close)

    g_UiElementsRowMeta := Map()
    UiElements_FillList(query)
    try g_UiElementsLv.ModifyCol(1, 220)
    try g_UiElementsLv.ModifyCol(2, 260)
    try g_UiElementsLv.ModifyCol(3, 260)
    g_UiElementsGui.Show("w800 h560")
    hwnd := 0
    try hwnd := g_UiElementsGui.Hwnd
    if (hwnd)
        UiElements_BindEnter(hwnd)
    if (query != "") {
        try g_UiElementsFilter.Value := query
        try g_UiElementsFilter.Focus()
    } else {
        try g_UiElementsLv.Focus()
    }
}

UiElements_OnFilter(*) {
    global g_UiElementsFilter, g_UiElementsActive
    if (!g_UiElementsActive || !IsObject(g_UiElementsFilter))
        return
    query := ""
    try query := g_UiElementsFilter.Value
    UiElements_FillList(query)
}

UiElements_FillList(query) {
    global g_UiElementsLv, g_UiElementsAppKey, g_UiElementsRowMeta
    if (!IsObject(g_UiElementsLv))
        return
    g_UiElementsLv.Delete()
    g_UiElementsRowMeta := Map()
    needle := StrLower(Trim(query))
    shown := 0
    for entry in UiElementCatalog_ForApp(g_UiElementsAppKey) {
        blob := StrLower(entry.label . " " . entry.id . " " . entry.source)
        if (needle != "" && !InStr(blob, needle))
            continue
        saved := UiElements_Summary(entry.section, entry.id)
        row := g_UiElementsLv.Add("", entry.label, saved, entry.source)
        g_UiElementsRowMeta[row] := entry
        shown++
    }
    if (shown = 0) {
        msg := (g_UiElementsAppKey = "") ? "No cheat sheet for this app." : "No cataloged UI elements for this app."
        g_UiElementsLv.Add("", msg, "", "")
    } else {
        try g_UiElementsLv.Modify(1, "Select Focus")
    }
}

UiElements_SelectedEntry() {
    global g_UiElementsLv, g_UiElementsRowMeta
    if (!IsObject(g_UiElementsLv))
        return 0
    row := 0
    try row := g_UiElementsLv.GetNext(0, "F")
    if (!row)
        return 0
    if (!g_UiElementsRowMeta.Has(row))
        return 0
    return g_UiElementsRowMeta[row]
}

UiElements_BindEnter(hwnd) {
    try HotIf((*) => WinActive("ahk_id " hwnd))
    try Hotkey("Enter", UiElements_BeginSelected, "On")
    try HotIf()
    catch {
        try HotIf()
    }
}

UiElements_UnbindEnter() {
    global g_UiElementsGui
    if (!IsObject(g_UiElementsGui))
        return
    hwnd := 0
    try hwnd := g_UiElementsGui.Hwnd
    if (!hwnd)
        return
    try HotIf((*) => WinActive("ahk_id " hwnd))
    try Hotkey("Enter", "Off")
    try HotIf()
    catch {
        try HotIf()
    }
}

UiElements_BeginSelected(*) {
    global g_UiElementsActive, g_UiElementsCaptureActive, g_UiElementsHwnd, g_UiElementsGui, g_UiElementsDisplay
    if (!g_UiElementsActive || g_UiElementsCaptureActive)
        return
    entry := UiElements_SelectedEntry()
    if (!IsObject(entry))
        return
    hwnd := g_UiElementsHwnd
    if (!hwnd || !WinExist("ahk_id " hwnd)) {
        try ShowCenteredOverlay_Utils("That window is no longer open.", 2200, BANNER_ACCENT_ERROR)
        return
    }
    UiElements_UnbindEnter()
    UiElements_SafeDestroyGui(g_UiElementsGui)
    g_UiElementsGui := 0
    try WinActivate("ahk_id " hwnd)
    UiElements_CaptureStart(entry, hwnd)
}

; ---------- Capture ----------------------------------------------------------

UiElements_EscapeIsDown() {
    return GetKeyState("Escape", "P") || ((DllCall("user32\GetAsyncKeyState", "int", 0x1B) & 0x8000) != 0)
}

UiElements_CaptureBindEscape() {
    global g_OnEscapePressed, g_UiElementsCaptureEscBound, g_UiElementsCaptureEscPrev
    if (g_UiElementsCaptureEscBound)
        return
    ; Rising edge starts from the current key, so a held Escape does not cancel immediately.
    g_UiElementsCaptureEscPrev := UiElements_EscapeIsDown()
    try HotIf()
    catch {
    }
    try Hotkey("$*Escape", UiElements_CaptureAbort, "On")
    catch {
    }
    try HotIf()
    catch {
    }
    g_OnEscapePressed := UiElements_CaptureAbort
    try Utils_EnsureGlobalEscapeHotkey()
    catch {
    }
    g_UiElementsCaptureEscBound := true
}

UiElements_CaptureClearEscapeCallback() {
    global g_OnEscapePressed, g_UiElementsCaptureEscSwallow
    g_UiElementsCaptureEscSwallow := false
    if (g_OnEscapePressed = UiElements_CaptureAbort)
        g_OnEscapePressed := ""
    try Utils_EnsureGlobalEscapeHotkey()
    catch {
    }
}

; The same physical press can reach both $*Escape and the I10 handler. Leave the
; callback in place until the key is up so the later path still returns true.
UiElements_CaptureEscapeRelease() {
    if (UiElements_EscapeIsDown())
        return
    SetTimer(UiElements_CaptureEscapeRelease, 0)
    UiElements_CaptureClearEscapeCallback()
}

UiElements_CaptureUnbindEscape() {
    global g_OnEscapePressed, g_UiElementsCaptureEscBound, g_UiElementsCaptureEscPrev, g_UiElementsCaptureEscSwallow
    if (!g_UiElementsCaptureEscBound)
        return
    g_UiElementsCaptureEscBound := false
    g_UiElementsCaptureEscPrev := false
    try HotIf()
    catch {
    }
    try Hotkey("$*Escape", "Off")
    catch {
    }
    try HotIf()
    catch {
    }
    if (UiElements_EscapeIsDown() && g_OnEscapePressed = UiElements_CaptureAbort) {
        g_UiElementsCaptureEscSwallow := true
        SetTimer(UiElements_CaptureEscapeRelease, 30)
        return
    }
    UiElements_CaptureClearEscapeCallback()
}

UiElements_CaptureStart(entry, hwnd) {
    global g_UiElementsCaptureActive, g_UiElementsCaptureArmed, g_UiElementsCaptureEntry, g_UiElementsDisplay
    g_UiElementsCaptureActive := true
    g_UiElementsCaptureArmed := false
    g_UiElementsCaptureEntry := entry
    UiElements_CaptureBindEscape()
    label := entry.label
    try StandardLoadingBar_Show(
        "Click " . label . " in " . g_UiElementsDisplay . ". Esc cancels.",
        BANNER_ACCENT_INTERMEDIATE, { passive: true, centerOnHwnd: hwnd, fontSize: 17, textWidth: 640,
            passiveBgColor: BANNER_ACCENT_INTERMEDIATE }
    )
    SetTimer(UiElements_CapturePoll, 30)
}

UiElements_CaptureStop() {
    global g_UiElementsCaptureActive, g_UiElementsCaptureArmed, g_UiElementsCaptureEntry
    g_UiElementsCaptureActive := false
    g_UiElementsCaptureArmed := false
    g_UiElementsCaptureEntry := 0
    try SetTimer(UiElements_CapturePoll, 0)
    UiElements_CaptureUnbindEscape()
}

UiElements_CaptureAbort(*) {
    global g_UiElementsCaptureActive, g_UiElementsActive, g_UiElementsCaptureEscSwallow
    if (!g_UiElementsCaptureActive)
        return g_UiElementsCaptureEscSwallow
    UiElements_CaptureStop()
    try StandardLoadingBar_Hide(0)
    if (g_UiElementsActive)
        UiElements_Rebuild()
    ; True tells the I10 Escape handler not to inject Escape into the target window.
    return true
}

UiElements_CapturePoll() {
    global g_UiElementsCaptureActive, g_UiElementsCaptureArmed, g_UiElementsCaptureEscPrev
    if (!g_UiElementsCaptureActive) {
        SetTimer(UiElements_CapturePoll, 0)
        return
    }
    escDown := UiElements_EscapeIsDown()
    if (escDown) {
        if (!g_UiElementsCaptureEscPrev) {
            g_UiElementsCaptureEscPrev := true
            UiElements_CaptureAbort()
            return
        }
    } else {
        g_UiElementsCaptureEscPrev := false
    }
    down := GetKeyState("LButton", "P") || ((DllCall("user32\GetAsyncKeyState", "int", 0x01) & 0x8000) != 0)
    if (!g_UiElementsCaptureArmed) {
        if (!down)
            g_UiElementsCaptureArmed := true
        return
    }
    if (down)
        UiElements_CaptureOnClick()
}

UiElements_CaptureOnClick() {
    global g_UiElementsCaptureActive, g_UiElementsCaptureEntry, g_UiElementsHwnd, g_UiElementsActive,
        g_UiElementsDisplay
    if (!g_UiElementsCaptureActive)
        return
    entry := g_UiElementsCaptureEntry
    hwnd := g_UiElementsHwnd
    UiElements_CaptureStop()
    label := IsObject(entry) ? entry.label : "element"
    if (!AiCompanionButtonCapture_CursorInTarget(hwnd)) {
        try StandardLoadingBar_Hide(0)
        try ShowCenteredOverlay_Utils("Click was outside " . g_UiElementsDisplay . ". Mapping unchanged.", 2200,
            BANNER_ACCENT_INTERMEDIATE)
        if (g_UiElementsActive)
            UiElements_Rebuild()
        return
    }
    kind := (IsObject(entry) && entry.HasProp("kind")) ? entry.kind : "button"
    el := UiElements_ElementFromClick(hwnd, kind)
    try StandardLoadingBar_Hide(0)
    if (!IsObject(el)) {
        try ShowCenteredOverlay_Utils("Could not read that element in " . g_UiElementsDisplay . ". Mapping unchanged.",
            2200, BANNER_ACCENT_INTERMEDIATE)
        if (g_UiElementsActive)
            UiElements_Rebuild()
        return
    }
    spec := AiCompanionButtons_SpecFromElement(el)
    if (!IsObject(entry) || !UiElements_Save(entry.section, entry.id, spec)) {
        try ShowCenteredOverlay_Utils("Could not save " . label . ".", 2200, BANNER_ACCENT_ERROR)
        if (g_UiElementsActive)
            UiElements_Rebuild()
        return
    }
    try ShowCenteredOverlay_Utils("Saved " . label . ": " . AiCompanionButtons_Summary(spec), 2000,
    BANNER_ACCENT_SUCCESS)
    if (g_UiElementsActive)
        UiElements_Rebuild()
}

UiElements_ElementFromClick(hwnd, kind) {
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
        if (kind = "button") {
            if (t = 50000 || t = 50005)
                return el
            if (t = 50032)
                return 0
        } else if (kind = "edit") {
            if (t = 50004 || t = 50003)
                return el
            if (t = 50032)
                return 0
        } else {
            name := ""
            aid := ""
            try name := Trim(el.Name)
            catch
                name := ""
            try aid := Trim(el.AutomationId)
            catch
                aid := ""
            ; Skip bare text and image nodes; keep the control that owns them.
            if (t != 50020 && t != 50025 && (name != "" || aid != ""))
                return el
            if (t = 50032)
                return 0
        }
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
