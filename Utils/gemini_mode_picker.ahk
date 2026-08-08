; =============================================================================
; Utils module: gemini_mode_picker.ahk
; Gemini mode picker mouse + UIA
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; --- Gemini mode picker (1.5 Flash-Lite / 1.5 Flash / 1.5 Pro + Extended thinking) ----
; UI tree reference: gemini-no-context-menu.md
; Family matching absorbs renames (1.5 / 3.1 / plain Flash-Lite, Pro Extended, etc.).

global GEMINI_MODEL_CANONICAL_NAMES := ["1.5 Flash-Lite", "1.5 Flash", "1.5 Pro", "Extended thinking",
    "3.1 Flash-Lite", "3.5 Flash", "3.1 Pro", "Thinking level"]
global GEMINI_MODE_PICKER_NAME_SUBSTR := "Open mode picker"
global GEMINI_MODE_MENU_WAIT_MS := 450
global GEMINI_MODE_MENU_POLL_MS := 40
global GEMINI_MODE_PICKER_LABEL_WAIT_MS := 700
global GEMINI_MODEL_SELECT_MAX_CYCLES := 2

FindGeminiChromeHwnd() {
    for hwnd in WinGetList("ahk_exe chrome.exe") {
        try {
            if IsConsumerGeminiChromeTitle(WinGetTitle("ahk_id " hwnd))
                return hwnd
        } catch {
        }
    }
    return 0
}

; One UIA_Browser bind per flow. lightAttach: skip activate/a11y when caller already focused Gemini.
GeminiAttachBrowser(geminiHwnd := 0, lightAttach := false) {
    if (!geminiHwnd)
        geminiHwnd := FindGeminiChromeHwnd()
    if (!geminiHwnd || !WinExist("ahk_id " geminiHwnd))
        return 0
    winTitle := "ahk_id " geminiHwnd
    if (!lightAttach) {
        if (!WinActive(winTitle)) {
            try WinActivate(winTitle)
            try WinWaitActive(winTitle, , 1)
        }
        try UIA.ActivateChromiumAccessibility(winTitle, 300)
        catch {
        }
    }
    try {
        return UIA_Browser(winTitle)
    } catch {
        return 0
    }
}

GeminiFindRenderWidgetControl(browserHwnd) {
    for ctrlName in ["Chrome_RenderWidgetHostHWND1", "Chrome_RenderWidgetHostHWND"] {
        try {
            if ControlGetHwnd(ctrlName, "ahk_id " browserHwnd)
                return ctrlName
        } catch {
        }
    }
    try {
        for ctrlName in WinGetControls("ahk_id " browserHwnd) {
            if InStr(ctrlName, "Chrome_RenderWidgetHostHWND")
                return ctrlName
        }
    } catch {
    }
    return ""
}

GeminiGetElementWindowClickCoords(el, browserHwnd, uia := 0) {
    if !IsObject(el) || !browserHwnd
        return 0
    elPos := 0
    try {
        elPos := el.GetPos("window", browserHwnd)
    } catch {
        return 0
    }
    if (!IsObject(elPos) || elPos.w <= 0 || elPos.h <= 0)
        return 0
    cx := elPos.x + elPos.w // 2
    cy := elPos.y + elPos.h // 2
    docOffset := "none"
    if IsObject(uia) {
        try {
            doc := uia.GetCurrentDocumentElement()
            if IsObject(doc) {
                docPos := doc.GetPos("window", browserHwnd)
                if (IsObject(docPos) && docPos.w > 0 && docPos.h > 0) {
                    cx := elPos.x - docPos.x + elPos.w // 2
                    cy := elPos.y - docPos.y + elPos.h // 2
                    docOffset := "yes"
                }
            }
        } catch {
        }
    }
    return { cx: cx, cy: cy, docOffset: docOffset, elPosX: elPos.x, elPosY: elPos.y }
}

; skipActivate: caller already activated Gemini (hot path for mode picker).
GeminiMouseClickElement(el, browserHwnd := 0, uia := 0, skipActivate := false) {
    if !IsObject(el)
        return false
    winTitle := browserHwnd ? "ahk_id " browserHwnd : ""
    if (browserHwnd && !skipActivate) {
        if (!WinExist(winTitle))
            return false
        if (!WinActive(winTitle)) {
            WinActivate(winTitle)
            WinWaitActive(winTitle, , 1)
        }
    }
    if (browserHwnd) {
        try {
            el.ControlClick("left", 1, "", winTitle)
            return true
        } catch {
        }
        try {
            if (el.Click())
                return true
        } catch {
        }
        try {
            pt := el.GetClickablePoint()
            prevCoordMode := A_CoordModeMouse
            CoordMode("Mouse", "Screen")
            Click(pt.x, pt.y)
            CoordMode("Mouse", prevCoordMode)
            return true
        } catch {
        }
        coords := GeminiGetElementWindowClickCoords(el, browserHwnd, uia)
        if IsObject(coords) {
            renderCtrl := GeminiFindRenderWidgetControl(browserHwnd)
            if (renderCtrl != "") {
                try {
                    ControlGetPos(&rwx, &rwy, , , renderCtrl, winTitle)
                    ControlClick("X" . (coords.cx - rwx) . " Y" . (coords.cy - rwy), winTitle, renderCtrl)
                    return true
                } catch {
                }
            }
        }
    }
    return false
}

GeminiGetBrowserHwndFromUia(uia) {
    if !IsObject(uia)
        return 0
    try {
        hwnd := uia.BrowserId
        if (hwnd && WinExist("ahk_id " hwnd))
            return hwnd
    } catch {
    }
    return WinExist("ahk_exe chrome.exe") ? WinExist("ahk_exe chrome.exe") : 0
}

GeminiDismissModePickerMenu(uia, browserHwnd := 0) {
    if !IsObject(uia)
        return false
    if (!browserHwnd)
        browserHwnd := GeminiGetBrowserHwndFromUia(uia)
    promptField := FindGeminiPromptField(uia)
    if !promptField
        return false
    return GeminiMouseClickElement(promptField, browserHwnd, uia, true)
}

FindGeminiModePickerButton(uia) {
    if !IsObject(uia)
        return 0
    for typeSpec in [50000, "Button"] {
        try {
            el := uia.FindFirst({ Type: typeSpec, Name: GEMINI_MODE_PICKER_NAME_SUBSTR, mm: 2, cs: 0 })
            if el
                return el
        } catch {
        }
    }
    ; Modern labels: "Pro Extended", "Flash-Lite", "1.5 Flash", etc. (no "Open mode picker" prefix).
    for typeSpec in [50000, "Button"] {
        try {
            buttons := uia.FindAll({ Type: typeSpec })
        } catch {
            continue
        }
        if !IsObject(buttons)
            continue
        for btn in buttons {
            try {
                n := Trim(btn.Name)
            } catch {
                continue
            }
            if (n = "")
                continue
            if RegExMatch(n, "i)mode\s*picker|currently\s+")
                return btn
            if (GeminiModelFamily(n) != "" && !RegExMatch(n, "i)microphone|Upload|tools|Search|Main menu"))
                return btn
        }
    }
    return 0
}

GeminiGetActiveModelFromPickerElement(picker) {
    if !IsObject(picker)
        return ""
    raw := ""
    try raw := Trim(picker.Name)
    catch {
        raw := ""
    }
    if (raw = "")
        return ""
    if RegExMatch(raw, "i)currently\s+(.+)$", &m) {
        norm := GeminiNormalizeModelLabel(Trim(m[1]))
        if (norm != "")
            return norm
    }
    ; Modern button labels may be plain "Pro Extended", "Flash-Lite", "1.5 Pro", etc.
    return GeminiNormalizeModelLabel(raw)
}

; Stable family key for version-tolerant compares (flash-lite / flash / pro / extended-thinking).
GeminiModelFamily(name) {
    name := Trim(name)
    if (name = "")
        return ""
    if RegExMatch(name, "i)extended\s*thinking|thinking\s*level")
        return "extended-thinking"
    if RegExMatch(name, "i)Flash-Lite|Flash Lite")
        return "flash-lite"
    if RegExMatch(name, "i)\bPro\b")
        return "pro"
    if RegExMatch(name, "i)\bFlash\b")
        return "flash"
    return ""
}

GeminiModelsMatch(a, b) {
    fa := GeminiModelFamily(a)
    fb := GeminiModelFamily(b)
    if (fa != "" && fb != "" && fa = fb)
        return true
    a := Trim(a)
    b := Trim(b)
    return (a != "" && b != "" && (a = b || InStr(a, b, false) || InStr(b, a, false)))
}

FindGeminiModelMenuItem(uia, modelName) {
    if !IsObject(uia)
        return 0
    needles := GeminiModelMenuNeedles(modelName)
    if (needles.Length = 0)
        return 0
    for needle in needles {
        if (needle = "")
            continue
        for typeSpec in [50011, "MenuItem", 50008, "RadioButton", 50003, "ListItem", 50000, "Button"] {
            try {
                el := uia.FindFirst({ Type: typeSpec, Name: needle, mm: 3, cs: 0 })
                if el
                    return el
            } catch {
            }
            try {
                el := uia.FindFirst({ Type: typeSpec, Name: needle, mm: 2, cs: 0 })
                if el
                    return el
            } catch {
            }
        }
    }
    return 0
}

; Gate D: scan all candidate rows and pick best family / substring match.
FindGeminiModelMenuItemByScan(uia, modelName) {
    if !IsObject(uia)
        return 0
    family := GeminiModelFamily(modelName)
    exp := GeminiNormalizeModelLabel(modelName)
    raw := Trim(modelName)
    bestEl := 0
    bestScore := 0
    for typeSpec in [50011, "MenuItem", 50008, "RadioButton", 50003, "ListItem", 50000, "Button"] {
        try {
            items := uia.FindAll({ Type: typeSpec })
        } catch {
            continue
        }
        if !IsObject(items)
            continue
        for mi in items {
            try {
                fullName := Trim(mi.Name)
            } catch {
                continue
            }
            if (fullName = "")
                continue
            score := 0
            if (exp != "" && (fullName = exp || InStr(fullName, exp, false)))
                score += 50
            if (raw != "" && InStr(fullName, raw, false))
                score += 40
            itemFam := GeminiModelFamily(fullName)
            if (family != "" && itemFam = family)
                score += 30
            if (score > bestScore) {
                bestScore := score
                bestEl := mi
            }
        }
        if (bestEl && bestScore >= 30)
            return bestEl
    }
    return (bestScore >= 20) ? bestEl : 0
}

GeminiModelMenuNeedles(modelName) {
    needles := []
    raw := Trim(modelName)
    if (raw != "")
        needles.Push(raw)
    norm := GeminiNormalizeModelLabel(raw)
    if (norm != "" && norm != raw)
        needles.Push(norm)
    family := GeminiModelFamily(raw)
    if (family = "flash-lite") {
        needles.Push("1.5 Flash-Lite")
        needles.Push("3.1 Flash-Lite")
        needles.Push("Flash-Lite")
    } else if (family = "flash") {
        needles.Push("1.5 Flash")
        needles.Push("3.5 Flash")
        needles.Push("3.1 Flash")
        needles.Push("Flash")
    } else if (family = "pro") {
        needles.Push("1.5 Pro")
        needles.Push("3.1 Pro")
        needles.Push("Pro")
    } else if (family = "extended-thinking") {
        needles.Push("Extended thinking")
        needles.Push("Thinking level")
    }
    ; De-dupe while preserving order.
    out := []
    seen := Map()
    for n in needles {
        key := StrLower(n)
        if seen.Has(key)
            continue
        seen[key] := true
        out.Push(n)
    }
    return out
}

GeminiWaitForModelMenuItem(uia, modelName, timeoutMs := 0) {
    if (timeoutMs <= 0)
        timeoutMs := GEMINI_MODE_MENU_WAIT_MS
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        el := FindGeminiModelMenuItem(uia, modelName)
        if el
            return el
        el := FindGeminiModelMenuItemByScan(uia, modelName)
        if el
            return el
        Sleep GEMINI_MODE_MENU_POLL_MS
    }
    return 0
}

GeminiWaitForPickerShowsModel(uia, expected, timeoutMs := 0) {
    if (timeoutMs <= 0)
        timeoutMs := GEMINI_MODE_PICKER_LABEL_WAIT_MS
    if (Trim(expected) = "")
        return false
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        if (GeminiPickerShowsModel(uia, expected))
            return true
        Sleep GEMINI_MODE_MENU_POLL_MS
    }
    return false
}

GeminiPickerShowsModel(uia, expected) {
    picker := FindGeminiModePickerButton(uia)
    if !picker
        return false
    active := GeminiGetActiveModelFromPickerElement(picker)
    if (GeminiModelsMatch(active, expected))
        return true
    ; Raw button name sometimes is the status itself ("Pro Extended").
    try {
        return GeminiModelsMatch(picker.Name, expected)
    } catch {
        return false
    }
}

GeminiNormalizeModelLabel(name) {
    if (name = "")
        return ""
    fam := GeminiModelFamily(name)
    if (fam = "extended-thinking")
        return "Extended thinking"
    if (fam = "flash-lite")
        return "1.5 Flash-Lite"
    if (fam = "pro")
        return "1.5 Pro"
    if (fam = "flash")
        return "1.5 Flash"
    return ""
}

; Click a menu item: prefer Invoke, then mouse/ControlClick stack.
GeminiActivateModelMenuItem(el, browserHwnd, uia) {
    if !IsObject(el)
        return false
    try {
        if (el.GetPropertyValue(UIA.Property.IsInvokePatternAvailable)) {
            el.InvokePattern.Invoke()
            return true
        }
    } catch {
    }
    try {
        if (el.GetPropertyValue(UIA.Property.IsSelectionItemPatternAvailable)) {
            el.SelectionItemPattern.Select()
            return true
        }
    } catch {
    }
    return GeminiMouseClickElement(el, browserHwnd, uia, true)
}

GetGeminiActiveModelFromPickerOnly(uia) {
    return GeminiGetActiveModelFromPickerElement(FindGeminiModePickerButton(uia))
}

GeminiCollectModelMenuItemState(mi) {
    className := ""
    try {
        className := mi.ClassName
    } catch {
        className := ""
    }
    isDisabled := false
    try {
        if (InStr(className, "disabled") || InStr(className, "mat-mdc-button-disabled"))
            isDisabled := true
        try {
            if (!mi.GetPropertyValue(UIA.Property.IsEnabled))
                isDisabled := true
        } catch {
        }
    } catch {
    }
    isSelected := false
    try {
        if (mi.GetPropertyValue(UIA.Property.IsSelectionItemPatternAvailable))
            isSelected := mi.SelectionItemPattern.IsSelected
    } catch {
    }
    hasLegacy := false
    legacyState := -1
    try {
        hasLegacy := mi.GetPropertyValue(UIA.Property.IsLegacyIAccessiblePatternAvailable)
    } catch {
        hasLegacy := false
    }
    if (hasLegacy) {
        try {
            legacyState := mi.LegacyIAccessiblePattern.State
        } catch {
            try {
                legacyState := mi.LegacyIAccessiblePattern.CurrentState
            } catch {
                legacyState := -1
            }
        }
    }
    if (!isSelected) {
        try {
            if (InStr(className, "is-selected") || InStr(className, "selected") || InStr(className, "active") ||
            InStr(className, "mdc-selected"))
                isSelected := true
        } catch {
        }
    }
    if (!isSelected) {
        try {
            if (hasLegacy && legacyState != -1 && (legacyState & 0x4))
                isSelected := true
        } catch {
        }
    }
    return { isDisabled: isDisabled, isSelected: isSelected, className: className }
}

GeminiCollectModelMenuItems(uia) {
    modelButtons := []
    menuItems := []
    for typeSpec in [50011, "MenuItem", 50008, "RadioButton", 50003, "ListItem"] {
        try {
            found := uia.FindAll({ Type: typeSpec })
        } catch {
            continue
        }
        if (IsObject(found) && found.Length) {
            menuItems := found
            break
        }
    }
    for mi in menuItems {
        try {
            fullName := mi.Name
            shortName := GeminiNormalizeModelLabel(fullName)
            if (shortName = "")
                continue
            st := GeminiCollectModelMenuItemState(mi)
            modelButtons.Push({ btn: mi, name: shortName, isSelected: st.isSelected, isDisabled: st.isDisabled,
                className: st.className })
        } catch {
        }
    }
    return modelButtons
}

; Back-compat alias for callers not yet updated
GeminiCollectModelOptionButtons(uia) {
    return GeminiCollectModelMenuItems(uia)
}

GeminiOpenModePickerMenu(uia, picker := "", browserHwnd := 0) {
    if !IsObject(uia)
        return false
    if (!browserHwnd)
        browserHwnd := GeminiGetBrowserHwndFromUia(uia)
    if !IsObject(picker)
        picker := FindGeminiModePickerButton(uia)
    if !picker
        return false
    return GeminiMouseClickElement(picker, browserHwnd, uia, true)
}

; Quality gates: already-selected → UIA Invoke → mouse stack → scan+click; verify picker each cycle.
EnsureGeminiModelViaMenu(expected, geminiHwnd := 0) {
    global GEMINI_MODEL_SELECT_MAX_CYCLES
    if (GeminiModelFamily(expected) = "extended-thinking")
        return EnsureGeminiExtendedThinkingToggle(geminiHwnd)

    target := GeminiNormalizeModelLabel(expected)
    if (target = "")
        target := Trim(expected)
    if (target = "")
        return false

    uia := GeminiAttachBrowser(geminiHwnd)
    if !IsObject(uia)
        return false
    browserHwnd := geminiHwnd ? geminiHwnd : GeminiGetBrowserHwndFromUia(uia)
    cycles := GEMINI_MODEL_SELECT_MAX_CYCLES > 0 ? GEMINI_MODEL_SELECT_MAX_CYCLES : 2

    loop cycles {
        ; Gate A — picker already shows family
        if (GeminiPickerShowsModel(uia, target))
            return true

        picker := FindGeminiModePickerButton(uia)
        if !picker {
            uia := GeminiAttachBrowser(geminiHwnd)
            if !IsObject(uia)
                return false
            picker := FindGeminiModePickerButton(uia)
            if !picker
                return false
        }

        if (!GeminiOpenModePickerMenu(uia, picker, browserHwnd)) {
            GeminiDismissModePickerMenu(uia, browserHwnd)
            continue
        }

        ; Gate B — FindFirst needles + Invoke/Select
        targetBtn := GeminiWaitForModelMenuItem(uia, target)
        clicked := false
        if (targetBtn)
            clicked := GeminiActivateModelMenuItem(targetBtn, browserHwnd, uia)

        ; Gate C — mouse/ControlClick stack if Invoke path failed
        if (!clicked && targetBtn)
            clicked := GeminiMouseClickElement(targetBtn, browserHwnd, uia, true)

        ; Gate D — reopen scan by family / substring
        if (!clicked) {
            scanBtn := FindGeminiModelMenuItemByScan(uia, target)
            if (scanBtn) {
                clicked := GeminiActivateModelMenuItem(scanBtn, browserHwnd, uia)
                if (!clicked)
                    clicked := GeminiMouseClickElement(scanBtn, browserHwnd, uia, true)
            }
        }

        if (!clicked) {
            GeminiDismissModePickerMenu(uia, browserHwnd)
            Sleep 80
            uia := GeminiAttachBrowser(geminiHwnd, true)
            continue
        }

        ; Verify picker label (family match)
        if (GeminiWaitForPickerShowsModel(uia, target))
            return true

        ; One verify miss: dismiss and retry full cycle
        GeminiDismissModePickerMenu(uia, browserHwnd)
        Sleep 100
        uia := GeminiAttachBrowser(geminiHwnd)
        if !IsObject(uia)
            return false
    }
    return GeminiPickerShowsModel(uia, target)
}

; Toggle Extended thinking (also covers legacy Thinking level label).
EnsureGeminiExtendedThinkingToggle(geminiHwnd := 0) {
    uia := GeminiAttachBrowser(geminiHwnd)
    if !IsObject(uia)
        return false
    browserHwnd := geminiHwnd ? geminiHwnd : GeminiGetBrowserHwndFromUia(uia)
    picker := FindGeminiModePickerButton(uia)
    if (!picker || !GeminiOpenModePickerMenu(uia, picker, browserHwnd)) {
        GeminiDismissModePickerMenu(uia, browserHwnd)
        return false
    }
    deadline := A_TickCount + Max(GEMINI_MODE_MENU_WAIT_MS, 800)
    thinkingItem := 0
    while (A_TickCount < deadline) {
        thinkingItem := FindGeminiModelMenuItem(uia, "Extended thinking")
        if !thinkingItem
            thinkingItem := FindGeminiModelMenuItemByScan(uia, "Extended thinking")
        if thinkingItem
            break
        Sleep GEMINI_MODE_MENU_POLL_MS
    }
    if !thinkingItem {
        GeminiDismissModePickerMenu(uia, browserHwnd)
        return false
    }
    ok := GeminiActivateModelMenuItem(thinkingItem, browserHwnd, uia)
    if (!ok)
        ok := GeminiMouseClickElement(thinkingItem, browserHwnd, uia, true)
    if (!ok) {
        GeminiDismissModePickerMenu(uia, browserHwnd)
        return false
    }
    ; Toggle may leave menu open or close it; dismiss so composer is usable.
    Sleep 120
    try GeminiDismissModePickerMenu(uia, browserHwnd)
    catch {
    }
    return true
}

; Back-compat alias
EnsureGeminiThinkingLevelMenuOpen(geminiHwnd := 0) {
    return EnsureGeminiExtendedThinkingToggle(geminiHwnd)
}
