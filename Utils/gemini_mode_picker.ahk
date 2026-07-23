; =============================================================================
; Utils module: gemini_mode_picker.ahk
; Gemini mode picker mouse + UIA
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; --- Gemini mode picker (3.1 Flash-Lite / 3.5 Flash / 3.1 Pro), mouse + UIA ----
; UI tree reference: gemini-no-context-menu.md

global GEMINI_MODEL_CANONICAL_NAMES := ["3.1 Flash-Lite", "3.5 Flash", "3.1 Pro"]
global GEMINI_MODE_PICKER_NAME_SUBSTR := "Open mode picker"
global GEMINI_MODE_MENU_WAIT_MS := 450
global GEMINI_MODE_MENU_POLL_MS := 40
global GEMINI_MODE_PICKER_LABEL_WAIT_MS := 500

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
    return 0
}

GeminiGetActiveModelFromPickerElement(picker) {
    if !IsObject(picker)
        return ""
    try {
        if RegExMatch(picker.Name, "i)currently\s+(.+)$", &m) {
            norm := GeminiNormalizeModelLabel(Trim(m[1]))
            if (norm != "")
                return norm
        }
    } catch {
    }
    return ""
}

FindGeminiModelMenuItem(uia, modelName) {
    if !IsObject(uia)
        return 0
    exp := GeminiNormalizeModelLabel(modelName)
    if (exp = "")
        return 0
    for typeSpec in [50011, "MenuItem", 50008, "RadioButton", 50003, "ListItem"] {
        try {
            el := uia.FindFirst({ Type: typeSpec, Name: exp, mm: 3, cs: 0 })
            if el
                return el
        } catch {
        }
        try {
            el := uia.FindFirst({ Type: typeSpec, Name: exp, mm: 2, cs: 0 })
            if el
                return el
        } catch {
        }
    }
    return 0
}

GeminiWaitForModelMenuItem(uia, modelName, timeoutMs := 0) {
    if (timeoutMs <= 0)
        timeoutMs := GEMINI_MODE_MENU_WAIT_MS
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        el := FindGeminiModelMenuItem(uia, modelName)
        if el
            return el
        Sleep GEMINI_MODE_MENU_POLL_MS
    }
    return 0
}

GeminiWaitForPickerShowsModel(uia, expected, timeoutMs := 0) {
    if (timeoutMs <= 0)
        timeoutMs := GEMINI_MODE_PICKER_LABEL_WAIT_MS
    exp := GeminiNormalizeModelLabel(expected)
    if (exp = "")
        return false
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        picker := FindGeminiModePickerButton(uia)
        if (picker && GeminiGetActiveModelFromPickerElement(picker) = exp)
            return true
        Sleep GEMINI_MODE_MENU_POLL_MS
    }
    return false
}

GeminiNormalizeModelLabel(name) {
    if (name = "")
        return ""
    for canonical in GEMINI_MODEL_CANONICAL_NAMES {
        if (name = canonical || RegExMatch(name, "i)^" . RegExReplace(canonical, "\.", "\.") . "(\s|$)"))
            return canonical
    }
    ; Picker button suffix after "currently" (e.g. Flash-Lite, Flash, Pro)
    if RegExMatch(name, "i)Flash-Lite")
        return "3.1 Flash-Lite"
    if RegExMatch(name, "i)\bFlash\b") && !RegExMatch(name, "i)Flash-Lite")
        return "3.5 Flash"
    if RegExMatch(name, "i)\bPro\b")
        return "3.1 Pro"
    return ""
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

EnsureGeminiModelViaMenu(expected, geminiHwnd := 0) {
    exp := GeminiNormalizeModelLabel(expected)
    if (exp = "")
        return false
    uia := GeminiAttachBrowser(geminiHwnd)
    if !IsObject(uia)
        return false
    browserHwnd := geminiHwnd ? geminiHwnd : GeminiGetBrowserHwndFromUia(uia)
    picker := FindGeminiModePickerButton(uia)
    if !picker
        return false
    if (GeminiGetActiveModelFromPickerElement(picker) = exp)
        return true
    if (!GeminiOpenModePickerMenu(uia, picker, browserHwnd)) {
        GeminiDismissModePickerMenu(uia, browserHwnd)
        return false
    }
    targetBtn := GeminiWaitForModelMenuItem(uia, exp)
    if !targetBtn {
        GeminiDismissModePickerMenu(uia, browserHwnd)
        return false
    }
    if (!GeminiMouseClickElement(targetBtn, browserHwnd, uia, true)) {
        GeminiDismissModePickerMenu(uia, browserHwnd)
        return false
    }
    return GeminiWaitForPickerShowsModel(uia, exp)
}

EnsureGeminiThinkingLevelMenuOpen(geminiHwnd := 0) {
    uia := GeminiAttachBrowser(geminiHwnd)
    if !IsObject(uia)
        return false
    browserHwnd := geminiHwnd ? geminiHwnd : GeminiGetBrowserHwndFromUia(uia)
    picker := FindGeminiModePickerButton(uia)
    if (!picker || !GeminiOpenModePickerMenu(uia, picker, browserHwnd)) {
        GeminiDismissModePickerMenu(uia, browserHwnd)
        return false
    }
    deadline := A_TickCount + GEMINI_MODE_MENU_WAIT_MS
    thinkingItem := 0
    while (A_TickCount < deadline) {
        try {
            thinkingItem := uia.FindFirst({ Name: "Thinking level", mm: 2, cs: 0, Type: 50011 })
        } catch {
            thinkingItem := 0
        }
        if thinkingItem
            break
        Sleep GEMINI_MODE_MENU_POLL_MS
    }
    if !thinkingItem {
        GeminiDismissModePickerMenu(uia, browserHwnd)
        return false
    }
    if (!GeminiMouseClickElement(thinkingItem, browserHwnd, uia, true)) {
        GeminiDismissModePickerMenu(uia, browserHwnd)
        return false
    }
    return true
}
