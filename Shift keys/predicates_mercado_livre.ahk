; =============================================================================
; Shift keys module: predicates_mercado_livre.ahk
; IsMercadoLivreActive and ML UIA helpers
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

;-------------------------------------------------------------------
; Mercado Livre (Brazil) Shortcuts
;-------------------------------------------------------------------
; Cache for IsMercadoLivreActive (per efficiency-canon: cache-first with validation).
; Invalidated when foreground HWND changes so we only run UIA once per window/tab focus.
global g_ML_CacheHwnd := 0
global g_ML_CacheResult := false
; Cache for initial-page-load workaround: right-click + close context menu to make hotkeys work (once per window).
global g_ML_ReceptivityHwnd := 0

; Workaround for ML: hotkeys fail on initial page load until page is "initialized". Right-click on empty area then close context menu.
; force=true: run even if cache says we already did (error-driven retry when hotkey action failed).
ML_EnsureHotkeyReceptivity(force := false) {
    global g_ML_ReceptivityHwnd
    if !WinActive("ahk_exe chrome.exe")
        return
    hwnd := WinExist("A")
    if (!hwnd)
        return
    if (!force && hwnd = g_ML_ReceptivityHwnd && WinExist("ahk_id " g_ML_ReceptivityHwnd))
        return
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        try
            root := uia.GetCurrentDocumentElement()
        catch
            root := uia.BrowserElement
        if (!root)
            return
        br := root.BoundingRectangle
        if (!br || (br.r <= br.l) || (br.b <= br.t))
            return
        x := br.l + (br.r - br.l) * 0.15
        y := br.t + (br.b - br.t) * 0.20
        prevMode := A_CoordModeMouse
        CoordMode("Mouse", "Screen")
        MouseClick("Right", x, y)
        SendEscape()
        CoordMode("Mouse", prevMode)
        g_ML_ReceptivityHwnd := hwnd
    } catch {
    }
}

IsMercadoLivreActive() {
    global g_ML_CacheHwnd, g_ML_CacheResult
    if !WinActive("ahk_exe chrome.exe")
        return false
    hwnd := WinExist("A")
    if (!hwnd)
        return false
    ; Cache hit: same window as last check (avoids UIA on every keystroke / cheat sheet open)
    if (hwnd = g_ML_CacheHwnd && WinExist("ahk_id " g_ML_CacheHwnd))
        return g_ML_CacheResult
    ; Platform identification by URL only (do not use window title; it changes to product name). See shopping uia3.md.
    ; URL check via UIA (address bar: Chrome exposes AcceleratorKey "Ctrl+L", not AccessKey). Bounded to this window only.
    try {
        root := UIA.ElementFromHandle(hwnd)
        addressBar := root.FindFirst({ Type: 50004, AcceleratorKey: "Ctrl+L" })
        if (addressBar) {
            url := addressBar.Value
            if InStr(url, "mercadolivre.com") || InStr(url, "mercadolibre.com") {
                g_ML_CacheHwnd := hwnd
                g_ML_CacheResult := true
                return true
            }
        }
    } catch {
        ; UIA failed; do not cache so next call retries
    }
    g_ML_CacheHwnd := hwnd
    g_ML_CacheResult := false
    return false
}

; Mercado Livre UIA helpers: get document root and find/invoke elements (bounded, no global state).
; When UIA_Browser init fails (document not ready), fallback: get Document from foreground window tree (Type 50030 = Document).
ML_GetDocRoot() {
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        try
            return uia.GetCurrentDocumentElement()
        catch
            return uia.BrowserElement
    } catch {
        ; Fallback: document not ready for UIA_Browser init (e.g. on first load). Get Document from active window.
        try {
            hwnd := WinExist("A")
            if (hwnd && WinActive("ahk_exe chrome.exe")) {
                root := UIA.ElementFromHandle(hwnd)
                doc := root.FindFirst({ Type: 50030 })
                if (doc)
                    return doc
            }
        } catch {
        }
        return 0
    }
}

; #region agent log - debug helper
ML_DebugLog(hypothesisId, message, detail, runId := "pre-fix") {
    ; Intentionally no-op.
    return
}
; #endregion agent log

; Locate the Mercado Livre Preço filter container, scoped to the block that
; contains the predefined price links (Até R$ 40 / R$ 40 a R$ 95 / Mais de R$ 95)
; and the Mínimo/Máximo edits. Returns the container element or 0.
ML_GetPriceFilter(root) {
    if !root
        return 0
    priceLinkNames := ["Até R$ 40", "R$ 40 a R$ 95", "Mais de R$ 95"]
    for _, name in priceLinkNames {
        link := ML_Find(root, { Type: 50005, Name: name, cs: false })
        if !link
            continue
        parent := link
        ; Walk up a few levels to find a container that has the Mínimo edit as a descendant.
        loop 6 {
            try parent := parent.Parent
            catch
                break
            if !parent
                break
            minimo := ML_Find(parent, { Type: 50004, Name: "Mínimo", cs: false })
            if (minimo) {
                return parent
            }
        }
    }
    return 0
}

; Best-effort helper to set text of an edit element. Tries ValuePattern first,
; then falls back to focus + clear + send keys.
ML_SetEditText(el, text) {
    if (!el)
        return false
    if (text = "")
        return false
    ok := false
    try {
        if el.GetPropertyValue(UIA.Property.IsValuePatternAvailable) {
            vp := el.ValuePattern
            vp.SetValue(text)
            ok := true
        }
    } catch {
        ML_DebugLog("C", "ValuePattern attempt failed", "exception in SetValue", "run1")
    }
    if (!ok) {
        try el.ScrollIntoView()
        catch {
        }
        try {
            el.SetFocus()
        } catch {
            try el.Click()
            catch {
            }
        }
        Sleep 50
        ML_DebugLog("C", "Focus acquired for edit", text, "run1")
        Send "^a{Del}"
        Sleep 30
        SendText(text)
        ok := true
    }
    return ok
}

; Try conditions in order; invoke or click first match. Returns true if invoked/clicked, false otherwise.
ML_FindAndInvoke(conditionList) {
    root := ML_GetDocRoot()
    if (!root)
        return false
    for cond in conditionList {
        try {
            el := root.FindElement(cond, UIA.TreeScope.Descendants)
            if (el) {
                try el.Invoke()
                catch {
                    try el.Click()
                    catch
                        return false
                }
                return true
            }
        } catch
            continue
    }
    return false
}

; Find single element by condition (Descendants). Returns element or 0.
ML_Find(root, condition) {
    try
        return root.FindElement(condition, UIA.TreeScope.Descendants)
    catch
        return 0
}
