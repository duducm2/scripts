; =============================================================================
; Shift keys module: hotif_chrome_pdf.ahk
; Chrome PDF viewer hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

; Foreground-hwnd session: one UIA_Browser + viewer root per Chrome window.
; Invalidated when the window changes, an element op fails, F5 runs, or the
; PDF predicate reports this tab is no longer a PDF. TTL stays in the predicate.
global g_ChromePdf_SessHwnd := 0
global g_ChromePdf_SessUia := 0
global g_ChromePdf_SessRoot := 0

ChromePdf_InvalidateSession() {
    global g_ChromePdf_SessHwnd, g_ChromePdf_SessUia, g_ChromePdf_SessRoot
    g_ChromePdf_SessHwnd := 0
    g_ChromePdf_SessUia := 0
    g_ChromePdf_SessRoot := 0
}

ChromePdf_SessionMap(hwnd, uia, root) {
    return Map("hwnd", hwnd, "uia", uia, "root", root)
}

#HotIf IsChromePdfViewerActive()

ChromePdf_GetViewerRoot(uia) {
    ; Prefer the extension's RootWebArea (most stable for the PDF viewer UI)
    if (!uia)
        return 0
    root := 0
    try root := uia.FindElement({ Type: 50030, Value: "chrome-extension://mhjfbmdgcfjbbpaeojofohoefgiehjai",
        matchmode: "Substring" })
    if (root)
        return root

    ; Fallbacks
    try root := uia.GetCurrentDocumentElement()
    if (root)
        return root
    try root := uia.BrowserElement
    return root
}

; Cache-first attach. Same hwnd + live root is reused. A missing root re-finds
; on the existing UIA_Browser (reload poll) instead of constructing a new one.
ChromePdf_GetSession() {
    global g_ChromePdf_SessHwnd, g_ChromePdf_SessUia, g_ChromePdf_SessRoot
    if !WinActive("ahk_exe chrome.exe")
        return 0
    hwnd := WinExist("A")
    if (!hwnd)
        return 0

    if (hwnd = g_ChromePdf_SessHwnd && g_ChromePdf_SessUia) {
        if (g_ChromePdf_SessRoot)
            return ChromePdf_SessionMap(hwnd, g_ChromePdf_SessUia, g_ChromePdf_SessRoot)
        root := 0
        try {
            root := ChromePdf_GetViewerRoot(g_ChromePdf_SessUia)
        } catch {
            g_ChromePdf_SessUia := 0
            g_ChromePdf_SessRoot := 0
        }
        if (g_ChromePdf_SessUia) {
            if (root) {
                g_ChromePdf_SessRoot := root
                return ChromePdf_SessionMap(hwnd, g_ChromePdf_SessUia, root)
            }
            return 0
        }
    }

    uia := 0
    try uia := UIA_Browser("ahk_id " hwnd)
    if (!uia)
        return 0
    root := 0
    try root := ChromePdf_GetViewerRoot(uia)
    g_ChromePdf_SessHwnd := hwnd
    g_ChromePdf_SessUia := uia
    g_ChromePdf_SessRoot := root
    if (!root)
        return 0
    return ChromePdf_SessionMap(hwnd, uia, root)
}

ChromePdf_InvokeElement(el) {
    if (!el)
        return false
    try {
        el.Invoke()
        return true
    } catch {
        try {
            el.Click()
            return true
        } catch {
        }
    }
    return false
}

ChromePdf_FocusElement(el) {
    if (!el)
        return false
    try {
        el.SetFocus()
        return true
    } catch {
        try {
            el.Click()
            return true
        } catch {
        }
    }
    return false
}

ChromePdf_TrySaved(elementId) {
    sess := ChromePdf_GetSession()
    if (!sess)
        return 0
    el := 0
    try el := UiElements_TrySaved(sess["root"], "Chrome PDF Viewer", elementId)
    return el
}

ChromePdf_ClickSaved(elementId) {
    el := ChromePdf_TrySaved(elementId)
    if (!el)
        return false
    if (ChromePdf_InvokeElement(el))
        return true
    ChromePdf_InvalidateSession()
    return false
}

; One typed FindFirst. Name list runs only after that miss.
; No second FindFirst under a different control type when typeHint is already set.
ChromePdf_FindByAutomationId(root, automationId, typeHint := 0, fallbackNames := 0) {
    if (!root || automationId = "")
        return 0
    el := 0
    try {
        if (typeHint)
            el := root.FindFirst({ Type: typeHint, AutomationId: automationId })
        else
            el := root.FindFirst({ AutomationId: automationId })
    } catch {
    }

    if (!el && IsObject(fallbackNames)) {
        nameType := typeHint ? typeHint : 50000
        for , name in fallbackNames {
            try el := root.FindFirst({ Type: nameType, Name: name })
            if (el)
                break
        }
    }
    return el
}

; Find once; on a miss, drop the session and resolve a single time.
ChromePdf_ResolveElement(automationId, typeHint := 0, fallbackNames := 0) {
    loop 2 {
        sess := ChromePdf_GetSession()
        if (!sess) {
            ChromePdf_InvalidateSession()
            continue
        }
        el := ChromePdf_FindByAutomationId(sess["root"], automationId, typeHint, fallbackNames)
        if (el)
            return el
        ChromePdf_InvalidateSession()
    }
    return 0
}

ChromePdf_ClickByAutomationId(automationId, fallbackNames := 0) {
    try {
        btn := ChromePdf_ResolveElement(automationId, 50000, fallbackNames)
        if (!btn)
            return false
        if (ChromePdf_InvokeElement(btn))
            return true
        ChromePdf_InvalidateSession()
    } catch {
        ChromePdf_InvalidateSession()
    }
    return false
}

ChromePdf_FindDownloadButton(root) {
    if (!root)
        return 0
    saved := UiElements_TrySaved(root, "Chrome PDF Viewer", "Download")
    if (saved)
        return saved
    downloadNames := ["Baixar", "Download"]
    btn := 0
    for , name in downloadNames {
        try btn := root.FindFirst({ Type: 50000, Name: name, cs: false })
        if (btn)
            return btn
    }

    ; Two toolbar buttons share AutomationId "save". Never take FindFirst(save) alone.
    try saves := root.FindAll({ Type: 50000, AutomationId: "save" })
    if (IsObject(saves)) {
        for , cand in saves {
            n := ""
            try n := cand.Name
            if (n = "")
                continue
            if InStr(StrLower(n), "drive")
                continue
            return cand
        }
    }
    return 0
}

; PDF toolbar: two buttons share AutomationId "save" (Save to Google Drive vs Download).
ChromePdf_ClickDownload() {
    try {
        loop 2 {
            sess := ChromePdf_GetSession()
            if (!sess) {
                ChromePdf_InvalidateSession()
                continue
            }
            btn := ChromePdf_FindDownloadButton(sess["root"])
            if (btn) {
                if (ChromePdf_InvokeElement(btn))
                    return true
                ChromePdf_InvalidateSession()
                return false
            }
            ChromePdf_InvalidateSession()
        }
    } catch {
    }
    return false
}

ChromePdf_FindPresentMenuItem(root) {
    if (!root)
        return 0
    saved := UiElements_TrySaved(root, "Chrome PDF Viewer", "Present")
    if (saved)
        return saved
    presentItem := 0
    selectorNames := [
        "Present",
        "Presentation mode",
        "Present mode",
        "Apresentar",
        "Modo de apresentação"
    ]

    ; One stable find, then localized names (menu item, then button) only on a miss.
    try presentItem := root.FindFirst({ Type: 50011, AutomationId: "present" })
    if (presentItem)
        return presentItem

    for , candidateName in selectorNames {
        try presentItem := root.FindFirst({ Type: 50011, Name: candidateName })
        if (presentItem)
            return presentItem
        try presentItem := root.FindFirst({ Type: 50000, Name: candidateName })
        if (presentItem)
            return presentItem
    }
    return 0
}

ChromePdf_WaitForMoreMenuReady(timeoutMs := 400) {
    ; Condition wait: More menu populated (any MenuItem under the cached viewer root).
    deadline := A_TickCount + timeoutMs
    while (A_TickCount <= deadline) {
        try {
            sess := ChromePdf_GetSession()
            if (sess) {
                item := 0
                try item := sess["root"].FindFirst({ Type: 50011 })
                if (!item)
                    try item := ChromePdf_FindPresentMenuItem(sess["root"])
                if (item)
                    return true
            }
        } catch {
            ChromePdf_InvalidateSession()
        }
        Sleep 25
    }
    return false
}

ChromePdf_TogglePresentMode() {
    ; Deterministic primary path: open More actions and invoke Present item by selector.
    ; Legacy directional-key fallback remains optional behind feature flag.
    global USE_CHROME_PDF_PRESENT_FALLBACK

    moreBtn := ChromePdf_ResolveElement("more", 50000, ["More actions", "Mais ações"])
    if !ChromePdf_InvokeElement(moreBtn)
        return false

    deadline := A_TickCount + 700
    while (A_TickCount <= deadline) {
        try {
            sess := ChromePdf_GetSession()
            if (sess) {
                presentItem := ChromePdf_FindPresentMenuItem(sess["root"])
                if (presentItem && ChromePdf_InvokeElement(presentItem))
                    return true
            }
        } catch {
            ChromePdf_InvalidateSession()
        }
        Sleep 40
    }

    if (USE_CHROME_PDF_PRESENT_FALLBACK) {
        Sleep 120
        Send "{Up}"
        Send "{Up}"
        Sleep 40
        Send "{Enter}"
        return true
    }

    return false
}

ChromePdf_PollPageSelector() {
    ; Reload poll: reuse the session UIA. A missing control clears only the root
    ; so the next tick re-finds the viewer without a new UIA_Browser.
    global g_ChromePdf_SessRoot
    sess := ChromePdf_GetSession()
    if (!sess)
        return 0
    el := 0
    try {
        el := sess["root"].FindFirst({ Type: 50004, AutomationId: "pageSelector" })
    } catch {
        ChromePdf_InvalidateSession()
        return 0
    }
    if (!el)
        g_ChromePdf_SessRoot := 0
    return el
}

ChromePdf_PageNumberFromElement(el) {
    if (!el)
        return ""
    val := ""
    try val := Trim(String(el.Value))
    if (val = "")
        return ""
    ; pageSelector is usually bare digits; tolerate "12 / 40" style values.
    if RegExMatch(val, "(\d+)", &m)
        return m[1]
    return ""
}

ChromePdf_GetCurrentPageNumber() {
    el := ChromePdf_ResolveElement("pageSelector", 50004)
    return ChromePdf_PageNumberFromElement(el)
}

ChromePdf_GotoPage(pageNum, el := 0) {
    pageNum := Trim(String(pageNum))
    if (pageNum = "" || !RegExMatch(pageNum, "^\d+$"))
        return false
    if (!el)
        el := ChromePdf_ResolveElement("pageSelector", 50004)
    if !ChromePdf_FocusElement(el) {
        ChromePdf_InvalidateSession()
        return false
    }
    Send "^a"
    Sleep 20
    SendText pageNum
    Sleep 20
    Send "{Enter}"
    return true
}

ChromePdf_WaitForReloadSettled(savedPage, timeoutMs := 10000) {
    ; Prefer: pageSelector vanishes then returns. Also accept Value reset to 1 when we
    ; were elsewhere (fast reload where the gap was never sampled).
    ; Returns the settled page-selector element, or 0.
    global g_ChromePdf_SessRoot
    deadline := A_TickCount + timeoutMs
    start := A_TickCount
    sawGone := false
    lastEl := 0
    while (A_TickCount <= deadline) {
        if !WinActive("ahk_exe chrome.exe")
            return 0
        el := ChromePdf_PollPageSelector()
        if (!el) {
            sawGone := true
            lastEl := 0
            Sleep 40
            continue
        }
        lastEl := el
        val := ""
        try {
            val := Trim(String(el.Value))
        } catch {
            ; Stale node after reload: drop the root and treat the viewer as gone.
            g_ChromePdf_SessRoot := 0
            sawGone := true
            lastEl := 0
            Sleep 40
            continue
        }
        if (val = "") {
            Sleep 40
            continue
        }
        cur := ""
        if RegExMatch(val, "(\d+)", &m)
            cur := m[1]
        if (sawGone)
            return el
        if (savedPage != "" && savedPage != "1" && cur = "1")
            return el
        ; Already on page 1: no Value flip to observe — settle briefly then continue.
        if (savedPage = "1" && cur = "1" && (A_TickCount - start) >= 500)
            return el
        Sleep 40
    }
    return lastEl
}

ChromePdf_RefreshKeepPage() {
    ; F5 reloads the PDF (often resets to page 1); restore the page we were on.
    page := ChromePdf_GetCurrentPageNumber()
    if (page = "") {
        ShowCenteredOverlay_Utils("❌ PDF: could not read current page", 2000, BANNER_ACCENT_ERROR)
        return false
    }

    try StandardLoadingBar_Show("🔄 Refreshing PDF (page " page ")…", BANNER_ACCENT_INTERMEDIATE, {
        passive: false })

    Send "{F5}"
    ChromePdf_InvalidateSession()
    ChromePdf_InvalidatePredicateCache()  ; title is unchanged across reload; force a re-probe

    try StandardLoadingBar_Update("⏳ Waiting for PDF viewer…", BANNER_ACCENT_INTERMEDIATE)
    el := ChromePdf_WaitForReloadSettled(page, 10000)
    if (!el) {
        try StandardLoadingBar_Update("❌ PDF: viewer did not return", BANNER_ACCENT_ERROR)
        try StandardLoadingBar_Hide(2000)
        return false
    }

    if !ChromePdf_GotoPage(page, el) {
        try StandardLoadingBar_Update("❌ PDF: refreshed, but could not go to page " page, BANNER_ACCENT_ERROR)
        try StandardLoadingBar_Hide(2500)
        return false
    }

    try StandardLoadingBar_Update("✅ PDF refreshed → page " page, BANNER_ACCENT_SUCCESS)
    try StandardLoadingBar_Hide(1400)
    return true
}

; Shift + F : Fit to page (Zoom to Fit) - Fit
+f::
{
    if ChromePdf_ClickSaved("Fit")
        return
    ; UIA tree: AutomationId "fit"
    ChromePdf_ClickByAutomationId("fit")
}

; Shift + P : Focus page number field and select contents - Page
+p::
{
    ; UIA tree: Edit AutomationId "pageSelector"
    ; SetFocus is the readiness gate; Ctrl+A selects so typing replaces the number.
    el := ChromePdf_TrySaved("Page")
    if (!el)
        el := ChromePdf_ResolveElement("pageSelector", 50004)
    if (!el)
        return
    if !ChromePdf_FocusElement(el) {
        ChromePdf_InvalidateSession()
        return
    }
    Send "^a"
}

; Shift + T : Toggle thumbnails sidebar - Thumbnails
+t::
{
    if ChromePdf_ClickSaved("Thumbnails")
        return
    ; UIA tree: AutomationId "sidenavToggle"
    ChromePdf_ClickByAutomationId("sidenavToggle")
}

; Shift + D : Download PDF - Download
+d::
{
    ; Toolbar duplicates AutomationId "save" (Drive save vs file download); use ChromePdf_ClickDownload.
    ChromePdf_ClickDownload()
}

; Shift + 2 : Two-page view (mnemonic: 2 = two pages)
+2::
{
    ; UIA: Button Type 50000, Name "More actions", AutomationId "more"
    ; Keep Down/Enter (no stable AutomationId for two-page in tree dump); wait for menu ready.
    moreBtn := ChromePdf_TrySaved("MoreActions")
    if !moreBtn
        moreBtn := ChromePdf_ResolveElement("more", 50000, ["More actions", "Mais ações"])
    if !ChromePdf_InvokeElement(moreBtn)
        return
    ChromePdf_WaitForMoreMenuReady(400)
    if ChromePdf_ClickSaved("TwoPage")
        return
    Send "{Down}"
    Send "{Enter}"
}

; Shift + E : Present mode (mnemonic: E from prEsent)
+E::
{
    ChromePdf_TogglePresentMode()
}

; Shift + R : Refresh PDF and restore current page - Refresh
+r::
{
    ChromePdf_RefreshKeepPage()
}

#HotIf