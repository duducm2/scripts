; =============================================================================
; Shift keys module: hotif_chrome_pdf.ahk
; Chrome PDF viewer hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

#HotIf IsChromePdfViewerActive()

ChromePdf_GetActiveUia() {
    ; Bind to the specific foreground Chrome window (efficiency-canon §11).
    if !WinActive("ahk_exe chrome.exe")
        return 0
    hwnd := WinExist("A")
    if (!hwnd)
        return 0
    try {
        return UIA_Browser("ahk_id " hwnd)
    } catch {
    }
    return 0
}

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

ChromePdf_FindByAutomationId(root, automationId, typeHint := 0, fallbackNames := 0) {
    if (!root || automationId = "")
        return 0
    el := 0
    if (typeHint) {
        try el := root.FindFirst({ Type: typeHint, AutomationId: automationId })
    }
    if (!el)
        try el := root.FindFirst({ Type: 50000, AutomationId: automationId })
    if (!el)
        try el := root.FindFirst({ AutomationId: automationId })

    if (!el && IsObject(fallbackNames)) {
        for , name in fallbackNames {
            try el := root.FindFirst({ Type: 50000, Name: name })
            if (el)
                break
        }
    }
    return el
}

ChromePdf_ClickByAutomationId(automationId, fallbackNames := 0) {
    try {
        uia := ChromePdf_GetActiveUia()
        if (!uia)
            return false

        root := ChromePdf_GetViewerRoot(uia)
        if (!root)
            return false

        btn := ChromePdf_FindByAutomationId(root, automationId, 50000, fallbackNames)
        if (btn)
            return ChromePdf_InvokeElement(btn)
    } catch {
    }
    return false
}

; PDF toolbar: two buttons share AutomationId "save" (Save to Google Drive vs Download). Never use FindFirst(save) alone.
ChromePdf_ClickDownload() {
    try {
        uia := ChromePdf_GetActiveUia()
        if (!uia)
            return false

        root := ChromePdf_GetViewerRoot(uia)
        if (!root)
            return false

        downloadNames := ["Baixar", "Download"]
        btn := 0
        for , name in downloadNames {
            try btn := root.FindFirst({ Type: 50000, Name: name, cs: false })
            if (btn)
                break
        }

        if (!btn) {
            try saves := root.FindAll({ Type: 50000, AutomationId: "save" })
            if (IsObject(saves)) {
                for , cand in saves {
                    n := ""
                    try n := cand.Name
                    if (n = "")
                        continue
                    if InStr(StrLower(n), "drive")
                        continue
                    btn := cand
                    break
                }
            }
        }

        if (btn)
            return ChromePdf_InvokeElement(btn)
    } catch {
    }
    return false
}

ChromePdf_FocusByAutomationId(automationId, controlType := 0) {
    try {
        uia := ChromePdf_GetActiveUia()
        if (!uia)
            return false

        root := ChromePdf_GetViewerRoot(uia)
        if (!root)
            return false

        el := ChromePdf_FindByAutomationId(root, automationId, controlType)
        if (el) {
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
        }
    } catch {
    }
    return false
}

ChromePdf_FindPresentMenuItem(root) {
    if (!root)
        return 0
    presentItem := 0
    selectorNames := [
        "Present",
        "Presentation mode",
        "Present mode",
        "Apresentar",
        "Modo de apresentação"
    ]

    ; Prefer stable attributes first, then localized names.
    try presentItem := root.FindFirst({ Type: 50011, AutomationId: "present" })
    if (!presentItem)
        try presentItem := root.FindFirst({ AutomationId: "present" })
    if (!presentItem)
        try presentItem := root.FindFirst({ Type: 50000, AutomationId: "present" })

    if (!presentItem) {
        for , candidateName in selectorNames {
            try presentItem := root.FindFirst({ Type: 50011, Name: candidateName })
            if (presentItem)
                break
            try presentItem := root.FindFirst({ Type: 50000, Name: candidateName })
            if (presentItem)
                break
        }
    }
    return presentItem
}

ChromePdf_WaitForMoreMenuReady(uia, timeoutMs := 400) {
    ; Condition wait: More menu populated (any MenuItem under the PDF viewer root).
    if (!uia)
        return false
    deadline := A_TickCount + timeoutMs
    while (A_TickCount <= deadline) {
        try {
            root := ChromePdf_GetViewerRoot(uia)
            if (root) {
                item := 0
                try item := root.FindFirst({ Type: 50011 })
                if (!item)
                    try item := ChromePdf_FindPresentMenuItem(root)
                if (item)
                    return true
            }
        } catch {
        }
        Sleep 25
    }
    return false
}

ChromePdf_TogglePresentMode() {
    ; Deterministic primary path: open More actions and invoke Present item by selector.
    ; Legacy directional-key fallback remains optional behind feature flag.
    global USE_CHROME_PDF_PRESENT_FALLBACK

    uia := ChromePdf_GetActiveUia()
    if (!uia)
        return false

    root := ChromePdf_GetViewerRoot(uia)
    if (!root)
        return false

    moreBtn := ChromePdf_FindByAutomationId(root, "more", 50000, ["More actions", "Mais ações"])
    if !ChromePdf_InvokeElement(moreBtn)
        return false

    deadline := A_TickCount + 700
    while (A_TickCount <= deadline) {
        try {
            root := ChromePdf_GetViewerRoot(uia)
            presentItem := ChromePdf_FindPresentMenuItem(root)
            if (presentItem) {
                if ChromePdf_InvokeElement(presentItem)
                    return true
            }
        } catch {
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

ChromePdf_GetPageSelectorElement() {
    try {
        uia := ChromePdf_GetActiveUia()
        if (!uia)
            return 0
        root := ChromePdf_GetViewerRoot(uia)
        if (!root)
            return 0
        return ChromePdf_FindByAutomationId(root, "pageSelector", 50004)
    } catch {
    }
    return 0
}

ChromePdf_GetCurrentPageNumber() {
    el := ChromePdf_GetPageSelectorElement()
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

ChromePdf_GotoPage(pageNum) {
    pageNum := Trim(String(pageNum))
    if (pageNum = "" || !RegExMatch(pageNum, "^\d+$"))
        return false
    if !ChromePdf_FocusByAutomationId("pageSelector", 50004)
        return false
    Sleep 40
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
    deadline := A_TickCount + timeoutMs
    start := A_TickCount
    sawGone := false
    while (A_TickCount <= deadline) {
        if !WinActive("ahk_exe chrome.exe")
            return false
        el := ChromePdf_GetPageSelectorElement()
        if (!el) {
            sawGone := true
            Sleep 40
            continue
        }
        val := ""
        try val := Trim(String(el.Value))
        if (val = "") {
            Sleep 40
            continue
        }
        cur := ""
        if RegExMatch(val, "(\d+)", &m)
            cur := m[1]
        if (sawGone)
            return true
        if (savedPage != "" && savedPage != "1" && cur = "1")
            return true
        ; Already on page 1: no Value flip to observe — settle briefly then continue.
        if (savedPage = "1" && cur = "1" && (A_TickCount - start) >= 500)
            return true
        Sleep 40
    }
    return !!ChromePdf_GetPageSelectorElement()
}

ChromePdf_RefreshKeepPage() {
    ; F5 reloads the PDF (often resets to page 1); restore the page we were on.
    global g_ChromePdf_CacheTick

    page := ChromePdf_GetCurrentPageNumber()
    if (page = "") {
        ShowCenteredOverlay_Utils("❌ PDF: could not read current page", 2000, BANNER_ACCENT_ERROR)
        return false
    }

    try StandardLoadingBar_Show("🔄 Refreshing PDF (page " page ")…", BANNER_ACCENT_INTERMEDIATE, {
        passive: false })

    Send "{F5}"
    g_ChromePdf_CacheTick := 0  ; force IsChromePdfViewerActive to re-probe after reload

    try StandardLoadingBar_Update("⏳ Waiting for PDF viewer…", BANNER_ACCENT_INTERMEDIATE)
    if !ChromePdf_WaitForReloadSettled(page, 10000) {
        try StandardLoadingBar_Update("❌ PDF: viewer did not return", BANNER_ACCENT_ERROR)
        try StandardLoadingBar_Hide(2000)
        return false
    }

    Sleep 80
    if !ChromePdf_GotoPage(page) {
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
    ; UIA tree: AutomationId "fit"
    ChromePdf_ClickByAutomationId("fit")
}

; Shift + P : Focus page number field and select contents - Page
+p::
{
    ; UIA tree: Edit AutomationId "pageSelector"
    ; Focus then Ctrl+A so typing replaces the current page number.
    if ChromePdf_FocusByAutomationId("pageSelector", 50004) {
        Sleep 40
        Send "^a"
    }
}

; Shift + T : Toggle thumbnails sidebar - Thumbnails
+t::
{
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
    uia := ChromePdf_GetActiveUia()
    if (!uia)
        return
    root := ChromePdf_GetViewerRoot(uia)
    if (!root)
        return
    moreBtn := ChromePdf_FindByAutomationId(root, "more", 50000, ["More actions", "Mais ações"])
    if !ChromePdf_InvokeElement(moreBtn)
        return
    ChromePdf_WaitForMoreMenuReady(uia, 400)
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