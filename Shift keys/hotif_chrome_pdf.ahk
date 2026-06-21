; =============================================================================
; Shift keys module: hotif_chrome_pdf.ahk
; Chrome PDF viewer hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

#HotIf IsChromePdfViewerActive()

ChromePdf_GetViewerRoot(uia) {
    ; Prefer the extension's RootWebArea (most stable for the PDF viewer UI)
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

ChromePdf_ClickByAutomationId(automationId, fallbackNames := 0) {
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 80

        root := ChromePdf_GetViewerRoot(uia)
        if (!root)
            return false

        btn := 0
        try btn := root.FindFirst({ Type: 50000, AutomationId: automationId })
        if (!btn)
            try btn := root.FindFirst({ AutomationId: automationId })

        if (!btn && IsObject(fallbackNames)) {
            for , name in fallbackNames {
                try btn := root.FindFirst({ Type: 50000, Name: name })
                if (btn)
                    break
            }
        }

        if (btn) {
            try btn.Invoke()
            catch {
                try btn.Click()
            }
            return true
        }
    } catch {
    }
    return false
}

; PDF toolbar: two buttons share AutomationId "save" (Save to Google Drive vs Download). Never use FindFirst(save) alone.
ChromePdf_ClickDownload() {
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 80

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

        if (btn) {
            try btn.Invoke()
            catch {
                try btn.Click()
            }
            return true
        }
    } catch {
    }
    return false
}

ChromePdf_FocusByAutomationId(automationId, controlType := 0) {
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 80

        root := ChromePdf_GetViewerRoot(uia)
        if (!root)
            return false

        el := 0
        if (controlType) {
            try el := root.FindFirst({ Type: controlType, AutomationId: automationId })
        }
        if (!el)
            try el := root.FindFirst({ AutomationId: automationId })

        if (el) {
            try el.SetFocus()
            catch {
                try el.Click()
            }
            return true
        }
    } catch {
    }
    return false
}

ChromePdf_TogglePresentMode() {
    ; Deterministic primary path: open More actions and invoke Present item by selector.
    ; Legacy directional-key fallback remains optional behind feature flag.
    global USE_CHROME_PDF_PRESENT_FALLBACK

    if !ChromePdf_ClickByAutomationId("more", ["More actions"])
        return false

    deadline := A_TickCount + 700
    selectorNames := [
        "Present",
        "Presentation mode",
        "Present mode",
        "Apresentar",
        "Modo de apresentação"
    ]

    while (A_TickCount <= deadline) {
        try {
            uia := UIA_Browser("ahk_exe chrome.exe")
            root := ChromePdf_GetViewerRoot(uia)
            if (root) {
                presentItem := 0

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

                if (presentItem) {
                    try presentItem.Invoke()
                    catch {
                        try presentItem.Click()
                    }
                    return true
                }
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

; Shift + F : Fit to page (Zoom to Fit) - Fit
+f::
{
    ; UIA tree: AutomationId "fit"
    ChromePdf_ClickByAutomationId("fit")
}

; Shift + P : Focus page number field - Page
+p::
{
    ; UIA tree: Edit AutomationId "pageSelector"
    ; Per requirement: focus only (no select-all)
    ChromePdf_FocusByAutomationId("pageSelector", 50004)
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
    if ChromePdf_ClickByAutomationId("more", ["More actions"]) {
        Sleep 150
        Send "{Down}"
        Sleep 50
        Send "{Enter}"
    }
}

; Shift + E : Present mode (mnemonic: E from prEsent)
+E::
{
    ChromePdf_TogglePresentMode()
}

#HotIf