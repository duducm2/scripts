; =============================================================================
; Utils module: handy_uia_helpers.ahk
; Handy UIA helper functions and ShowAiModelSelector support
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; Handy UIA Helper Functions
; =============================================================================

; Activate existing Handy window or launch it; returns hwnd or 0
Handy_ActivateOrLaunch() {
    targetPath := GetHandyShortcutPath()
    expectedExePath := GetHandyProcessPath()

    ; Find existing Handy window
    matchingHwnd := 0
    for hwnd in WinGetList("Handy ahk_class Tauri Window") {
        try {
            procPath := WinGetProcessPath(hwnd)
            if (expectedExePath = "" || StrCompare(procPath, expectedExePath, false) = 0) {
                matchingHwnd := hwnd
                break
            }
        } catch {
            if (expectedExePath = "") {
                matchingHwnd := hwnd
                break
            }
        }
    }

    if (matchingHwnd) {
        WinActivate("ahk_id " . matchingHwnd)
        WinWaitActive("ahk_id " . matchingHwnd, , 2)
        Handy_WaitForMainUiReady(matchingHwnd, 2000)
        return matchingHwnd
    }

    ; Launch Handy
    if (targetPath = "" || !FileExist(targetPath))
        return 0

    Run targetPath
    if !WinWait("Handy ahk_class Tauri Window", , 8)
        return 0

    ; Find the window we just launched
    for h in WinGetList("Handy ahk_class Tauri Window") {
        try {
            procPath := WinGetProcessPath(h)
            if (expectedExePath = "" || StrCompare(procPath, expectedExePath, false) = 0) {
                WinActivate("ahk_id " . h)
                WinWaitActive("ahk_id " . h, , 2)
                if (!Handy_WaitForMainUiReady(h, 9000))
                    return 0
                return h
            }
        } catch {
            if (expectedExePath = "") {
                WinActivate("ahk_id " . h)
                WinWaitActive("ahk_id " . h, , 2)
                if (!Handy_WaitForMainUiReady(h, 9000))
                    return 0
                return h
            }
        }
    }
    return 0
}

; Wait for Handy main UI to be interactive (needed most on cold launch).
Handy_WaitForMainUiReady(hwnd, maxWaitMs := 9000) {
    global UIA
    start := A_TickCount
    pollMs := 120
    loop {
        if ((A_TickCount - start) >= maxWaitMs)
            return false
        el := UIA.ElementFromHandle(hwnd)
        if (el) {
            try {
                if (el.FindFirst({ Type: 50000, Name: "Check for updates" }))
                    return true
            }
            try {
                if (el.FindFirst({ Type: 50000, Name: "Verificar atualizações" }))
                    return true
            }
            try {
                if (el.FindFirst({ Type: 50000, Name: "Update available" }))
                    return true
            }
        }
        Sleep pollMs
    }
}

; True when General tab content (COHERE SETTINGS) is visible.
Handy_GeneralTabVisible(el) {
    if !el
        return false
    try {
        return el.FindFirst({ Type: 50020, Name: "COHERE SETTINGS" }) != 0
    } catch {
        return false
    }
}

; Click sidebar "General" so COHERE SETTINGS is shown (needed from Models/About/etc.).
Handy_EnsureGeneralTab(hwnd) {
    el := UIA.ElementFromHandle(hwnd)
    if !el
        return false
    if (Handy_GeneralTabVisible(el))
        return true
    try {
        gen := el.FindFirst({ Type: 50020, Name: "General" })
        if gen {
            try gen.Click()
            catch {
                try gen.Invoke()
            }
            Sleep 220
            el2 := UIA.ElementFromHandle(hwnd)
            return Handy_GeneralTabVisible(el2)
        }
    } catch {
    }
    return false
}

; Language dropdown under COHERE SETTINGS: class uses "rounded min-w-[200px]" (Microphone uses rounded-md).
Handy_FindHandyLanguageButton(el) {
    if !el
        return 0
    try {
        buttons := el.FindAll({ Type: 50000 })
        for btn in buttons {
            cn := ""
            try cn := btn.ClassName
            if (cn != "" && InStr(cn, "rounded min-w-[200px]"))
                return btn
        }
    } catch {
    }
    return 0
}

; Current Cohere language label on General tab ("" if unknown).
Handy_ReadCohereLanguage(hwnd) {
    el := UIA.ElementFromHandle(hwnd)
    if !el
        return ""
    langBtn := Handy_FindHandyLanguageButton(el)
    if !langBtn
        return ""
    try return langBtn.Name
    return ""
}

; Poll until language button shows langName (short window; no-op if already correct).
Handy_WaitCohereLanguage(hwnd, langName, maxWaitMs := 450) {
    if (langName = "")
        return false
    pollMs := 50
    start := A_TickCount
    loop {
        if (Handy_ReadCohereLanguage(hwnd) = langName)
            return true
        if ((A_TickCount - start) >= maxWaitMs)
            break
        Sleep pollMs
    }
    return Handy_ReadCohereLanguage(hwnd) = langName
}

; Open the COHERE language dropdown on General tab.
Handy_OpenCohereLanguageDropdown(hwnd) {
    el := UIA.ElementFromHandle(hwnd)
    if !el
        return false
    langBtn := Handy_FindHandyLanguageButton(el)
    if !langBtn
        return false
    try langBtn.Click()
    catch {
        try langBtn.Invoke()
    }
    Sleep 200
    return true
}

; With language dropdown open: focus search, type langName, choose row or Enter.
Handy_SetCohereLanguage_PickFromOpenDropdown(hwnd, langName) {
    el := UIA.ElementFromHandle(hwnd)
    if !el
        return false
    searchEl := 0
    try {
        for ed in el.FindAll({ Type: UIA.Type.Edit }) {
            searchEl := ed
            break
        }
    } catch {
    }
    if (searchEl) {
        try {
            searchEl.SetFocus()
        } catch {
            try searchEl.Click()
        }
        Sleep 50
    }
    Send "^a"
    SendText langName
    Sleep 120
    picked := false
    try {
        for btn in el.FindAll({ Type: 50000 }) {
            n := ""
            try n := btn.Name
            if (n != langName)
                continue
            cn := ""
            try cn := btn.ClassName
            if (InStr(cn, "w-full px-3 py-2 text-left") || InStr(cn, "w-full px-3 py-2 text-start")) {
                try btn.Click()
                picked := true
                break
            }
        }
    } catch {
    }
    if !picked
        Send "{Enter}"
    Sleep 80
    return true
}

; Set Cohere transcription language on General tab (explicit list pick, not Auto Detect).
; Retries with verify-after-pick until correct or max attempts (slots 3–4 / English & Portuguese).
Handy_SetCohereLanguage(hwnd, langName) {
    if !hwnd || langName = ""
        return false
    if !Handy_EnsureGeneralTab(hwnd)
        return false
    if (Handy_ReadCohereLanguage(hwnd) = langName)
        return true

    maxAttempts := 3
    loop maxAttempts {
        if (A_Index > 1) {
            Send "{Escape}"
            Sleep 80
            if !Handy_EnsureGeneralTab(hwnd)
                continue
        }
        if !Handy_OpenCohereLanguageDropdown(hwnd)
            continue
        Handy_SetCohereLanguage_PickFromOpenDropdown(hwnd, langName)
        if (Handy_WaitCohereLanguage(hwnd, langName))
            return true
        Send "{Escape}"
        Sleep 80
    }
    return false
}

; Open the AI model dropdown menu using keyboard navigation
Handy_OpenAiModelMenu(hwnd) {
    el := UIA.ElementFromHandle(hwnd)
    if !el {
        return false
    }

    ; Find anchor: primary "Check for updates" button
    anchor := 0
    try anchor := el.FindFirst({
        Type: 50000,
        ClassName: "transition-colors disabled:opacity-50 tabular-nums text-text/60 hover:text-text/80"
    })
    if (!anchor) {
        try anchor := el.FindFirst({ Type: 50000, Name: "Check for updates" })
    }
    if (!anchor) {
        try anchor := el.FindFirst({ Type: 50000, Name: "Verificar atualizações" })
    }

    ; Fallback: "Update available" anchor when a system update banner is shown
    if (!anchor) {
        try anchor := el.FindFirst({
            Type: 50000,
            ClassName: "transition-colors disabled:opacity-50 tabular-nums text-logo-primary hover:text-logo-primary/80 font-medium"
        })
    }
    if (!anchor) {
        try anchor := el.FindFirst({ Type: 50000, Name: "Update available" })
    }
    if (!anchor) {
        ; Last-resort: use technical condition path to reach the "Update available" button
        try anchor := el.ElementFromPath({ T: 33 }, { T: 33 }, { T: 33 }, { T: 33, CN: "BrowserRootView" }, { T: 33 }, { T: 33,
            CN: "EmbeddedBrowserFrameView" }, { T: 33, CN: "BrowserView" }, { T: 33, CN: "SidebarContentsSplitView" }, { T: 33 }, { T: 33 }, { T: 33 }, { T: 30 }, { T: 26 }, { T: 0,
                CN: "transition-colors disabled:opacity-50 tabular-nums text-logo-primary hover:text-logo-primary/80 font-medium" }
        )
    }

    if (!anchor) {
        return false
    }

    ; Focus anchor, Shift+Tab to model button, Enter to open menu
    try anchor.SetFocus()
    catch {
        try anchor.Click()
    }
    Sleep 60
    Send "+{Tab}"
    Sleep 60
    Send "{Enter}"

    ; Context menu can open slowly in Handy; wait for menu row(s) to actually exist.
    return Handy_WaitForAiModelMenuOpen(hwnd, 2500)
}

; Wait for AI model context menu rows to appear after opening the menu.
Handy_WaitForAiModelMenuOpen(hwnd, maxWaitMs := 2500) {
    global UIA
    start := A_TickCount
    pollMs := 100
    loop {
        if ((A_TickCount - start) >= maxWaitMs)
            return false
        el := UIA.ElementFromHandle(hwnd)
        if (el) {
            try {
                for btn in el.FindAll({ Type: 50000 }) {
                    cn := ""
                    try cn := btn.ClassName
                    if (InStr(cn, "w-full px-3 py-2 text-left") || InStr(cn, "w-full px-3 py-2 text-start"))
                        return true
                }
            }
        }
        Sleep pollMs
    }
}

; Find and click the AI model button by partial name match
Handy_ClickAiModel(hwnd, modelName) {
    el := UIA.ElementFromHandle(hwnd)
    if !el {
        return false
    }

    ; Model buttons have class containing "w-full px-3 py-2 text-left"
    ; and names starting with the model name (e.g., "Whisper Large Good accuracy...")
    ; Try to find by partial name match
    modelBtn := 0
    buttonCount := 0
    nameMatchNoClass := ""

    ; Strategy 1: Find button whose Name starts with modelName
    try {
        buttons := el.FindAll({ Type: 50000 })
        for btn in buttons {
            buttonCount++
            btnName := ""
            try btnName := btn.Name
            if (btnName != "" && InStr(btnName, modelName) = 1) {
                btnClass := ""
                try btnClass := btn.ClassName
                ; Menu items: w-full px-3 py-2 text-left (legacy) or text-start (new Handy UI); header: flex items-center gap-2
                if (InStr(btnClass, "w-full px-3 py-2 text-left") || InStr(btnClass, "w-full px-3 py-2 text-start") ||
                InStr(btnClass, "flex items-center gap-2")) {
                    modelBtn := btn
                    break
                }
                if (nameMatchNoClass = "")
                    nameMatchNoClass := btnClass
            }
        }
    }

    if (!modelBtn)
        return false

    ; Click the model button
    try {
        modelBtn.Click()
        return true
    } catch as e {
        return false
    }
}

; Poll the AI model selection button until Name no longer contains "loading", or maxWaitMs elapses.
; Button: Type 50000, ClassName "flex items-center gap-2 hover:text-text/80 transition-colors "
; Returns true when loading text disappeared, false on timeout or if button not found.
Handy_WaitForModelReady(hwnd, maxWaitMs) {
    global UIA
    pollInterval := 250
    start := A_TickCount
    firstLog := true
    loop {
        if ((A_TickCount - start) >= maxWaitMs) {
            return false
        }
        el := UIA.ElementFromHandle(hwnd)
        if !el {
            Sleep pollInterval
            continue
        }
        btn := 0
        try btn := el.FindFirst({ Type: 50000, ClassName: "flex items-center gap-2 hover:text-text/80 transition-colors " })
        if (!btn) {
            if (firstLog) {
                firstLog := false
            }
            Sleep pollInterval
            continue
        }
        btnName := ""
        try btnName := btn.Name
        if (InStr(btnName, "loading") = 0) {
            return true
        }
        Sleep pollInterval
    }
}
