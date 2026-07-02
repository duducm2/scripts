; =============================================================================
; Utils module: hotstring_gemini_banner.ahk
; Hotstring Gemini banner and D2C preset helpers
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; Hotstring Selector: Gemini Redirect Banner (non-blocking; uses standard loading indicator)
; =============================================================================
HotstringGeminiBanner_Show(text := "📤 Gemini: inserting prompt...") {
    StandardLoadingBar_CloseKeysOverlay()
    Sleep 50
    StandardLoadingBar_Show(text, BANNER_ACCENT_INTERMEDIATE, { passive: true, centerOnHwnd: 0, textWidth: 280,
        fontSize: 17,
        alpha: 204 })
}

HotstringGeminiBanner_Hide(*) {
    StandardLoadingBar_Hide(0)
}

; Dictation → Gemini: join preset prompt and dictated text for InsertText into Gemini prompt field.
D2C_CombinePresetWithDictation(presetText, dictationText) {
    p := Trim(presetText)
    d := Trim(dictationText)
    if (p = "")
        return d
    if (d = "")
        return p
    return p . "`n`n" . d
}

; Buffer after dictation ends before "Send to Gemini?" (linear loading bar + accidental key buffer).
D2C_SUBMIT_MENU_DELAY_MS := 2000
D2C_SUBMIT_MENU_PROGRESS_TICK_MS := 50
global g_D2C_PreparingMenuCancelledByV := false

D2C_OnPreparingMenuCtrlV(*) {
    global g_D2C_PreparingMenuCancelledByV
    g_D2C_PreparingMenuCancelledByV := true
}

; Linear 0–100% loading bar for D2C_SUBMIT_MENU_DELAY_MS, then hide (no keys). Stops StandardLoadingBar_Tick spinner.
; Returns true when delay completes; false when user presses Ctrl+V (same as menu [V] — paste dictated text only).
D2C_RunSubmitMenuDelayBar() {
    global g_StandardLoadingBarGui, g_D2C_PreparingMenuCancelledByV
    delayMs := D2C_SUBMIT_MENU_DELAY_MS
    tickMs := D2C_SUBMIT_MENU_PROGRESS_TICK_MS
    g_D2C_PreparingMenuCancelledByV := false
    StandardLoadingBar_CloseKeysOverlay()
    StandardLoadingBar_Hide(0)
    StandardLoadingBar_Show("⏳ Preparing menu…", BANNER_ACCENT_INTERMEDIATE, { textWidth: 520, fontSize: 17,
        noBorder: true, trackActiveMonitor: true })
    SetTimer(StandardLoadingBar_Tick, 0)
    try {
        if IsObject(g_StandardLoadingBarGui)
            g_StandardLoadingBarGui["OverlayProg"].Value := 0
    } catch {
    }
    steps := delayMs // tickMs
    if (steps < 1)
        steps := 1
    try {
        Hotkey("$^v", D2C_OnPreparingMenuCtrlV, "On")
        loop steps {
            if g_D2C_PreparingMenuCancelledByV
                break
            Sleep(tickMs)
            if g_D2C_PreparingMenuCancelledByV
                break
            pct := Min(100, Round(100 * A_Index / steps))
            try {
                if IsObject(g_StandardLoadingBarGui)
                    g_StandardLoadingBarGui["OverlayProg"].Value := pct
            } catch {
                break
            }
        }
        if !g_D2C_PreparingMenuCancelledByV {
            try {
                if IsObject(g_StandardLoadingBarGui)
                    g_StandardLoadingBarGui["OverlayProg"].Value := 100
            } catch {
            }
        }
    } finally {
        try Hotkey("$^v", "Off")
        catch {
        }
        StandardLoadingBar_Hide(0)
    }
    return !g_D2C_PreparingMenuCancelledByV
}

TryGetSelectedTextViaUIA_QuickLook() {
    hwnd := WinExist("A")
    focused := 0
    try {
        focused := UIA.GetFocusedElement()
    } catch {
        focused := 0
    }

    if (focused && focused.IsTextPatternAvailable) {
        try {
            ranges := focused.TextPattern.GetSelection()
            if (IsObject(ranges) && ranges.Length >= 1) {
                txt := ranges[1].GetText(512)
                return Trim(txt)
            }
        } catch {
        }
    }

    try {
        root := UIA.ElementFromHandle(hwnd)
        doc := root.FindFirst({ Type: UIA.ControlType.Document })
        if (doc && doc.IsTextPatternAvailable) {
            try {
                ranges := doc.TextPattern.GetSelection()
                if (IsObject(ranges) && ranges.Length >= 1) {
                    txt := ranges[1].GetText(512)
                    if (Trim(txt) != "")
                        return Trim(txt)
                }
            } catch {
            }
            try {
                txt := doc.TextPattern.DocumentRange.GetText(512)
                return Trim(txt)
            } catch {
            }
        }
    } catch {
    }
    return ""
}

TryCopySelectionToClipboard_QuickLookAware() {
    proc := ""
    try {
        proc := WinGetProcessName("A")
    } catch {
        proc := ""
    }

    A_Clipboard := ""
    Send "^c"
    if ClipWait(0.7)
        return true

    A_Clipboard := ""
    Send "^{Insert}"
    if ClipWait(0.7)
        return true

    if (proc = "QuickLook.exe") {
        A_Clipboard := ""
        Send "{AppsKey}"
        Sleep 60
        Send "c"
        if ClipWait(0.9)
            return true

        txt := TryGetSelectedTextViaUIA_QuickLook()
        if (txt != "" && StrLen(Trim(txt)) > 0) {
            A_Clipboard := txt
            return true
        }
    }

    return false
}

; Heuristic pt/en/de when Python daemon or lingua is unavailable (#!+8 auto-detect timeout path).
DetectLang_AhkFallback(text) {
    static inited := false
    static ptMap := Map()
    static enMap := Map()
    static deMap := Map()
    if (!inited) {
        inited := true
        for w in StrSplit("que nao com para uma dos das por mas sao esta ser tem foi", " ") {
            if (w != "")
                ptMap[w] := true
        }
        for w in StrSplit("the and of to in is that it for on with as by this are", " ") {
            if (w != "")
                enMap[w] := true
        }
        for w in StrSplit("der die das und ist nicht ein eine mit auf fur sich von zu als", " ") {
            if (w != "")
                deMap[w] := true
        }
    }
    t := Trim(text)
    if (t = "")
        return "en"
    tl := StrLower(t)
    if RegExMatch(tl, "[ãõçáàâéêíóôú]")
        return "pt"
    if RegExMatch(tl, "[äöüß]")
        return "de"
    clean := RegExReplace(tl, "[^a-zà-ÿ]+", " ")
    scorePt := 0
    scoreEn := 0
    scoreDe := 0
    for word in StrSplit(RegExReplace(clean, " +", " "), " ") {
        if (word = "")
            continue
        if ptMap.Has(word)
            scorePt++
        if enMap.Has(word)
            scoreEn++
        if deMap.Has(word)
            scoreDe++
    }
    m := Max(scorePt, scoreEn, scoreDe)
    if (m = 0)
        return "en"
    wins := []
    if (scorePt = m)
        wins.Push("pt")
    if (scoreEn = m)
        wins.Push("en")
    if (scoreDe = m)
        wins.Push("de")
    if (wins.Length = 1)
        return wins[1]
    return "en"
}
