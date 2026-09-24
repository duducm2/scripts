; =============================================================================
; Utils module: handy_uia_helpers.ahk
; Handy UIA helper functions and ShowAiModelSelector support
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; Handy UIA Helper Functions
; =============================================================================

; Background suppress (ClipAngel BeginFavoriteSuppress pattern): opacity 0 + off-screen.
; Spike: settings_store.json has selected_model, but Handy has no hot-reload IPC; CLI
; --model only applies to --transcribe-file. Live switches use suppressed UIA.
HANDY_SUPPRESS_OPACITY := 0
HANDY_SUPPRESS_OFFSCREEN_X := -32000
HANDY_SUPPRESS_OFFSCREEN_Y := -32000
HANDY_SUPPRESS_OFFSCREEN_W := 900
HANDY_SUPPRESS_OFFSCREEN_H := 700
HANDY_SUPPRESS_MIN_W := 640
HANDY_SUPPRESS_MIN_H := 480

global g_HandySuppressActive := false
global g_HandySuppressHadSavedPos := false
global g_HandySuppressSavedX := 0
global g_HandySuppressSavedY := 0
global g_HandySuppressSavedW := 0
global g_HandySuppressSavedH := 0
; g_HandyModelSwitchBusy is declared in handy_ai_model_config.ahk (loaded first).

; Restore the window that was focused before Handy automation (paste/dictation target).
; excludeHwnd: skip if it is Handy itself (already closed or still the only option).
Handy_RestorePrevWindow(restoreHwnd, excludeHwnd := 0) {
    if (!restoreHwnd || (excludeHwnd && restoreHwnd = excludeHwnd))
        return false
    if !WinExist("ahk_id " restoreHwnd)
        return false
    try {
        WinActivate("ahk_id " restoreHwnd)
        WinWaitActive("ahk_id " restoreHwnd, , 1)
        return !!WinActive("ahk_id " restoreHwnd)
    } catch {
        return false
    }
}

; Main Handy hwnd (includes hidden / suppressed). Environment-aware exe path filter.
Handy_MainHwnd() {
    expectedExePath := GetHandyProcessPath()
    prevDetect := A_DetectHiddenWindows
    DetectHiddenWindows true
    try {
        for hwnd in WinGetList("Handy ahk_class Tauri Window") {
            try {
                procPath := WinGetProcessPath(hwnd)
                if (expectedExePath = "" || StrCompare(procPath, expectedExePath, false) = 0)
                    return hwnd
            } catch {
                if (expectedExePath = "")
                    return hwnd
            }
        }
    } finally {
        DetectHiddenWindows prevDetect
    }
    return 0
}

; Resolve handy.exe for Run with --start-hidden (lnk targets and process path).
Handy_ResolveExePath() {
    expected := GetHandyProcessPath()
    if (expected != "" && FileExist(expected))
        return expected
    hwnd := Handy_MainHwnd()
    if (hwnd) {
        try {
            p := WinGetProcessPath("ahk_id " hwnd)
            if (p != "" && FileExist(p))
                return p
        } catch {
        }
    }
    shortcut := GetHandyShortcutPath()
    if (shortcut = "" || !FileExist(shortcut))
        return ""
    if (RegExMatch(shortcut, "i)\.exe$"))
        return shortcut
    try {
        FileGetShortcut(shortcut, &outTarget)
        if (outTarget != "" && FileExist(outTarget))
            return outTarget
    } catch {
    }
    return ""
}

Handy_ApplySuppressOpacity(hwnd) {
    if !hwnd
        return
    try WinSetTransparent(HANDY_SUPPRESS_OPACITY, "ahk_id " hwnd)
    catch {
    }
}

Handy_ClearSuppressOpacity(hwnd := 0) {
    if !hwnd
        hwnd := Handy_MainHwnd()
    if !hwnd
        return
    try WinSetTransparent("Off", "ahk_id " hwnd)
    catch {
    }
}

; Park Handy off-screen at opacity 0 for model-switch automation. Does not activate.
Handy_BeginSuppress(hwnd) {
    global g_HandySuppressActive, g_HandySuppressHadSavedPos
    global g_HandySuppressSavedX, g_HandySuppressSavedY
    global g_HandySuppressSavedW, g_HandySuppressSavedH
    if !hwnd
        return false
    g_HandySuppressActive := true
    Handy_ApplySuppressOpacity(hwnd)
    if !g_HandySuppressHadSavedPos {
        try {
            WinGetPos(&sx, &sy, &sw, &sh, "ahk_id " hwnd)
            if (sw > 0 && sh > 0
                && (sx > HANDY_SUPPRESS_OFFSCREEN_X + 1000 || sy > HANDY_SUPPRESS_OFFSCREEN_Y + 1000)) {
                g_HandySuppressSavedX := sx
                g_HandySuppressSavedY := sy
                g_HandySuppressSavedW := sw
                g_HandySuppressSavedH := sh
                g_HandySuppressHadSavedPos := true
            }
        } catch {
        }
    }
    w := g_HandySuppressHadSavedPos ? Max(HANDY_SUPPRESS_MIN_W, g_HandySuppressSavedW)
        : HANDY_SUPPRESS_OFFSCREEN_W
    h := g_HandySuppressHadSavedPos ? Max(HANDY_SUPPRESS_MIN_H, g_HandySuppressSavedH)
        : HANDY_SUPPRESS_OFFSCREEN_H
    if (w < HANDY_SUPPRESS_OFFSCREEN_W)
        w := HANDY_SUPPRESS_OFFSCREEN_W
    if (h < HANDY_SUPPRESS_OFFSCREEN_H)
        h := HANDY_SUPPRESS_OFFSCREEN_H
    try WinMove(HANDY_SUPPRESS_OFFSCREEN_X, HANDY_SUPPRESS_OFFSCREEN_Y, w, h, "ahk_id " hwnd)
    catch {
    }
    Handy_ApplySuppressOpacity(hwnd)
    try {
        mm := WinGetMinMax("ahk_id " hwnd)
        if (mm = -1 || mm = 1)
            WinRestore("ahk_id " hwnd)
    } catch {
    }
    Handy_ApplySuppressOpacity(hwnd)
    try WinMove(HANDY_SUPPRESS_OFFSCREEN_X, HANDY_SUPPRESS_OFFSCREEN_Y, w, h, "ahk_id " hwnd)
    catch {
    }
    Handy_ApplySuppressOpacity(hwnd)
    try WinShow("ahk_id " hwnd)
    catch {
    }
    Handy_ApplySuppressOpacity(hwnd)
    try WinMove(HANDY_SUPPRESS_OFFSCREEN_X, HANDY_SUPPRESS_OFFSCREEN_Y, w, h, "ahk_id " hwnd)
    catch {
    }
    Handy_ApplySuppressOpacity(hwnd)
    return true
}

; End suppress: restore saved geometry + opacity. showNormal=true activates on-screen.
Handy_EndSuppress(hwnd := 0, showNormal := false) {
    global g_HandySuppressActive, g_HandySuppressHadSavedPos
    global g_HandySuppressSavedX, g_HandySuppressSavedY
    global g_HandySuppressSavedW, g_HandySuppressSavedH
    if !hwnd
        hwnd := Handy_MainHwnd()
    if (hwnd) {
        Handy_ApplySuppressOpacity(hwnd)
        if g_HandySuppressHadSavedPos {
            try {
                WinMove(g_HandySuppressSavedX, g_HandySuppressSavedY,
                    g_HandySuppressSavedW, g_HandySuppressSavedH, "ahk_id " hwnd)
            } catch {
            }
        }
        Handy_ClearSuppressOpacity(hwnd)
        if (showNormal) {
            try WinShow("ahk_id " hwnd)
            catch {
            }
        }
    }
    g_HandySuppressActive := false
    g_HandySuppressHadSavedPos := false
}

; Ensure Handy exists for background automation: launch if needed, suppress, wait UI.
; Does not activate / steal focus. Returns hwnd or 0.
Handy_EnsureForAutomation(waitUiMs := 2000) {
    hwnd := Handy_MainHwnd()
    launched := false
    if (!hwnd) {
        exePath := Handy_ResolveExePath()
        targetPath := GetHandyShortcutPath()
        if (exePath != "" && FileExist(exePath)) {
            try Run('"' exePath '" --start-hidden')
            catch {
                try Run('"' exePath '"')
                catch
                    return 0
            }
        } else if (targetPath != "" && FileExist(targetPath)) {
            try Run targetPath
            catch
                return 0
        } else {
            return 0
        }
        launched := true
        prevDetect := A_DetectHiddenWindows
        DetectHiddenWindows true
        try {
            if !WinWait("Handy ahk_class Tauri Window", , 8)
                return 0
        } finally {
            DetectHiddenWindows prevDetect
        }
        hwnd := Handy_MainHwnd()
        if (!hwnd)
            return 0
        waitUiMs := Max(waitUiMs, 9000)
    }
    Handy_BeginSuppress(hwnd)
    if (Handy_WaitForMainUiReady(hwnd, waitUiMs))
        return hwnd
    ; WebView2 can throttle at -32000: park opacity-0 on the farthest monitor instead.
    Handy_SuppressMonitorFallback(hwnd)
    if (Handy_WaitForMainUiReady(hwnd, Min(waitUiMs, 2500)))
        return hwnd
    if (launched)
        return 0
    ; Warm instance: probe can flake while suppressed; still return hwnd for UIA attempts.
    return hwnd
}

; Fallback when off-screen (-32000) UIA is dead: opacity 0 on the monitor farthest from the cursor.
Handy_SuppressMonitorFallback(hwnd) {
    if !hwnd
        return false
    global g_HandySuppressActive
    g_HandySuppressActive := true
    Handy_ApplySuppressOpacity(hwnd)
    CoordMode "Mouse", "Screen"
    MouseGetPos &mx, &my
    bestIdx := 1
    bestDist := -1
    try {
        loop MonitorGetCount() {
            MonitorGetWorkArea(A_Index, &l, &t, &r, &b)
            cx := (l + r) // 2
            cy := (t + b) // 2
            dist := (cx - mx) * (cx - mx) + (cy - my) * (cy - my)
            if (dist > bestDist) {
                bestDist := dist
                bestIdx := A_Index
            }
        }
        MonitorGetWorkArea(bestIdx, &ml, &mt, &mr, &mb)
        ; Park just inside the far monitor (opacity 0 — invisible but WebView keeps painting).
        x := mr - HANDY_SUPPRESS_OFFSCREEN_W - 8
        y := mb - HANDY_SUPPRESS_OFFSCREEN_H - 8
        if (x < ml)
            x := ml
        if (y < mt)
            y := mt
        try {
            mm := WinGetMinMax("ahk_id " hwnd)
            if (mm = -1 || mm = 1)
                WinRestore("ahk_id " hwnd)
        } catch {
        }
        try WinMove(x, y, HANDY_SUPPRESS_OFFSCREEN_W, HANDY_SUPPRESS_OFFSCREEN_H, "ahk_id " hwnd)
        catch {
        }
        try WinShow("ahk_id " hwnd)
        catch {
        }
        Handy_ApplySuppressOpacity(hwnd)
        return true
    } catch {
        return false
    }
}

; True when Handy is on-screen and not in background suppress (user can see/edit it).
Handy_IsUserVisible(hwnd := 0) {
    global g_HandySuppressActive
    if !hwnd
        hwnd := Handy_MainHwnd()
    if !hwnd
        return false
    if (g_HandySuppressActive)
        return false
    prevDetect := A_DetectHiddenWindows
    DetectHiddenWindows false
    try {
        if !WinExist("ahk_id " hwnd)
            return false
    } finally {
        DetectHiddenWindows prevDetect
    }
    try {
        WinGetPos(&x, &y, &w, &h, "ahk_id " hwnd)
        if (w < 80 || h < 80)
            return false
        if (x <= HANDY_SUPPRESS_OFFSCREEN_X + 1000 && y <= HANDY_SUPPRESS_OFFSCREEN_Y + 1000)
            return false
    } catch {
        return false
    }
    try {
        if !DllCall("IsWindowVisible", "ptr", hwnd)
            return false
    } catch {
    }
    return true
}

; #region agent log
Handy_DebugLog(hypothesisId, location, message, data := "") {
    try {
        path := A_ScriptDir "\debug-e946b7.log"
        ts := A_TickCount
        dataJson := "{}"
        if (data != "") {
            if (data is String)
                dataJson := '{"raw":"' . StrReplace(StrReplace(data, "\", "\\"), '"', '\"') . '"}'
            else if (IsObject(data)) {
                parts := []
                for k, v in data.OwnProps() {
                    vv := v
                    if !(vv is Number)
                        vv := '"' . StrReplace(StrReplace(String(vv), "\", "\\"), '"', '\"') . '"'
                    parts.Push('"' k '":' vv)
                }
                dataJson := "{"
                for i, p in parts {
                    if (i > 1)
                        dataJson .= ","
                    dataJson .= p
                }
                dataJson .= "}"
            }
        }
        line := '{"sessionId":"e946b7","hypothesisId":"' hypothesisId '","location":"' location '","message":"' .
            StrReplace(StrReplace(message, "\", "\\"), '"', '\"') . '","data":' dataJson ',"timestamp":' ts ',"runId":"pre-fix"}`n'
        FileAppend(line, path)
    } catch {
    }
}
; #endregion

; Utility Shortcuts: toggle Handy visible ↔ background-suppressed (for manual settings edits).
Handy_ToggleVisible(*) {
    global g_HandySuppressActive, g_HandySuppressHadSavedPos
    ; #region agent log
    Handy_DebugLog("A", "Handy_ToggleVisible:entry", "toggle invoked", {
        script: A_ScriptName, suppress: g_HandySuppressActive, hadSaved: g_HandySuppressHadSavedPos })
    ; #endregion
    hwnd := Handy_MainHwnd()
    ; #region agent log
    posX := "", posY := "", posW := "", posH := "", visDll := -1
    if (hwnd) {
        try {
            WinGetPos(&posX, &posY, &posW, &posH, "ahk_id " hwnd)
            visDll := DllCall("IsWindowVisible", "ptr", hwnd)
        } catch {
        }
    }
    Handy_DebugLog("B", "Handy_ToggleVisible:hwnd", "main hwnd probe", {
        hwnd: hwnd, x: posX, y: posY, w: posW, h: posH, isVisibleDll: visDll })
    ; #endregion
    if (!hwnd) {
        hwnd := Handy_ActivateOrLaunch()
        ; #region agent log
        Handy_DebugLog("B", "Handy_ToggleVisible:launch", "no hwnd path ActivateOrLaunch", { hwnd: hwnd })
        ; #endregion
        if (!hwnd) {
            ShowCenteredOverlay_Utils("❌ Handy not available", 2000, BANNER_ACCENT_ERROR)
            return false
        }
        ShowCenteredOverlay_Utils("✅ Handy open", 900, BANNER_ACCENT_SUCCESS)
        return true
    }
    userVis := Handy_IsUserVisible(hwnd)
    ; #region agent log
    Handy_DebugLog("C", "Handy_ToggleVisible:branch", "visibility branch", {
        userVisible: userVis, suppress: g_HandySuppressActive })
    ; #endregion
    if (userVis) {
        Handy_BeginSuppress(hwnd)
        ShowCenteredOverlay_Utils("👻 Handy hidden (background)", 900, BANNER_ACCENT_INFO)
        return true
    }
    hwnd := Handy_ActivateOrLaunch()
    ; #region agent log
    ax := "", ay := "", aw := "", ah := "", afterSuppress := g_HandySuppressActive
    if (hwnd) {
        try WinGetPos(&ax, &ay, &aw, &ah, "ahk_id " hwnd)
        catch {
        }
    }
    Handy_DebugLog("D", "Handy_ToggleVisible:afterShow", "ActivateOrLaunch result", {
        hwnd: hwnd, x: ax, y: ay, w: aw, h: ah, suppressAfter: afterSuppress,
        active: WinActive("ahk_id " hwnd) })
    ; #endregion
    if (!hwnd) {
        ShowCenteredOverlay_Utils("❌ Could not show Handy", 2000, BANNER_ACCENT_ERROR)
        return false
    }
    ShowCenteredOverlay_Utils("✅ Handy visible", 900, BANNER_ACCENT_SUCCESS)
    return true
}

; Activate existing Handy window or launch it; returns hwnd or 0.
; Ends suppress first so the window is visible for interactive flows.
Handy_ActivateOrLaunch() {
    matchingHwnd := Handy_MainHwnd()
    if (matchingHwnd) {
        global g_HandySuppressActive
        ; #region agent log
        Handy_DebugLog("D", "Handy_ActivateOrLaunch:existing", "found hwnd", {
            hwnd: matchingHwnd, suppress: g_HandySuppressActive })
        ; #endregion
        if (g_HandySuppressActive)
            Handy_EndSuppress(matchingHwnd, true)
        ; #region agent log
        ; Also detect physical off-screen even if suppress flag cleared (e.g. after reload).
        try {
            WinGetPos(&ex, &ey, &ew, &eh, "ahk_id " matchingHwnd)
            Handy_DebugLog("E", "Handy_ActivateOrLaunch:preActivate", "geometry before activate", {
                x: ex, y: ey, w: ew, h: eh, suppress: g_HandySuppressActive })
        } catch {
        }
        ; #endregion
        WinActivate("ahk_id " . matchingHwnd)
        WinWaitActive("ahk_id " . matchingHwnd, , 2)
        Handy_WaitForMainUiReady(matchingHwnd, 2000)
        return matchingHwnd
    }

    exePath := Handy_ResolveExePath()
    targetPath := GetHandyShortcutPath()
    if (exePath != "" && FileExist(exePath)) {
        try Run('"' exePath '"')
        catch {
            if (targetPath = "" || !FileExist(targetPath))
                return 0
            Run targetPath
        }
    } else if (targetPath != "" && FileExist(targetPath)) {
        Run targetPath
    } else {
        return 0
    }

    prevDetect := A_DetectHiddenWindows
    DetectHiddenWindows true
    try {
        if !WinWait("Handy ahk_class Tauri Window", , 8)
            return 0
    } finally {
        DetectHiddenWindows prevDetect
    }

    h := Handy_MainHwnd()
    if (!h)
        return 0
    WinActivate("ahk_id " . h)
    WinWaitActive("ahk_id " . h, , 2)
    if (!Handy_WaitForMainUiReady(h, 9000))
        return 0
    return h
}

; Wait for Handy main UI to be interactive (needed most on cold launch).
Handy_WaitForMainUiReady(hwnd, maxWaitMs := 9000) {
    global UIA
    start := A_TickCount
    pollMs := 100
    prevDetect := A_DetectHiddenWindows
    DetectHiddenWindows true
    try {
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
                ; Model button alone is enough when already on History/Models tabs.
                try {
                    if (Handy_FindActiveAiModelButton(el))
                        return true
                }
            }
            Sleep pollMs
        }
    } finally {
        DetectHiddenWindows prevDetect
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

; Open the AI model dropdown via direct UIA click on the header model button (no focus steal).
; Falls back to keyboard navigation only when direct click fails (requires activation).
Handy_OpenAiModelMenu(hwnd) {
    el := UIA.ElementFromHandle(hwnd)
    if !el {
        return false
    }

    ; Preferred: click the active model selector button (handy.md header button).
    modelBtn := Handy_FindActiveAiModelButton(el)
    if (modelBtn) {
        try modelBtn.Click()
        catch {
            try modelBtn.Invoke()
            catch {
                modelBtn := 0
            }
        }
        if (modelBtn && Handy_WaitForAiModelMenuOpen(hwnd, 1800))
            return true
    }

    ; Fallback: anchor + Shift+Tab + Enter (needs foreground focus).
    global g_HandySuppressActive
    if (g_HandySuppressActive) {
        ; Re-assert suppress after any accidental activate from fallback prep.
        Handy_BeginSuppress(hwnd)
    }
    try WinActivate("ahk_id " . hwnd)
    catch {
    }

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
        if (g_HandySuppressActive)
            Handy_BeginSuppress(hwnd)
        return false
    }

    try anchor.SetFocus()
    catch {
        try anchor.Click()
    }
    Sleep 40
    ; Prefer ControlSend so keys go to Handy even if focus races.
    try ControlSend("+{Tab}", , "ahk_id " hwnd)
    catch
        Send "+{Tab}"
    Sleep 40
    try ControlSend("{Enter}", , "ahk_id " hwnd)
    catch
        Send "{Enter}"

    opened := Handy_WaitForAiModelMenuOpen(hwnd, 1800)
    if (g_HandySuppressActive)
        Handy_BeginSuppress(hwnd)
    return opened
}

; Wait for AI model context menu rows to appear after opening the menu.
Handy_WaitForAiModelMenuOpen(hwnd, maxWaitMs := 1800) {
    global UIA
    start := A_TickCount
    pollMs := 50
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

; Active model selector button in Handy header (not menu row).
Handy_FindActiveAiModelButton(el) {
    if !el
        return 0
    try {
        return el.FindFirst({ Type: 50000, ClassName: "flex items-center gap-2 hover:text-text/80 transition-colors " })
    } catch {
        return 0
    }
}

; Returns the active model button label, or "" if not found.
Handy_ReadActiveAiModelName(hwnd) {
    global UIA
    el := UIA.ElementFromHandle(hwnd)
    if !el
        return ""
    btn := Handy_FindActiveAiModelButton(el)
    if !btn
        return ""
    try {
        return btn.Name
    } catch {
        return ""
    }
}

; Quality gate: active model name must contain modelName and not be loading.
Handy_VerifyAiModelActive(hwnd, modelName) {
    if (modelName = "")
        return false
    name := Handy_ReadActiveAiModelName(hwnd)
    if (name = "")
        return false
    if (InStr(name, "loading"))
        return false
    return InStr(name, modelName)
}

; Open menu, click model, wait for load — all must succeed.
Handy_TrySelectAiModel(hwnd, modelClickName, maxWaitMs := 20000) {
    if !Handy_OpenAiModelMenu(hwnd)
        return false
    if !Handy_ClickAiModel(hwnd, modelClickName)
        return false
    return Handy_WaitForModelReady(hwnd, maxWaitMs)
}

; Close a stuck model menu before retrying. Prefer ControlSend; re-assert suppress after.
Handy_DismissOpenUi(hwnd) {
    if !hwnd
        return
    global g_HandySuppressActive
    try ControlSend("{Escape}", , "ahk_id " hwnd)
    catch {
        try WinActivate("ahk_id " . hwnd)
        Sleep 40
        Send "{Escape}"
    }
    Sleep 40
    try ControlSend("{Escape}", , "ahk_id " hwnd)
    catch
        Send "{Escape}"
    Sleep 40
    if (g_HandySuppressActive)
        Handy_BeginSuppress(hwnd)
}

; Poll the AI model selection button until Name no longer contains "loading", or maxWaitMs elapses.
; Button: Type 50000, ClassName "flex items-center gap-2 hover:text-text/80 transition-colors "
; Returns true when loading text disappeared, false on timeout or if button not found.
Handy_WaitForModelReady(hwnd, maxWaitMs) {
    global UIA
    pollInterval := 100
    start := A_TickCount
    loop {
        if ((A_TickCount - start) >= maxWaitMs)
            return false
        el := UIA.ElementFromHandle(hwnd)
        if !el {
            Sleep pollInterval
            continue
        }
        btn := Handy_FindActiveAiModelButton(el)
        if (!btn) {
            Sleep pollInterval
            continue
        }
        btnName := ""
        try btnName := btn.Name
        if (InStr(btnName, "loading") = 0)
            return true
        Sleep pollInterval
    }
}
