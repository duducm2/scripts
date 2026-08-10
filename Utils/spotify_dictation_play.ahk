; =============================================================================
; Utils module: spotify_dictation_play.ahk
; Send dictation? [P] — open Spotify, Ctrl+K search, type query, Enter, immerse
; =============================================================================

; #region agent log
SpotifyDictation_DebugLog(hypothesisId, location, message, data := "") {
    try {
        dataStr := "{}"
        if (data is Map) {
            parts := ""
            for k, v in data {
                if (parts != "")
                    parts .= ","
                if (v is String) {
                    vv := StrReplace(v, "\", "\\")
                    vv := StrReplace(vv, "`"", "\`"")
                    vv := StrReplace(vv, "`n", "\n")
                    vv := StrReplace(vv, "`r", "")
                    parts .= "`"" k "`":`"" vv "`""
                } else
                    parts .= "`"" k "`":" v
            }
            dataStr := "{" parts "}"
        }
        line := "{`"sessionId`":`"79788c`",`"hypothesisId`":`"" hypothesisId "`",`"location`":`"" location "`",`"message`":`"" message "`",`"data`":" dataStr ",`"timestamp`":" A_TickCount ",`"runId`":`"post-fix`"}`n"
        FileAppend(line, A_ScriptDir "\debug-79788c.log", "UTF-8")
    } catch {
    }
}
SpotifyDictation_DebugActive(hypothesisId, location, message) {
    try {
        hwnd := WinExist("A")
        title := "", proc := ""
        try title := WinGetTitle("ahk_id " hwnd)
        try proc := WinGetProcessName("ahk_id " hwnd)
        SpotifyDictation_DebugLog(hypothesisId, location, message, Map(
            "hwnd", hwnd,
            "proc", proc,
            "title", SubStr(title, 1, 80),
            "clipLen", StrLen(A_Clipboard)
        ))
    } catch {
    }
}
; #endregion

SpotifyDictation_ResolveMainHwnd() {
    bestWithTitle := 0
    bestWithTitleArea := 0
    bestAny := 0
    bestAnyArea := 0
    for hwnd in WinGetList("ahk_exe Spotify.exe") {
        if !(hwnd is Integer) || hwnd <= 0
            continue
        if !DllCall("IsWindowVisible", "Ptr", hwnd)
            continue
        try {
            WinGetPos(, , &w, &h, "ahk_id " hwnd)
            if (w < 200 || h < 200)
                continue
            area := w * h
        } catch {
            continue
        }
        try title := WinGetTitle(hwnd)
        catch
            title := ""
        if InStr(title, "Spotify") {
            if (area > bestWithTitleArea) {
                bestWithTitleArea := area
                bestWithTitle := hwnd
            }
        } else if (area > bestAnyArea) {
            bestAnyArea := area
            bestAny := hwnd
        }
    }
    return bestWithTitle ? bestWithTitle : bestAny
}

SpotifyDictation_ResolveLaunchPath() {
    path := A_AppData "\Microsoft\Windows\Start Menu\Programs\Spotify.lnk"
    if FileExist(path)
        return path
    try {
        path := A_ProgramsCommon "\Spotify.lnk"
        if FileExist(path)
            return path
    } catch {
    }
    localAppData := EnvGet("LocalAppData")
    for candidate in [A_AppData "\Spotify\Spotify.exe", localAppData "\Microsoft\WindowsApps\Spotify.exe"] {
        if FileExist(candidate)
            return candidate
    }
    return ""
}

SpotifyDictation_IsWindowResponsive(hwnd, timeoutMs := 300) {
    if !(hwnd is Integer) || hwnd <= 0 || !WinExist("ahk_id " hwnd)
        return false
    try {
        if DllCall("IsHungAppWindow", "Ptr", hwnd)
            return false
    } catch {
    }
    result := 0
    ok := DllCall("SendMessageTimeout", "Ptr", hwnd, "UInt", 0, "Ptr", 0, "Ptr", 0, "UInt", 0x0002, "UInt",
        timeoutMs, "Ptr*", &result)
    return ok != 0
}

; Shell gate: responsive + Edit + at least one nav chrome control (Home/Search/Library).
SpotifyDictation_IsUiReady(hwnd) {
    if !(hwnd is Integer) || hwnd <= 0
        return false
    if !SpotifyDictation_IsWindowResponsive(hwnd)
        return false
    try {
        spot := UIA_Browser("ahk_id " hwnd)
        if (!spot)
            return false
        hasEdit := false
        hasNav := false
        try {
            if (spot.FindFirst({ Type: 50004 }))
                hasEdit := true
        } catch {
        }
        for nm in ["Search", "Home", "Your Library", "Biblioteca", "Início"] {
            try {
                if (spot.FindFirst({ Type: 50000, Name: nm })) {
                    hasNav := true
                    break
                }
            } catch {
            }
        }
        if (!hasNav) {
            try {
                if (spot.FindFirst({ Type: 50004, Name: "Search" }))
                    hasNav := true
            } catch {
            }
        }
        return hasEdit && hasNav
    } catch {
    }
    return false
}

; After Ctrl+K: is keyboard focus on an Edit (search field)?
SpotifyDictation_IsSearchEditFocused() {
    try {
        fe := UIA.GetFocusedElement()
        if (!fe)
            return false
        tp := 0
        nm := ""
        try tp := fe.Type
        try nm := fe.Name
        if (tp = 50004 || tp = "Edit")
            return true
        if (nm && RegExMatch(nm, "i)(search|what do you want to (play|listen)|ouvir|buscar)"))
            return true
    } catch {
    }
    return false
}

; Poll until shell gate passes. Updates loading bar with elapsed wait.
SpotifyDictation_WaitUntilUiReady(initialHwnd := 0, timeoutMs := 25000, neededStreak := 2, openStart := 0) {
    if (!openStart)
        openStart := A_TickCount
    deadline := A_TickCount + timeoutMs
    readyStreak := 0
    lastUpdate := 0
    while (A_TickCount < deadline) {
        hwnd := SpotifyDictation_ResolveMainHwnd()
        if (hwnd <= 0 && initialHwnd > 0 && WinExist("ahk_id " initialHwnd))
            hwnd := initialHwnd
        elapsed := Round((A_TickCount - openStart) / 1000)
        if ((A_TickCount - lastUpdate) >= 800) {
            StandardLoadingBar_Update("⏳ Waiting until Spotify is fully ready... (" elapsed "s)")
            lastUpdate := A_TickCount
        }
        if (hwnd > 0 && SpotifyDictation_IsUiReady(hwnd)) {
            readyStreak += 1
            if (readyStreak >= neededStreak) {
                WinActivate("ahk_id " hwnd)
                WinWaitActive("ahk_id " hwnd, , 2)
                ; Extra settle after gate — shell exists but SPA may still hydrate.
                StandardLoadingBar_Update("⏳ Spotify shell ready — finishing load...")
                Sleep 1200
                return hwnd
            }
        } else {
            readyStreak := 0
        }
        Sleep 200
    }
    return 0
}

; Quality gate: Ctrl+K until search Edit is focused (retries). Returns true if focused.
SpotifyDictation_OpenSearchUntilFocused(hwnd, maxAttempts := 8) {
    loop maxAttempts {
        StandardLoadingBar_Update("⏳ Opening search (attempt " A_Index "/" maxAttempts ")...")
        if (hwnd > 0) {
            WinActivate("ahk_id " hwnd)
            WinWaitActive("ahk_id " hwnd, , 1)
        }
        Send "{LWin Up}{RWin Up}{LAlt Up}{RAlt Up}{LShift Up}{RShift Up}{LCtrl Up}{RCtrl Up}"
        Sleep 50
        Send "^k"
        Sleep 500
        if (SpotifyDictation_IsSearchEditFocused()) {
            ; #region agent log
            SpotifyDictation_DebugLog("I", "spotify_dictation_play.ahk:search_gate", "search edit focused", Map(
                "attempt", A_Index
            ))
            ; #endregion
            Sleep 200
            return true
        }
        Sleep 350
    }
    ; #region agent log
    SpotifyDictation_DebugLog("I", "spotify_dictation_play.ahk:search_gate", "search edit NOT focused", Map(
        "attempt", maxAttempts
    ))
    ; #endregion
    return false
}

; Type query into Spotify (ControlSendText proven; bare ControlSend "v" reached search).
SpotifyDictation_TypeSearchQuery(hwnd, messageText) {
    StandardLoadingBar_Update("⏳ Typing search query...")
    if (hwnd > 0) {
        try {
            ControlSendText messageText, , "ahk_id " hwnd
            return true
        } catch {
        }
        try {
            ControlSend "{Text}" messageText, , "ahk_id " hwnd
            return true
        } catch {
        }
    }
    SendText messageText
    return true
}

; Activate existing Spotify or launch. Loading bar must already be showing.
SpotifyDictation_ActivateOrOpen() {
    hwnd := SpotifyDictation_ResolveMainHwnd()
    coldStart := false
    openStart := A_TickCount
    if (hwnd > 0) {
        StandardLoadingBar_Update("⏳ Activating Spotify...")
        WinActivate("ahk_id " hwnd)
        if !WinWaitActive("ahk_id " hwnd, , 3) {
            WinActivate("ahk_id " hwnd)
            if !WinWaitActive("ahk_id " hwnd, , 2) {
                StandardLoadingBar_Hide(0)
                ShowCenteredOverlay_Utils("❌ Could not activate Spotify.", 2000, BANNER_ACCENT_ERROR)
                return false
            }
        }
    } else {
        coldStart := true
        StandardLoadingBar_Update("⏳ Opening Spotify — please wait...")
        launch := SpotifyDictation_ResolveLaunchPath()
        try {
            if (launch != "")
                Run('"' launch '"')
            else
                Run("spotify:")
        } catch {
            try Run("shell:AppsFolder\SpotifyAB.SpotifyMusic_zpdnekdrzrea0!Spotify")
            catch {
                StandardLoadingBar_Hide(0)
                ShowCenteredOverlay_Utils("❌ Could not launch Spotify.", 2000, BANNER_ACCENT_ERROR)
                return false
            }
        }

        StandardLoadingBar_Update("⏳ Waiting for Spotify window...")
        deadline := A_TickCount + 40000
        hwnd := 0
        while (A_TickCount < deadline) {
            hwnd := SpotifyDictation_ResolveMainHwnd()
            if (hwnd > 0)
                break
            elapsed := Round((A_TickCount - openStart) / 1000)
            if (Mod(elapsed, 2) = 0)
                StandardLoadingBar_Update("⏳ Waiting for Spotify window... (" elapsed "s)")
            Sleep 200
        }
        if (hwnd <= 0) {
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ Spotify did not start in time.", 2000, BANNER_ACCENT_ERROR)
            return false
        }

        WinActivate("ahk_id " hwnd)
        if !WinWaitActive("ahk_id " hwnd, , 5) {
            WinActivate("ahk_id " hwnd)
            if !WinWaitActive("ahk_id " hwnd, , 3) {
                StandardLoadingBar_Hide(0)
                ShowCenteredOverlay_Utils("❌ Could not activate Spotify.", 2000, BANNER_ACCENT_ERROR)
                return false
            }
        }
    }

    ; Cold start: stricter shell gate + longer streak. Warm: lighter.
    StandardLoadingBar_Update(coldStart ? "⏳ Waiting until Spotify is fully ready..." : "⏳ Checking Spotify UI...")
    readyTimeout := coldStart ? 45000 : 10000
    neededStreak := coldStart ? 10 : 3  ; cold ≈ 2s consecutive ready samples
    readyHwnd := SpotifyDictation_WaitUntilUiReady(hwnd, readyTimeout, neededStreak, openStart)
    ; #region agent log
    SpotifyDictation_DebugLog("G", "spotify_dictation_play.ahk:ActivateOrOpen:ready", "UI ready wait finished", Map(
        "coldStart", coldStart ? 1 : 0,
        "readyOk", readyHwnd > 0 ? 1 : 0,
        "waitMs", A_TickCount - openStart,
        "neededStreak", neededStreak
    ))
    ; #endregion
    if (readyHwnd <= 0) {
        StandardLoadingBar_Hide(0)
        ShowCenteredOverlay_Utils("❌ Spotify UI did not become ready in time.", 2500, BANNER_ACCENT_ERROR)
        return false
    }
    return true
}

; End-to-end: open Spotify → Ctrl+K (gated) → type → Enter → immerse.
; Loading Indication stays visible for the whole open/search/type path.
SpotifyDictation_PlayFromClipboard(messageText) {
    ; #region agent log
    SpotifyDictation_DebugLog("A", "spotify_dictation_play.ahk:PlayFromClipboard:entry", "messageText snapshot", Map(
        "msgLen", StrLen(messageText),
        "msgPreview", SubStr(messageText, 1, 40),
        "clipLenAtEntry", StrLen(A_Clipboard)
    ))
    ; #endregion
    StandardLoadingBar_Show("⏳ Opening Spotify — please wait...", BANNER_ACCENT_INTERMEDIATE, {
        fontSize: 17,
        trackActiveMonitor: true
    })

    clipSaved := ClipboardAll()
    barOwned := true
    try {
        if (!SpotifyDictation_ActivateOrOpen()) {
            ; #region agent log
            SpotifyDictation_DebugLog("D", "spotify_dictation_play.ahk:ActivateOrOpen", "activate_or_open_failed", Map(
                "ok", 0))
            ; #endregion
            barOwned := false
            return false
        }
        ; #region agent log
        SpotifyDictation_DebugLog("D", "spotify_dictation_play.ahk:ActivateOrOpen", "activate_or_open_ok", Map("ok", 1))
        ; #endregion

        Send "{LWin Up}{RWin Up}{LAlt Up}{RAlt Up}{LShift Up}{RShift Up}"
        Sleep 80

        hwnd := SpotifyDictation_ResolveMainHwnd()
        if (hwnd > 0) {
            WinActivate("ahk_id " hwnd)
            WinWaitActive("ahk_id " hwnd, , 2)
        }

        ; Keep loading bar visible (ControlSendText does not need the overlay hidden).
        if (!SpotifyDictation_OpenSearchUntilFocused(hwnd, 8)) {
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ Spotify search field did not accept focus", 2500, BANNER_ACCENT_ERROR)
            barOwned := false
            return false
        }

        ; #region agent log
        SpotifyDictation_DebugActive("B", "spotify_dictation_play.ahk:after_search_gate", "search gate passed")
        SpotifyDictation_DebugLog("H", "spotify_dictation_play.ahk:before_paste", "about to ControlSendText", Map(
            "msgLen", StrLen(messageText),
            "method", "controlsendtext_gated",
            "hwnd", hwnd
        ))
        ; #endregion

        try {
            SpotifyDictation_TypeSearchQuery(hwnd, messageText)
            Sleep 500
            ; #region agent log
            SpotifyDictation_DebugActive("C", "spotify_dictation_play.ahk:after_paste", "after ControlSendText")
            ; #endregion
        } catch as pasteErr {
            ; #region agent log
            SpotifyDictation_DebugLog("H", "spotify_dictation_play.ahk:paste_exception", "exception during type", Map(
                "err", SubStr(pasteErr.Message, 1, 120)
            ))
            ; #endregion
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ Could not type into Spotify search", 2200, BANNER_ACCENT_ERROR)
            barOwned := false
            return false
        }

        StandardLoadingBar_Update("⏳ Waiting for search suggestions...")
        Sleep 1000
        if (hwnd > 0) {
            WinActivate("ahk_id " hwnd)
            WinWaitActive("ahk_id " hwnd, , 1)
        }
        Send "{Enter}"

        StandardLoadingBar_Update("⏳ Waiting for Spotify results...")
        Sleep 1800

        StandardLoadingBar_Update("⏳ Immersing...")
        SpotifyImmerse(false)
        barOwned := false
        return true
    } finally {
        if (barOwned) {
            try StandardLoadingBar_Hide(0)
        }
        A_Clipboard := clipSaved
        if (ClipWait(1)) {
        }
    }
}
