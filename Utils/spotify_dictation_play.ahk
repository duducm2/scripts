; =============================================================================
; Utils module: spotify_dictation_play.ahk
; Send dictation? [P] — open Spotify, Ctrl+K search, paste, Enter, immerse
; =============================================================================

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

; Activate existing Spotify or launch without #!+s (avoids Spotify.ahk banner fight).
SpotifyDictation_ActivateOrOpen() {
    hwnd := SpotifyDictation_ResolveMainHwnd()
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
        return true
    }

    StandardLoadingBar_Update("⏳ Opening Spotify...")
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

    StandardLoadingBar_Update("⏳ Waiting for Spotify...")
    deadline := A_TickCount + 30000
    hwnd := 0
    while (A_TickCount < deadline) {
        hwnd := SpotifyDictation_ResolveMainHwnd()
        if (hwnd > 0)
            break
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
    ; Cold start: let the UI become interactive before Ctrl+K.
    StandardLoadingBar_Update("⏳ Spotify loading...")
    Sleep 2000
    WinActivate("ahk_id " hwnd)
    Sleep 150
    return true
}

; End-to-end: open Spotify → Ctrl+K → paste → Enter → SpotifyImmerse(false).
SpotifyDictation_PlayFromClipboard(messageText) {
    StandardLoadingBar_Show("⏳ Opening Spotify...", BANNER_ACCENT_INTERMEDIATE, {
        fontSize: 17,
        trackActiveMonitor: true
    })

    clipSaved := ClipboardAll()
    barOwned := true
    try {
        if (!SpotifyDictation_ActivateOrOpen()) {
            ; ActivateOrOpen already hid the bar and showed an error overlay.
            barOwned := false
            return false
        }

        Send "{LWin Up}{RWin Up}{LAlt Up}{RAlt Up}{LShift Up}{RShift Up}"
        Sleep 80

        hwnd := SpotifyDictation_ResolveMainHwnd()
        if (hwnd > 0) {
            WinActivate("ahk_id " hwnd)
            WinWaitActive("ahk_id " hwnd, , 2)
        }

        ; Native Spotify search (Shift+S maps to this).
        StandardLoadingBar_Update("⏳ Opening Spotify search...")
        Send "^k"
        Sleep 300

        StandardLoadingBar_Update("⏳ Pasting search query...")
        loop 5 {
            A_Clipboard := ""
            A_Clipboard := messageText
            if ClipWait(2) && (A_Clipboard = messageText)
                break
            if A_Index = 5 {
                StandardLoadingBar_Hide(0)
                ShowCenteredOverlay_Utils("❌ CLIPBOARD ERROR - TRY AGAIN", 3000, BANNER_ACCENT_ERROR)
                barOwned := false
                return false
            }
            Sleep 100
        }
        Send "^v"
        ; Wait for Spotify search suggestions to populate before confirming with Enter.
        StandardLoadingBar_Update("⏳ Waiting for search suggestions...")
        Sleep 900
        Send "{Enter}"

        ; Let search → entity page settle before header Play UIA.
        StandardLoadingBar_Update("⏳ Waiting for Spotify results...")
        Sleep 1800

        StandardLoadingBar_Update("⏳ Immersing...")
        SpotifyImmerse(false)
        ; Immersion shows its own result toast — do not Hide over it.
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
