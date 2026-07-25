; =============================================================================
; Shift keys module: hotif_spotify.ahk
; Spotify hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

#HotIf WinActive("ahk_exe Spotify.exe") && !IsFileDialogActive()

; Shift + C : Toggle Connect to a device - Connect to a device
+c::
{
    ; #region agent log
    SafeDebugLog Format(
        "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:7574`",`"message`":`"Connect button handler entry`",`"data`":{},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"A`"}`n",
        A_TickCount, Random(1000, 9999), A_TickCount)
    ; #endregion
    try {
        spot := UIA_Browser("ahk_exe Spotify.exe")
        Sleep 300

        ; Find and click the Connect button
        connectPattern := "i)^(Connect to a device|Conectar a um dispositivo|Connect)$"
        ; #region agent log
        SafeDebugLog Format(
            "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:7582`",`"message`":`"Before WaitForButton call`",`"data`":{`"pattern`":`"{4}`"},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"B`"}`n",
            A_TickCount, Random(1000, 9999), A_TickCount, connectPattern)
        ; #endregion
        if (connectBtn := WaitForButton(spot, connectPattern)) {
            ; #region agent log
            btnName := "", btnType := "", btnClassName := "", btnLocalizedType := "", supportsInvoke := false,
                supportsToggle := false
            try btnName := connectBtn.Name
            try btnType := connectBtn.ControlType
            try btnClassName := connectBtn.ClassName
            try btnLocalizedType := connectBtn.LocalizedControlType
            try supportsInvoke := connectBtn.GetPropertyValue(UIA.Property.IsInvokePatternAvailable)
            try supportsToggle := connectBtn.GetPropertyValue(UIA.Property.IsTogglePatternAvailable)
            SafeDebugLog Format(
                "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:7583`",`"message`":`"Button found before Invoke`",`"data`":{`"name`":`"{4}`",`"type`":`"{5}`",`"className`":`"{6}`",`"localizedType`":`"{7}`",`"supportsInvoke`":{8},`"supportsToggle`":{9}},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"A,C`"}`n",
                A_TickCount, Random(1000, 9999), A_TickCount, btnName, btnType, btnClassName, btnLocalizedType,
                supportsInvoke, supportsToggle)
            ; #endregion
            ; Try multi-strategy activation: prefer Invoke when available, fallback to Click
            clicked := false
            if (supportsInvoke) {
                try {
                    connectBtn.Invoke()
                    clicked := true
                    ; #region agent log
                    SafeDebugLog Format(
                        "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:7663`",`"message`":`"Invoke() succeeded`",`"data`":{},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"A`"}`n",
                        A_TickCount, Random(1000, 9999), A_TickCount)
                    ; #endregion
                } catch Error as invokeErr {
                    ; #region agent log
                    SafeDebugLog Format(
                        "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:7663`",`"message`":`"Invoke() failed, trying Click()`",`"data`":{`"error`":`"{4}`"},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"A,E`"}`n",
                        A_TickCount, Random(1000, 9999), A_TickCount, invokeErr.Message)
                    ; #endregion
                }
            }
            if (!clicked) {
                try {
                    connectBtn.Click()
                    clicked := true
                    ; #region agent log
                    SafeDebugLog Format(
                        "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:7675`",`"message`":`"Click() succeeded`",`"data`":{},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"A,E`"}`n",
                        A_TickCount, Random(1000, 9999), A_TickCount)
                    ; #endregion
                } catch Error as clickErr {
                    ; #region agent log
                    SafeDebugLog Format(
                        "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:7675`",`"message`":`"Click() failed`",`"data`":{`"error`":`"{4}`"},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"A,E`"}`n",
                        A_TickCount, Random(1000, 9999), A_TickCount, clickErr.Message)
                    ; #endregion
                    MsgBox "Error: " clickErr.Message
                    return
                }
            }

            ; After connecting, wait a moment for the device list to load
            Sleep 500

            ; Search for Office button (e.g., "Office Google Cast")
            officePattern := "i)Office"
            if (officeBtn := WaitForButton(spot, officePattern, 3000)) {
                officeBtn.Invoke()
            }
            ; If Office button not found, continue without error (as requested)
        }
        else {
            ; #region agent log
            SafeDebugLog Format(
                "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:7595`",`"message`":`"Button not found by WaitForButton`",`"data`":{`"pattern`":`"{4}`"},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"B,D`"}`n",
                A_TickCount, Random(1000, 9999), A_TickCount, connectPattern)
            ; #endregion
            MsgBox "Couldn't find the Connect-to-device button."
        }
    } catch Error as e {
        ; #region agent log
        SafeDebugLog Format(
            "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:7597`",`"message`":`"Exception caught`",`"data`":{`"error`":`"{4}`"},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"A,B,C,D,E`"}`n",
            A_TickCount, Random(1000, 9999), A_TickCount, e.Message)
        ; #endregion
        MsgBox "Error: " e.Message
    }
    ; #region agent log
    SafeDebugLog Format(
        "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:7600`",`"message`":`"Connect button handler exit`",`"data`":{},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"A`"}`n",
        A_TickCount, Random(1000, 9999), A_TickCount)
    ; #endregion
}

; Shift + F : Toggle full screen - Fullscreen
+f:: SpotifyToggleFullscreen()

; Shift + S : Open Search - Search
+s:: Send "^k"

; Shift + P : Go to Playlists - Playlists
+p:: Send "!+1"

; Shift + A : Go to Artists - Artists
+a:: Send "!+3"

; Shift + B : Go to Albums - Albums
+b:: Send "!+4"

; Shift + H : Go to Home - Home
+h:: Send "!+h"

; Shift + N : Go to Now Playing - Now Playing
+n:: Send "!+j"

; Shift + M : Go to Made For You - Made For You
+m:: Send "!+m"

; Shift + R : Go to New Releases - Releases
+r:: Send "!+n"

; Shift + X : Go to Charts - Charts
+x:: Send "!+c"

; Shift + V : Toggle Now Playing View Sidebar - View
+v:: Send "!+r"

; Shift + L : Toggle Your Library Sidebar - Library
+l:: Send "!+l"

; Shift + E : Toggle Fullscreen Library - Expand Library
+e::
{
    try {
        spot := UIA_Browser("ahk_exe Spotify.exe")
        Sleep 300

        ; First, try to find and click "Open Your Library" button (if available)
        try {
            openLibBtn := spot.FindElement({ Name: "Open Your Library", Type: "Button" })
            if (openLibBtn) {
                openLibBtn.Click()
                Sleep 500  ; Wait for the library to open and UI to adjust
            }
        } catch {
            ; "Open Your Library" button not found - this is normal, continue to next step
        }

        ; Then, try to find and click "Expand Your Library" button
        try {
            expandLibBtn := spot.FindElement({ Name: "Expand Your Library", Type: "Button" })
            if (expandLibBtn) {
                expandLibBtn.Click()
                Sleep 300  ; Wait for the expansion to complete
            } else {
                MsgBox "Could not find the 'Expand Your Library' button.", "Spotify Navigation", "IconX"
            }
        } catch {
            MsgBox "Could not find the 'Expand Your Library' button.", "Spotify Navigation", "IconX"
        }
    } catch Error as e {
        MsgBox "Error toggling fullscreen library: " e.Message, "Spotify Error", "IconX"
    }
}

; Shift + I : Immerse — topmost header Play in main entity region, then fullscreen
+i:: {
    try {
        StandardLoadingBar_Show("⏳ Immersing — finding header Play…", BANNER_ACCENT_INTERMEDIATE, {
            passive: false, centerOnHwnd: 0, fontSize: 17 })
        spot := UIA_Browser("ahk_exe Spotify.exe")
        Sleep(200)
        if (!spot) {
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ Spotify UIA root not found", 2000, BANNER_ACCENT_ERROR)
            return
        }
        StandardLoadingBar_Update("⏳ Immersing — locating header Play…", BANNER_ACCENT_INTERMEDIATE)
        playBtn := FindHeaderPlayButton(spot)
        if (!playBtn) {
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ Header Play button not found", 2000, BANNER_ACCENT_ERROR)
            return
        }
        ; Already playing this context (header shows Pause) — skip click so we don't pause.
        if (!IsHeaderPauseLike(playBtn)) {
            StandardLoadingBar_Update("⏳ Immersing — activating Play…", BANNER_ACCENT_INTERMEDIATE)
            activated := ActivateElement(playBtn)
            if (!activated) {
                StandardLoadingBar_Hide(0)
                ShowCenteredOverlay_Utils("❌ Failed to activate header Play", 2000, BANNER_ACCENT_ERROR)
                return
            }
            Sleep(250)
        }
        StandardLoadingBar_Update("⏳ Immersing — entering fullscreen…", BANNER_ACCENT_INTERMEDIATE)
        fsOk := SpotifyToggleFullscreen()
        StandardLoadingBar_Hide(0)
        if (fsOk) {
            ShowCenteredOverlay_Utils("✅ Immersed", 1200, BANNER_ACCENT_SUCCESS)
            Send("!{Tab}")
        } else
            ShowCenteredOverlay_Utils("⚠ Play started; fullscreen control not found", 2000, BANNER_ACCENT_INTERMEDIATE)
    } catch Error as e {
        try StandardLoadingBar_Hide(0)
        ShowCenteredOverlay_Utils("❌ Immerse error: " e.Message, 2500, BANNER_ACCENT_ERROR)
    }
}

; Shift + Y : Toggle lyrics - Lyrics
+y:: Send("^+")

; Shift + T : Toggle Play/Pause - Play/Pause
; Improvements:
; - Robust word-based detection: matches any Button 50000 whose Name contains the word "play"
;   (also supports "reproduzir" / "tocar"), prefers Play over Pause when both are seen.
; - No click on the anchor. Only SetFocus/Select.
+t:: {
    try {
        spot := UIA_Browser("ahk_exe Spotify.exe")
        Sleep(200)
        if (!spot) {
            Send("{Media_Play_Pause}")
            return
        }
        if (!WinExist("ahk_exe Spotify.exe")) {
            ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        WinActivate("ahk_exe Spotify.exe")
        Sleep(150)

        ; 1) Find the exact anchor: Button 50000 named "Download"
        anchor := FindDownloadAnchor(spot)
        if (!anchor) {
            ; Fallback: best-scored Play/Pause button via scan
            btn := FindBestPlayPauseButton(spot)
            if (btn) {
                ActivateElement(btn)
                return
            }
            Send("{Media_Play_Pause}")
            return
        }

        ; 2) Focus the anchor WITHOUT clicking it
        if (!FocusAnchorWithoutClick(anchor)) {
            btn := FindBestPlayPauseButton(spot)
            if (btn) {
                ActivateElement(btn)
                return
            }
            Send("{Media_Play_Pause}")
            return
        }
        Sleep(160)

        ; 3) From the anchor, back-tab up to 12 steps to locate Play/Pause
        if (HuntBackToPlayPausePreferPlay(12))  ; Shift+Tab only
            return

        ; 4) Direct lookup fallback - scan and pick the best-scoring Play/Pause button
        btn := FindBestPlayPauseButton(spot)
        if (btn) {
            ActivateElement(btn)
            return
        }

        ; 5) Final fallback – OS media key
        Send("{Media_Play_Pause}")

    } catch Error as e {
        Send("{Media_Play_Pause}")
    }
}

; ---------------------------
; Helpers
; ---------------------------

FindDownloadAnchor(root) {
    ; Exact spec: Type 50000 (Button), Name "Download"
    try {
        el := root.FindElement({ Type: 50000, Name: "Download" })
        if (el)
            return el
    } catch {
    }
    try {
        el := root.FindElement({ Type: "Button", Name: "Download" })
        if (el)
            return el
    } catch {
    }
    return ""
}

FocusAnchorWithoutClick(el) {
    ; Do NOT click the anchor
    try {
        el.SetFocus()
        return true
    } catch {
    }
    try {
        el.Select()   ; Safe, non-click focus in many UIA wrappers
        return true
    } catch {
    }
    return false
}

HuntBackToPlayPausePreferPlay(steps) {
    global UIA
    ; Prefer Play (> Pause). Keep the first Pause seen as a fallback.
    pauseCandidate := ""

    ; Check current focus first
    try {
        fe := UIA.GetFocusedElement()
        sc := PlayPauseScore(fe)
        if (sc >= 2)
            return ActivateElement(fe)  ; Found Play (or equivalent)
        else if (sc = 1)
            pauseCandidate := fe
    } catch {
    }

    loop steps {
        Send("+{Tab}")            ; Shift+Tab only
        Sleep(80)
        try {
            fe := UIA.GetFocusedElement()
            sc := PlayPauseScore(fe)
            if (sc >= 2)
                return ActivateElement(fe)  ; Prefer Play immediately
            else if (!pauseCandidate && sc = 1)
                pauseCandidate := fe        ; Remember first Pause
        } catch {
        }
    }
    if (pauseCandidate)
        return ActivateElement(pauseCandidate)
    return false
}

; Score-based detector:
; 2 = Play-like (play/reproduzir/tocar)
; 1 = Pause-like (pause/pausar/pausa)
; 0 = not a target
PlayPauseScore(el) {
    try {
        tp := el.Type
        if !(tp == 50000 || tp == "Button")
            return 0

        nm := el.Name
        if (!nm)
            return 0

        norm := NormalizeName(nm)

        if (ContainsWord(norm, "play") || ContainsWord(norm, "reproduzir") || ContainsWord(norm, "tocar"))
            return 2
        if (ContainsWord(norm, "pause") || ContainsWord(norm, "pausar") || ContainsWord(norm, "pausa"))
            return 1
    } catch {
    }
    return 0
}

ActivateElement(el) {
    try {
        el.Invoke()     ; Preferred - UIA Invoke pattern
        return true
    } catch {
    }
    try {
        el.Click()      ; Acceptable on the target (not the anchor)
        return true
    } catch {
    }
    Send("{Enter}")     ; Last resort
    Sleep(60)
    return true
}

; Scan all buttons and pick the best-scoring Play/Pause control
FindBestPlayPauseButton(root) {
    best := ""
    bestScore := 0

    ; First try numeric ControlType
    try {
        btns := root.FindAll({ Type: 50000 })
        if (btns && btns.Length) {
            for _, b in btns {
                sc := PlayPauseScore(b)
                if (sc > bestScore) {
                    best := b, bestScore := sc
                    if (bestScore >= 2)  ; Play found - early exit
                        return best
                }
            }
        }
    } catch {
    }

    ; Then try textual ControlType
    try {
        btns := root.FindAll({ Type: "Button" })
        if (btns && btns.Length) {
            for _, b in btns {
                sc := PlayPauseScore(b)
                if (sc > bestScore) {
                    best := b, bestScore := sc
                    if (bestScore >= 2)
                        return best
                }
            }
        }
    } catch {
    }

    return best
}

; --- header Play immersion (Shift+I) ---

; Shared fullscreen toggle used by +f and +i (Send("+f") does not reliably re-enter +f).
SpotifyToggleFullscreen() {
    try {
        spot := UIA_Browser("ahk_exe Spotify.exe")
        Sleep 300
        enterFsPattern := "i)^Enter Full[- ]?screen$"
        exitFsPattern := "i)^Exit Full[- ]?screen$"

        ClickFsButton(btn) {
            if (!btn)
                return false
            supportsInvoke := false
            try supportsInvoke := btn.GetPropertyValue(UIA.Property.IsInvokePatternAvailable)
            catch
                supportsInvoke := false
            if (supportsInvoke) {
                try {
                    btn.Invoke()
                    return true
                } catch {
                }
            }
            try {
                btn.Click()
                return true
            } catch {
            }
            return false
        }

        enterFsBtn := WaitForButton(spot, enterFsPattern, 500)
        if (enterFsBtn) {
            clickedEnter := ClickFsButton(enterFsBtn)
            Sleep 300
            Send "{Escape}"
            return clickedEnter
        }

        exitFsBtn := WaitForButton(spot, exitFsPattern, 500)
        if (!exitFsBtn) {
            Sleep 1000
            exitFsBtn := WaitForButton(spot, exitFsPattern, 500)
        }
        if (exitFsBtn)
            return ClickFsButton(exitFsBtn)
        return false
    } catch {
        return false
    }
}

; Header Play/Pause name: bare or entity-qualified (EN + PT).
IsHeaderPlayPauseName(nm) {
    if (!nm)
        return false
    return !!RegExMatch(nm, "i)^(Play|Pause|Reproduzir|Pausar)(\s|$)")
}

IsHeaderPauseLike(el) {
    try {
        nm := el.Name
        if (!nm)
            return false
        return !!RegExMatch(nm, "i)^(Pause|Pausar)(\s|$)")
    } catch {
        return false
    }
}

IsHeaderPlayPausePrimary(el) {
    try {
        nm := el.Name
        if (!IsHeaderPlayPauseName(nm))
            return false
        cn := ""
        try cn := el.ClassName
        if (cn != "" && !InStr(cn, "legacy-button-primary"))
            return false
        return true
    } catch {
        return false
    }
}

; Visible on-screen element with a non-empty bounding rectangle.
IsVisibleUiaElement(el) {
    if (!el)
        return false
    try {
        try {
            if (el.GetPropertyValue(UIA.Property.IsOffscreen))
                return false
        } catch {
        }
        br := el.BoundingRectangle
        if (!IsObject(br))
            return false
        if ((br.r - br.l) <= 0 || (br.b - br.t) <= 0)
            return false
        return true
    } catch {
        return false
    }
}

; Primary Play/Pause buttons (entity-qualified or bare). Excludes track-row plays (no primary class).
CollectHeaderPlayCandidates(root) {
    candidates := []
    if (!root)
        return candidates
    try {
        btns := root.FindAll({ Type: 50000 })
        if (btns && btns.Length) {
            for b in btns {
                try {
                    if (IsHeaderPlayPausePrimary(b))
                        candidates.Push(b)
                } catch {
                    continue
                }
            }
        }
    } catch {
    }
    return candidates
}

; Main content entity group, e.g. "best cartoon songs - playlist by DON | Spotify".
FindMainEntityGroup(root) {
    if (!root)
        return ""
    try {
        groups := root.FindAll({ Type: 50026 })
        if (groups && groups.Length) {
            for g in groups {
                try {
                    nm := g.Name
                    if (nm && RegExMatch(nm, "i)\| Spotify$"))
                        return g
                } catch {
                    continue
                }
            }
            ; Fallback: LocalizedType "principal" / "main" when Name is missing.
            for g in groups {
                try {
                    lt := ""
                    try lt := g.LocalizedType
                    if (lt && RegExMatch(lt, "i)^(principal|main)$"))
                        return g
                } catch {
                    continue
                }
            }
        }
    } catch {
    }
    return ""
}

; Topmost visible primary Play/Pause inside the entity action area.
; Action-bar control sits below metadata and above track/recommendation Play buttons.
FindHeaderPlayInGroup(group) {
    if (!group)
        return ""
    candidates := CollectHeaderPlayCandidates(group)
    if (!candidates.Length)
        return ""
    best := ""
    bestT := 0.0
    bestL := 0.0
    for c in candidates {
        try {
            if (!IsVisibleUiaElement(c))
                continue
            br := c.BoundingRectangle
            if (best = "" || br.t < bestT || (br.t = bestT && br.l < bestL)) {
                best := c
                bestT := br.t
                bestL := br.l
            }
        } catch {
            continue
        }
    }
    return best
}

; Main entity region only — no app-wide fallback (avoids related-card / player hits).
FindHeaderPlayButton(root) {
    if (!root)
        return ""
    mainGroup := FindMainEntityGroup(root)
    if (!mainGroup)
        return ""
    return FindHeaderPlayInGroup(mainGroup)
}

; --- text utils ---

NormalizeName(s) {
    ; Lowercase and collapse non-word chars (punctuation, hyphens) to single spaces
    try s := StrLower(s)
    catch {
    }
    try s := RegExReplace(s, "[^\w]+", " ")
    catch {
    }
    return Trim(s)
}

ContainsWord(norm, word) {
    ; Match a whole word boundary after normalization
    return RegExMatch(norm, "(^|\s)" . word . "(\s|$)")
}
