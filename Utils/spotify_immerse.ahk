; =============================================================================
; Utils module: spotify_immerse.ahk
; Shared Spotify Immersion (header Play + fullscreen) for Shift+I and D2C [P]
; =============================================================================

SpotifyImmerse_ActivateElement(el) {
    try {
        el.Invoke()
        return true
    } catch {
    }
    try {
        el.Click()
        return true
    } catch {
    }
    Send("{Enter}")
    Sleep(60)
    return true
}

; Slim WaitForButton for fullscreen controls (no Shift-keys WaitForButton dependency).
SpotifyImmerse_WaitForButton(root, pattern, timeout := 5000) {
    if !IsObject(root)
        return 0
    deadline := A_TickCount + timeout
    while (A_TickCount < deadline) {
        try {
            buttons := root.FindAll({ Type: "Button" })
            if (buttons && buttons.Length) {
                for btn in buttons {
                    try {
                        nm := btn.Name
                        if (nm && RegExMatch(nm, pattern))
                            return btn
                    } catch {
                        continue
                    }
                }
            }
        } catch {
        }
        Sleep 50
    }
    return 0
}

SpotifyImmerse_ClickFsButton(btn) {
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

; Shared fullscreen toggle used by +f and Immersion (Send("+f") does not reliably re-enter +f).
SpotifyToggleFullscreen() {
    try {
        spot := UIA_Browser("ahk_exe Spotify.exe")
        Sleep 300
        enterFsPattern := "i)^Enter Full[- ]?screen$"
        exitFsPattern := "i)^Exit Full[- ]?screen$"

        enterFsBtn := SpotifyImmerse_WaitForButton(spot, enterFsPattern, 500)
        if (enterFsBtn) {
            clickedEnter := SpotifyImmerse_ClickFsButton(enterFsBtn)
            Sleep 300
            Send "{Escape}"
            return clickedEnter
        }

        exitFsBtn := SpotifyImmerse_WaitForButton(spot, exitFsPattern, 500)
        if (!exitFsBtn) {
            Sleep 1000
            exitFsBtn := SpotifyImmerse_WaitForButton(spot, exitFsPattern, 500)
        }
        if (exitFsBtn)
            return SpotifyImmerse_ClickFsButton(exitFsBtn)
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

FindHeaderPlayButton(root) {
    if (!root)
        return ""
    mainGroup := FindMainEntityGroup(root)
    if (!mainGroup)
        return ""
    return FindHeaderPlayInGroup(mainGroup)
}

SpotifyImmerse_UpdateLoading(state, ownLoading) {
    global g_StandardLoadingBarGui
    if (ownLoading || IsObject(g_StandardLoadingBarGui))
        StandardLoadingBar_Update(state, BANNER_ACCENT_INTERMEDIATE)
}

; ownLoading true (Shift+I): manage Show/Hide. false (D2C [P]): Update only; caller owns the bar.
SpotifyImmerse(ownLoading := true) {
    try {
        if (ownLoading) {
            StandardLoadingBar_Show("⏳ Immersing — finding header Play…", BANNER_ACCENT_INTERMEDIATE, {
                passive: false, centerOnHwnd: 0, fontSize: 17, trackActiveMonitor: true })
        } else {
            SpotifyImmerse_UpdateLoading("⏳ Immersing — finding header Play…", ownLoading)
        }

        spot := UIA_Browser("ahk_exe Spotify.exe")
        Sleep(200)
        if (!spot) {
            if (ownLoading)
                StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ Spotify UIA root not found", 2000, BANNER_ACCENT_ERROR)
            return false
        }

        SpotifyImmerse_UpdateLoading("⏳ Immersing — locating header Play…", ownLoading)
        playBtn := FindHeaderPlayButton(spot)
        if (!playBtn) {
            if (ownLoading)
                StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ Header Play button not found", 2000, BANNER_ACCENT_ERROR)
            return false
        }

        ; Already playing this context (header shows Pause) — skip click so we don't pause.
        if (!IsHeaderPauseLike(playBtn)) {
            SpotifyImmerse_UpdateLoading("⏳ Immersing — activating Play…", ownLoading)
            activated := SpotifyImmerse_ActivateElement(playBtn)
            if (!activated) {
                if (ownLoading)
                    StandardLoadingBar_Hide(0)
                ShowCenteredOverlay_Utils("❌ Failed to activate header Play", 2000, BANNER_ACCENT_ERROR)
                return false
            }
            Sleep(250)
        }

        SpotifyImmerse_UpdateLoading("⏳ Immersing — entering fullscreen…", ownLoading)
        fsOk := SpotifyToggleFullscreen()
        if (ownLoading)
            StandardLoadingBar_Hide(0)

        if (fsOk) {
            ShowCenteredOverlay_Utils("✅ Immersed", 1200, BANNER_ACCENT_SUCCESS)
            Send("!{Tab}")
            return true
        }
        ShowCenteredOverlay_Utils("⚠ Play started; fullscreen control not found", 2000, BANNER_ACCENT_INTERMEDIATE)
        return false
    } catch Error as e {
        if (ownLoading) {
            try StandardLoadingBar_Hide(0)
        }
        ShowCenteredOverlay_Utils("❌ Immerse error: " e.Message, 2500, BANNER_ACCENT_ERROR)
        return false
    }
}
