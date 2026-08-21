; =============================================================================
; Utils module: handy_retranscribe_last.ahk
; Send dictation? [B] — toggle Parakeet/Cohere, History re-transcribe, copy
; Reuses HandyReplay_* History/click helpers from handy_replay_last.ahk
; =============================================================================

global HANDY_RETRANSCRIBE_MAX_WAIT_MS := 90000
global HANDY_RETRANSCRIBE_POLL_MS := 250

HandyRetranscribe_FindFirstRetranscribe(el) {
    return HandyReplay_FindNamed(el, 50000, ["Re-transcribe", "Retranscrever"])
}

HandyRetranscribe_FindFirstCopy(el) {
    return HandyReplay_FindNamed(el, 50000, [
        "Copy transcription to clipboard",
        "Copiar transcrição para a área de transferência"
    ])
}

HandyRetranscribe_IsNoiseHistoryText(name) {
    if (name = "" || name = "HISTORY" || name = "Histórico" || name = "•" || name = "v")
        return true
    low := StrLower(name)
    if (InStr(low, "transcribing") || InStr(low, "transcrevendo"))
        return true
    switch name {
        case "General", "History", "Histórico", "Models", "Modelos", "Advanced", "Avançado",
            "Post Process", "Pós-processamento", "About", "Sobre", "Open Recordings Folder",
            "Abrir pasta de gravações":
            return true
    }
    if RegExMatch(name, "^\d+:\d+$")
        return true
    if InStr(name, " at ")
        return true
    if InStr(name, " às ")
        return true
    return false
}

; Newest History entry transcription: first body Text after the newest entry's date header.
HandyRetranscribe_ReadFirstEntryText(el) {
    if !el
        return ""
    try {
        sawEntryChrome := false
        for txt in el.FindAll({ Type: 50020 }) {
            nm := ""
            try nm := txt.Name
            if (nm = "")
                continue
            if (!sawEntryChrome) {
                if InStr(nm, " at ") || InStr(nm, " às ") {
                    sawEntryChrome := true
                    continue
                }
                continue
            }
            if HandyRetranscribe_IsNoiseHistoryText(nm)
                continue
            return nm
        }
    } catch {
    }
    return ""
}

HandyRetranscribe_ButtonEnabled(btn) {
    if !btn
        return false
    try {
        if !btn.GetPropertyValue(UIA.Property.IsEnabled)
            return false
    } catch {
        try {
            if !btn.IsEnabled
                return false
        } catch {
        }
    }
    return true
}

HandyRetranscribe_WaitFirstRetranscribe(hwnd, maxWaitMs := 2000) {
    global UIA
    start := A_TickCount
    loop {
        try {
            el := UIA.ElementFromHandle(hwnd)
            btn := HandyRetranscribe_FindFirstRetranscribe(el)
            if btn
                return btn
        } catch {
        }
        if ((A_TickCount - start) >= maxWaitMs)
            break
        Sleep 100
    }
    return 0
}

HandyRetranscribe_ClickFirstRetranscribe(hwnd) {
    btn := HandyRetranscribe_WaitFirstRetranscribe(hwnd, 2000)
    if !btn
        return false
    if HandyReplay_ClickWithStrategies(btn, hwnd)
        return true
    try btn.Invoke()
    catch {
        return false
    }
    return true
}

HandyRetranscribe_ClickFirstCopy(hwnd) {
    global UIA
    start := A_TickCount
    btn := 0
    loop {
        try {
            el := UIA.ElementFromHandle(hwnd)
            btn := HandyRetranscribe_FindFirstCopy(el)
            if (btn && HandyRetranscribe_ButtonEnabled(btn))
                break
            btn := 0
        } catch {
            btn := 0
        }
        if ((A_TickCount - start) >= 5000)
            break
        Sleep 150
    }
    if !btn
        return false
    if HandyReplay_ClickWithStrategies(btn, hwnd)
        return true
    try btn.Invoke()
    catch {
        return false
    }
    return true
}

; Busy = visible Transcribing label only. Ignore hidden/offscreen leftover nodes.
HandyRetranscribe_ElementIsOnScreen(el) {
    if !el
        return false
    try {
        if el.GetPropertyValue(UIA.Property.IsOffscreen)
            return false
    } catch {
        try {
            if el.IsOffscreen
                return false
        } catch {
        }
    }
    try {
        pos := el.Location
        if (pos.w <= 0 || pos.h <= 0)
            return false
    } catch {
    }
    return true
}

HandyRetranscribe_IsTranscribingUiVisible(el) {
    if !el
        return false
    needles := ["Transcribing", "Transcrevendo"]
    for nm in needles {
        try {
            found := el.FindFirst({ Type: 50020, Name: nm })
            if (found && HandyRetranscribe_ElementIsOnScreen(found))
                return true
        } catch {
        }
    }
    for nm in needles {
        try {
            found := el.FindFirst({ Type: 50020, Name: nm, matchmode: "Substring", cs: false })
            if !found
                continue
            n := ""
            try n := found.Name
            low := StrLower(n)
            if (InStr(low, "re-transcribe") || InStr(low, "transcription") || InStr(low, "retranscrever"))
                continue
            if !(InStr(low, "transcribing") || InStr(low, "transcrevendo"))
                continue
            if !HandyRetranscribe_ElementIsOnScreen(found)
                continue
            return true
        } catch {
        }
    }
    return false
}

; Wait until Transcribing appears then clears, and entry text has changed from prior mistranslation.
HandyRetranscribe_WaitRetranscribeDone(hwnd, priorText := "", maxWaitMs := 0) {
    global UIA, HANDY_RETRANSCRIBE_MAX_WAIT_MS, HANDY_RETRANSCRIBE_POLL_MS
    if (maxWaitMs <= 0)
        maxWaitMs := HANDY_RETRANSCRIBE_MAX_WAIT_MS
    start := A_TickCount
    sawBusy := false
    clearStreak := 0
    requiredClearPolls := 4
    textStableStreak := 0
    lastSeenText := ""

    loop {
        try {
            el := UIA.ElementFromHandle(hwnd)
            busy := HandyRetranscribe_IsTranscribingUiVisible(el)
            curText := HandyRetranscribe_ReadFirstEntryText(el)
            textChanged := (priorText != "" && curText != "" && curText != priorText)
            || (priorText = "" && curText != "")

            if (busy) {
                sawBusy := true
                clearStreak := 0
                textStableStreak := 0
            } else if (sawBusy) {
                clearStreak += 1
                ; After busy clears, require the History entry text to differ from the old one
                ; and stay stable briefly (UI can flash empty then fill).
                if (clearStreak >= requiredClearPolls && textChanged) {
                    if (curText = lastSeenText)
                        textStableStreak += 1
                    else {
                        lastSeenText := curText
                        textStableStreak := 1
                    }
                    if (textStableStreak >= 3)
                        return true
                } else {
                    textStableStreak := 0
                }
            } else if (!busy && textChanged) {
                ; Missed Transcribing label — still require changed text to stay stable ~1s.
                if (curText = lastSeenText)
                    textStableStreak += 1
                else {
                    lastSeenText := curText
                    textStableStreak := 1
                }
                if (textStableStreak >= 6)
                    return true
            } else {
                textStableStreak := 0
            }
        } catch {
        }
        if ((A_TickCount - start) >= maxWaitMs)
            break
        Sleep HANDY_RETRANSCRIBE_POLL_MS
    }
    return false
}

; History → Re-transcribe newest → Copy to clipboard. Assumes Handy already open on any tab.
HandyRetranscribe_HistoryRetranscribeAndCopy(hwnd) {
    global UIA
    if !hwnd
        return false

    if (!HandyReplay_EnsureHistoryTab(hwnd)) {
        ShowCenteredOverlay_Utils("❌ Handy History tab not found.", 2200, BANNER_ACCENT_ERROR)
        return false
    }

    priorText := ""
    try {
        el := UIA.ElementFromHandle(hwnd)
        priorText := HandyRetranscribe_ReadFirstEntryText(el)
    } catch {
    }
    ; Prefer OS clipboard as the mistranslation baseline when History text is hard to read.
    clipBefore := ""
    try clipBefore := A_Clipboard
    if (priorText = "" && clipBefore != "")
        priorText := clipBefore

    if (!HandyRetranscribe_ClickFirstRetranscribe(hwnd)) {
        ShowCenteredOverlay_Utils("❌ Could not click Re-transcribe.", 2200, BANNER_ACCENT_ERROR)
        return false
    }

    StandardLoadingBar_Show("⏳ Re-transcribing — please wait...", BANNER_ACCENT_INTERMEDIATE, {
        fontSize: 17,
        trackActiveMonitor: true
    })
    barOwned := true
    try {
        if (!HandyRetranscribe_WaitRetranscribeDone(hwnd, priorText)) {
            StandardLoadingBar_Hide(0)
            barOwned := false
            ShowCenteredOverlay_Utils("❌ Re-transcribe timed out.", 2200, BANNER_ACCENT_ERROR)
            return false
        }

        StandardLoadingBar_Update("⏳ Copying transcription...")
        StandardLoadingBar_Hide(0)
        barOwned := false

        if (!HandyReplay_EnsureHistoryTab(hwnd)) {
            ShowCenteredOverlay_Utils("❌ Handy History tab not found after re-transcribe.", 2200, BANNER_ACCENT_ERROR)
            return false
        }
        Sleep 200

        newText := ""
        try {
            el := UIA.ElementFromHandle(hwnd)
            newText := HandyRetranscribe_ReadFirstEntryText(el)
        } catch {
        }
        if (newText = "" || (priorText != "" && newText = priorText)) {
            ShowCenteredOverlay_Utils("❌ Transcription did not update.", 2200, BANNER_ACCENT_ERROR)
            return false
        }

        if (!HandyRetranscribe_ClickFirstCopy(hwnd)) {
            ShowCenteredOverlay_Utils("❌ Could not copy transcription.", 2200, BANNER_ACCENT_ERROR)
            return false
        }

        ; ClipWait alone is wrong: old dictation text is already on the clipboard.
        clipOk := false
        waitEnd := A_TickCount + 4000
        loop {
            try {
                if (A_Clipboard != "" && A_Clipboard != clipBefore
                    && (priorText = "" || A_Clipboard != priorText)) {
                    clipOk := true
                    break
                }
            } catch {
            }
            if ((A_TickCount) >= waitEnd)
                break
            Sleep 100
        }
        ; Prefer Handy's Copy; if it left the old clip, push the verified History text.
        if (!clipOk && newText != "" && newText != priorText) {
            try {
                A_Clipboard := newText
                clipOk := (A_Clipboard = newText)
            } catch {
                clipOk := false
            }
        }
        if !clipOk {
            ShowCenteredOverlay_Utils("❌ Clipboard still has old transcription.", 2200, BANNER_ACCENT_ERROR)
            return false
        }

        ShowCenteredOverlay_Utils("✅ Re-transcribed & copied", 1500, BANNER_ACCENT_SUCCESS)
        return true
    } finally {
        if (barOwned) {
            try StandardLoadingBar_Hide(0)
        }
    }
}

; End-to-end for Send dictation? [B]: toggle model (keep open) → re-transcribe → copy → close Handy.
HandyRetranscribe_ToggleModelAndCopy() {
    targetSlot := Handy_GetDictationToggleTargetSlot()

    if (!ExecuteHandyAiModelSelection(targetSlot, true)) {
        ShowCenteredOverlay_Utils("❌ Could not switch Handy model.", 2200, BANNER_ACCENT_ERROR)
        return false
    }

    hwnd := Handy_ActivateOrLaunch()
    if (!hwnd) {
        ShowCenteredOverlay_Utils("❌ Could not open Handy.", 2000, BANNER_ACCENT_ERROR)
        return false
    }

    ok := HandyRetranscribe_HistoryRetranscribeAndCopy(hwnd)
    if (ok) {
        try WinClose("ahk_id " . hwnd)
    }
    return ok
}
