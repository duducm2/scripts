; =============================================================================
; Utils module: handy_retranscribe_last.ahk
; Send dictation? [B] — toggle Parakeet/Cohere, History re-transcribe, copy
; Reuses HandyReplay_* History/click helpers from handy_replay_last.ahk
; =============================================================================

global HANDY_RETRANSCRIBE_MAX_WAIT_MS := 90000
global HANDY_RETRANSCRIBE_POLL_MS := 250
global g_HandyRetranscribeLastMatchLog := 0

; #region agent log
AgentDebugLog(hypothesisId, location, message, dataMap := "") {
    logPath := A_ScriptDir "\debug-f7e6c2.log"
    dataJson := "{}"
    if (IsObject(dataMap)) {
        parts := []
        for k, v in dataMap {
            vv := v
            if (Type(vv) = "String") {
                vv := StrReplace(vv, "\", "\\")
                vv := StrReplace(vv, '"', '\"')
                vv := '"' vv '"'
            } else if (Type(vv) = "Integer" || Type(vv) = "Float") {
                ; keep numeric
            } else if (vv = true) {
                vv := "true"
            } else if (vv = false) {
                vv := "false"
            } else {
                vv := '"' StrReplace(String(vv), '"', '\"') '"'
            }
            parts.Push('"' k '":' vv)
        }
        dataJson := "{" . ArrJoin(parts, ",") . "}"
    }
    line := '{"sessionId":"f7e6c2","hypothesisId":"' hypothesisId '","location":"' location '","message":"' message '","data":' dataJson ',"timestamp":' A_TickCount ',"runId":"post-fix"}`n'
    try FileAppend(line, logPath, "UTF-8")
}

ArrJoin(arr, sep) {
    out := ""
    for i, s in arr
        out .= (i = 1 ? "" : sep) . s
    return out
}
; #endregion

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
    ; #region agent log
    AgentDebugLog("L", "handy_retranscribe_last.ahk:ClickFirstCopy", "found", Map(
        "found", btn ? 1 : 0,
        "waited", A_TickCount - start
    ))
    ; #endregion
    if !btn
        return false
    if HandyReplay_ClickWithStrategies(btn, hwnd) {
        ; #region agent log
        AgentDebugLog("L", "handy_retranscribe_last.ahk:ClickFirstCopy", "strategies_ok", Map())
        ; #endregion
        return true
    }
    try {
        btn.Invoke()
        ; #region agent log
        AgentDebugLog("L", "handy_retranscribe_last.ahk:ClickFirstCopy", "invoke_ok", Map())
        ; #endregion
        return true
    } catch {
        ; #region agent log
        AgentDebugLog("L", "handy_retranscribe_last.ahk:ClickFirstCopy", "all_fail", Map())
        ; #endregion
        return false
    }
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
    global g_HandyRetranscribeLastMatchLog
    if !el
        return false
    ; Exact names first, then substring — but reject Re-transcribe / transcription false positives.
    needles := ["Transcribing", "Transcrevendo"]
    for nm in needles {
        try {
            found := el.FindFirst({ Type: 50020, Name: nm })
            if (found && HandyRetranscribe_ElementIsOnScreen(found)) {
                ; #region agent log
                if ((A_TickCount - g_HandyRetranscribeLastMatchLog) > 2500) {
                    g_HandyRetranscribeLastMatchLog := A_TickCount
                    mn := ""
                    try mn := found.Name
                    AgentDebugLog("S", "handy_retranscribe_last.ahk:IsTranscribingUiVisible", "match_exact_text", Map(
                        "name", mn
                    ))
                }
                ; #endregion
                return true
            }
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
            ; #region agent log
            if ((A_TickCount - g_HandyRetranscribeLastMatchLog) > 2500) {
                g_HandyRetranscribeLastMatchLog := A_TickCount
                AgentDebugLog("S", "handy_retranscribe_last.ahk:IsTranscribingUiVisible", "match_substr_text", Map(
                    "name", n
                ))
            }
            ; #endregion
            return true
        } catch {
        }
    }
    return false
}

; Wait until Transcribing appears then clears; then caller clicks Copy. Do not gate on entry text.
HandyRetranscribe_WaitRetranscribeDone(hwnd, priorText, maxWaitMs := 0) {
    global UIA, HANDY_RETRANSCRIBE_MAX_WAIT_MS, HANDY_RETRANSCRIBE_POLL_MS
    if (maxWaitMs <= 0)
        maxWaitMs := HANDY_RETRANSCRIBE_MAX_WAIT_MS
    start := A_TickCount
    sawBusy := false
    clearStreak := 0
    requiredClearPolls := 4
    appearDeadline := start + 12000
    lastBeat := 0

    ; #region agent log
    AgentDebugLog("F", "handy_retranscribe_last.ahk:WaitRetranscribeDone", "start", Map(
        "priorLen", StrLen(priorText),
        "maxWaitMs", maxWaitMs
    ))
    ; #endregion

    loop {
        try {
            el := UIA.ElementFromHandle(hwnd)
            busy := HandyRetranscribe_IsTranscribingUiVisible(el)
            copyBtn := HandyRetranscribe_FindFirstCopy(el)
            copyEn := (copyBtn && HandyRetranscribe_ButtonEnabled(copyBtn)) ? 1 : 0

            ; #region agent log
            if ((A_TickCount - lastBeat) >= 3000) {
                lastBeat := A_TickCount
                AgentDebugLog("N", "handy_retranscribe_last.ahk:WaitRetranscribeDone", "heartbeat", Map(
                    "elapsed", A_TickCount - start,
                    "busy", busy ? 1 : 0,
                    "sawBusy", sawBusy ? 1 : 0,
                    "clearStreak", clearStreak,
                    "copyEn", copyEn
                ))
            }
            ; #endregion

            if (busy) {
                if !sawBusy {
                    ; #region agent log
                    AgentDebugLog("F", "handy_retranscribe_last.ahk:WaitRetranscribeDone", "saw_busy", Map(
                        "elapsed", A_TickCount - start
                    ))
                    ; #endregion
                }
                sawBusy := true
                clearStreak := 0
            } else if (sawBusy) {
                clearStreak += 1
                if (clearStreak >= requiredClearPolls) {
                    ; #region agent log
                    AgentDebugLog("F", "handy_retranscribe_last.ahk:WaitRetranscribeDone", "done_after_busy_clear", Map(
                        "elapsed", A_TickCount - start,
                        "copyEn", copyEn
                    ))
                    ; #endregion
                    return true
                }
            } else if (!sawBusy && (A_TickCount >= appearDeadline) && copyEn) {
                ; #region agent log
                AgentDebugLog("H", "handy_retranscribe_last.ahk:WaitRetranscribeDone", "done_no_busy_copy_ready", Map(
                    "elapsed", A_TickCount - start
                ))
                ; #endregion
                return true
            }
        } catch as err {
            ; #region agent log
            if ((A_TickCount - lastBeat) >= 3000) {
                lastBeat := A_TickCount
                AgentDebugLog("N", "handy_retranscribe_last.ahk:WaitRetranscribeDone", "poll_error", Map(
                    "elapsed", A_TickCount - start,
                    "err", err.Message
                ))
            }
            ; #endregion
        }
        if ((A_TickCount - start) >= maxWaitMs)
            break
        Sleep HANDY_RETRANSCRIBE_POLL_MS
    }
    ; #region agent log
    AgentDebugLog("F", "handy_retranscribe_last.ahk:WaitRetranscribeDone", "timeout", Map(
        "elapsed", A_TickCount - start,
        "sawBusy", sawBusy ? 1 : 0
    ))
    ; #endregion
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
        Sleep 120

        copyOk := HandyRetranscribe_ClickFirstCopy(hwnd)
        ; #region agent log
        AgentDebugLog("I", "handy_retranscribe_last.ahk:HistoryRetranscribeAndCopy", "copy_click", Map(
            "copyOk", copyOk ? 1 : 0
        ))
        ; #endregion
        if (!copyOk) {
            ShowCenteredOverlay_Utils("❌ Could not copy transcription.", 2200, BANNER_ACCENT_ERROR)
            return false
        }

        clipOk := false
        clipLen := 0
        if ClipWait(3) {
            try {
                clipOk := (A_Clipboard != "")
                clipLen := StrLen(A_Clipboard)
            } catch {
                clipOk := false
            }
        }
        ; #region agent log
        AgentDebugLog("I", "handy_retranscribe_last.ahk:HistoryRetranscribeAndCopy", "clipboard", Map(
            "clipOk", clipOk ? 1 : 0,
            "clipLen", clipLen
        ))
        ; #endregion
        if !clipOk {
            ShowCenteredOverlay_Utils("❌ Clipboard did not update.", 2200, BANNER_ACCENT_ERROR)
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
    ; #region agent log
    AgentDebugLog("B", "handy_retranscribe_last.ahk:ToggleModelAndCopy", "enter", Map(
        "targetSlot", targetSlot,
        "scriptName", A_ScriptName,
        "isOwner", HandyAi_IsOwnerProcess() ? 1 : 0,
        "persistedSlot", Handy_GetPersistedAiModelSlot()
    ))
    ; #endregion

    selOk := ExecuteHandyAiModelSelection(targetSlot, true)
    ; #region agent log
    AgentDebugLog("C", "handy_retranscribe_last.ahk:ToggleModelAndCopy", "after ExecuteHandyAiModelSelection", Map(
        "selOk", selOk ? 1 : 0,
        "targetSlot", targetSlot
    ))
    ; #endregion
    if (!selOk) {
        ShowCenteredOverlay_Utils("❌ Could not switch Handy model.", 2200, BANNER_ACCENT_ERROR)
        return false
    }

    hwnd := Handy_ActivateOrLaunch()
    ; #region agent log
    AgentDebugLog("D", "handy_retranscribe_last.ahk:ToggleModelAndCopy", "after ActivateOrLaunch post-select", Map(
        "hwnd", hwnd
    ))
    ; #endregion
    if (!hwnd) {
        ShowCenteredOverlay_Utils("❌ Could not open Handy.", 2000, BANNER_ACCENT_ERROR)
        return false
    }

    ok := HandyRetranscribe_HistoryRetranscribeAndCopy(hwnd)
    ; #region agent log
    AgentDebugLog("I", "handy_retranscribe_last.ahk:ToggleModelAndCopy", "before_close", Map(
        "ok", ok ? 1 : 0,
        "hwnd", hwnd
    ))
    ; #endregion
    if (ok) {
        try WinClose("ahk_id " . hwnd)
        ; #region agent log
        AgentDebugLog("I", "handy_retranscribe_last.ahk:ToggleModelAndCopy", "after_close", Map(
            "ok", 1
        ))
        ; #endregion
    } else {
        ; #region agent log
        AgentDebugLog("J", "handy_retranscribe_last.ahk:ToggleModelAndCopy", "skip_close_on_failure", Map(
            "hwnd", hwnd
        ))
        ; #endregion
    }
    return ok
}
