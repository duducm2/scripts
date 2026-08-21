; =============================================================================
; Utils module: handy_retranscribe_last.ahk
; Send dictation? [B] — toggle Parakeet/Cohere, History re-transcribe, copy
; Reuses HandyReplay_* History/click helpers from handy_replay_last.ahk
; =============================================================================

global HANDY_RETRANSCRIBE_MAX_WAIT_MS := 90000
global HANDY_RETRANSCRIBE_POLL_MS := 250

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
    if RegExMatch(name, "^\d+:\d+$")
        return true
    ; Date labels e.g. "August 18, 2026 at 10:53 PM"
    if InStr(name, " at ")
        return true
    return false
}

; Newest History entry transcription (first non-noise Text in tree order under History list).
HandyRetranscribe_ReadFirstEntryText(el) {
    if !el
        return ""
    try {
        for txt in el.FindAll({ Type: 50020 }) {
            nm := ""
            try nm := txt.Name
            if (nm = "" || HandyRetranscribe_IsNoiseHistoryText(nm))
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
    try {
        el := UIA.ElementFromHandle(hwnd)
        btn := HandyRetranscribe_FindFirstCopy(el)
        if !btn
            return false
        if HandyReplay_ClickWithStrategies(btn, hwnd)
            return true
        try btn.Invoke()
        catch {
            return false
        }
        return true
    } catch {
        return false
    }
}

; Poll until newest entry text changes, or Re-transcribe disables then re-enables.
; Identical transcripts: after a disable→enable cycle, return true even if text unchanged.
HandyRetranscribe_WaitRetranscribeDone(hwnd, priorText, maxWaitMs := 0) {
    global UIA, HANDY_RETRANSCRIBE_MAX_WAIT_MS, HANDY_RETRANSCRIBE_POLL_MS
    if (maxWaitMs <= 0)
        maxWaitMs := HANDY_RETRANSCRIBE_MAX_WAIT_MS
    start := A_TickCount
    sawDisabled := false
    loop {
        try {
            el := UIA.ElementFromHandle(hwnd)
            curText := HandyRetranscribe_ReadFirstEntryText(el)
            if (priorText != "" && curText != "" && curText != priorText)
                return true
            btn := HandyRetranscribe_FindFirstRetranscribe(el)
            if btn {
                en := HandyRetranscribe_ButtonEnabled(btn)
                if !en
                    sawDisabled := true
                else if (sawDisabled && en)
                    return true
            }
            ; Model button may show "loading" while Cohere works.
            try {
                modelName := Handy_ReadActiveAiModelName(hwnd)
                if (modelName != "" && InStr(modelName, "loading"))
                    sawDisabled := true
                else if (sawDisabled && modelName != "" && !InStr(modelName, "loading")
                && (priorText = "" || curText != priorText || (A_TickCount - start) > 1500))
                    return true
            } catch {
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

        if (!HandyRetranscribe_ClickFirstCopy(hwnd)) {
            ShowCenteredOverlay_Utils("❌ Could not copy transcription.", 2200, BANNER_ACCENT_ERROR)
            return false
        }

        clipOk := false
        if ClipWait(3) {
            try clipOk := (A_Clipboard != "")
            catch {
                clipOk := false
            }
        }
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

    ; ExecuteHandyAiModelSelection shows its own AiModelBanner during switch.
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
    try WinClose("ahk_id " . hwnd)
    return ok
}
