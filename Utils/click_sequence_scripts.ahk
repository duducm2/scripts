; =============================================================================
; Utils module: click_sequence_scripts.ahk
; Hardcoded Script registry for Slot chains. Pick-list only; no runtime authoring.
; =============================================================================

global g_ClickSeqRunCtx := ""

ClickSeqScript_Catalog() {
    return [{ id: "scrollFeedBottom", title: "Scroll feed to bottom" }, { id: "desktopWait", title: "Wait for new Desktop file" }, { id: "desktopCut",
        title: "Cut newest Desktop item" }, { id: "focusCompanion", title: "Focus AI companion" }
    ]
}

ClickSeqScript_Title(scriptId) {
    for row in ClickSeqScript_Catalog() {
        if (row.id = scriptId)
            return row.title
    }
    return scriptId
}

ClickSeqScript_IsValid(scriptId) {
    id := ClickSeqData_SanitizeId(scriptId)
    for row in ClickSeqScript_Catalog() {
        if (row.id = id)
            return true
    }
    return false
}

ClickSeqScript_Ctx() {
    global g_ClickSeqRunCtx
    if (!IsObject(g_ClickSeqRunCtx))
        g_ClickSeqRunCtx := {}
    return g_ClickSeqRunCtx
}

; Returns true on success.
ClickSeqScript_Run(scriptId) {
    id := ClickSeqData_SanitizeId(scriptId)
    ctx := ClickSeqScript_Ctx()
    ; #region agent log
    try {
        eqI := (id = "scrollFeedBottom")
        eqS := (id == "scrollFeedBottom")
        t := Type(id)
        valid := ClickSeqScript_IsValid(scriptId)
        FileAppend(
            '{"sessionId":"dbd429","runId":"run1","hypothesisId":"A","location":"click_sequence_scripts.ahk:ClickSeqScript_Run","message":"script id before switch","data":{"raw":"' .
            scriptId . '","sanitized":"' . id . '","type":"' . t . '","eqInsensitive":' . (eqI ? "true" : "false") .
            ',"eqSensitive":' . (eqS ? "true" : "false") . ',"isValid":' . (valid ? "true" : "false") .
            '},"timestamp":' . A_TickCount . '}`n', A_ScriptDir "\debug-dbd429.log", "UTF-8")
    } catch {
    }
    ; #endregion
    switch id {
        case "scrollfeedbottom":
            ; #region agent log
            try FileAppend(
                '{"sessionId":"dbd429","runId":"run1","hypothesisId":"A","location":"click_sequence_scripts.ahk:switch","message":"matched scrollFeedBottom case","data":{"sanitized":"' .
                id . '"},"timestamp":' . A_TickCount . '}`n', A_ScriptDir "\debug-dbd429.log", "UTF-8")
            catch {
            }
            ; #endregion
            hwnd := ctx.HasProp("hwnd") ? ctx.hwnd : 0
            companion := ctx.HasProp("companion") ? ctx.companion : ""
            try AiQuickDownload_ScrollFeedToBottom(hwnd, companion)
            catch {
            }
            return true
        case "desktopwait":
            return ClickSeqScript_DesktopWait()
        case "desktopcut":
            return ClickSeqScript_DesktopCut()
        case "focuscompanion":
            return ClickSeqScript_FocusCompanion()
        default:
            ; #region agent log
            try FileAppend(
                '{"sessionId":"dbd429","runId":"run1","hypothesisId":"A","location":"click_sequence_scripts.ahk:switch","message":"hit default unknown script","data":{"sanitized":"' .
                id . '","raw":"' . scriptId . '"},"timestamp":' . A_TickCount . '}`n', A_ScriptDir "\debug-dbd429.log",
                "UTF-8")
            catch {
            }
            ; #endregion
            ClickSeq_SetFail("❌ Unknown Hardcoded Script: " . id)
            return false
    }
}

ClickSeqScript_DesktopWait() {
    ctx := ClickSeqScript_Ctx()
    desktopPath := ctx.HasProp("desktopPath") ? ctx.desktopPath : ""
    beforePath := ctx.HasProp("beforePath") ? ctx.beforePath : ""
    beforeStamp := ctx.HasProp("beforeStamp") ? ctx.beforeStamp : ""
    try StandardLoadingBar_Update("⏳ Waiting for Desktop file…", BANNER_ACCENT_INTERMEDIATE)
    catch {
    }
    newPath := ""
    try newPath := AiQuickDownload_WaitForNewDesktopFile(desktopPath, beforePath, beforeStamp)
    catch {
        newPath := ""
    }
    if (newPath = "") {
        ClickSeq_SetFail("❌ Quick Download: file did not appear on Desktop")
        return false
    }
    ctx.lastPath := newPath
    return true
}

ClickSeqScript_DesktopCut() {
    ctx := ClickSeqScript_Ctx()
    if (ctx.HasProp("doCut") && !ctx.doCut)
        return true
    try StandardLoadingBar_Hide(0)
    catch {
    }
    try DesktopCutNewest_Trigger()
    catch as e {
        ClickSeq_SetFail("❌ Quick Download: cut failed — " . SubStr(e.Message, 1, 60))
        return false
    }
    return true
}

ClickSeqScript_FocusCompanion() {
    ctx := ClickSeqScript_Ctx()
    focus := ""
    try focus := AiQuickDownload_FocusCompanion()
    catch {
        ClickSeq_SetFail("❌ Could not focus AI companion.")
        return false
    }
    if (!IsObject(focus) || !focus.ok) {
        ClickSeq_SetFail(IsObject(focus) ? focus.err : "❌ Could not focus AI companion.")
        return false
    }
    ctx.hwnd := focus.hwnd
    ctx.companion := focus.companion
    return true
}
