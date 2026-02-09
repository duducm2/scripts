#Requires AutoHotkey v2.0
; =============================================================================
; GeminiToCursorBridge.ahk — Self-contained module: copy Gemini last message
; to clipboard, go back to any Cursor window, focus AI field, paste, send.
; No project-path matching: just find any Cursor window and paste there.
; Single entry: CopyFromGeminiToCursor(). projectPath used only to launch Cursor if no window exists.
;
; Verification layers (fail fast with clear reason):
;   V1 — After copy: re-check clipboard length and result file before returning.
;   V2 — After WinActivate/WinWaitActive: confirm active window = target; after
;        focus AI: confirm still on target (retry activate if not).
;   V3 — Before Cursor: clipboard still valid; before paste: active = target and
;        clipboard still valid (retry activate once if wrong window).
; =============================================================================

; Cursor executable paths (copied config)
BRIDGE_CursorExePersonal := "C:\Users\eduev\AppData\Local\Programs\cursor\Cursor.exe"
BRIDGE_CursorExeWork     := "C:\Users\fie7ca\AppData\Local\Programs\cursor\Cursor.exe"
WM_COPY_LAST_GEMINI     := 0x8001
; Path for guarantee layer: Gemini.ahk writes "1" here when Copy Last Response (same as #!+p) succeeds
BRIDGE_GeminiCopyResultPath := A_ScriptDir "\.cursor\gemini_copy_result.txt"
BRIDGE_MinClipboardLength   := 10

; #region agent log
Bridge_Log(loc, msg, data, hypothesisId := "") {
    p := A_ScriptDir "\.cursor\debug.log"
    j := '{"location":"' . loc . '","message":"' . msg . '","data":' . (data is String ? data : "{}") . ',"hypothesisId":"' . hypothesisId . '","timestamp":' . A_TickCount . '}'
    try
        FileAppend j "`n", p
    catch
        return
}
; #endregion

; Move mouse to window center (minimal copy; no halo)
Bridge_MoveMouseToCenter(hwnd) {
    if (!hwnd)
        return
    rect := Buffer(16, 0)
    if (!DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect))
        return
    left := NumGet(rect, 0, "int")
    top := NumGet(rect, 4, "int")
    right := NumGet(rect, 8, "int")
    bottom := NumGet(rect, 12, "int")
    centerX := left + (right - left) // 2
    centerY := top + (bottom - top) // 2
    DllCall("SetCursorPos", "int", centerX, "int", centerY)
}

; Extract path segments for Cursor window title matching. More specific first so we prefer exact project.
Bridge_ExtractProjectMatchSegments(projectPath) {
    normalizedPath := RTrim(projectPath, "\")
    pathSegments := StrSplit(normalizedPath, "\")
    lastSegment := pathSegments[pathSegments.Length]
    matchSegments := []
    if (pathSegments.Length >= 2) {
        lastTwoJoined := pathSegments[pathSegments.Length - 1] . " - " . pathSegments[pathSegments.Length]
        if (lastTwoJoined != lastSegment)
            matchSegments.Push(lastTwoJoined)
    }
    matchSegments.Push(lastSegment)
    return matchSegments
}

; Return any Cursor window hwnd (skip preview). Used to "go back to Cursor" and paste.
Bridge_GetAnyCursorHwnd() {
    DetectHiddenWindows true
    try {
        for hwnd in WinGetList("ahk_exe Cursor.exe") {
            try {
                winTitle := WinGetTitle("ahk_id " hwnd)
                if (InStr(StrLower(winTitle), "preview"))
                    continue
                DetectHiddenWindows false
                return hwnd
            } catch {
                continue
            }
        }
    } catch {
    }
    DetectHiddenWindows false
    return 0
}

; Return Cursor hwnd for project path, or 0. Prefer match on most specific segment so we don't pick another project.
Bridge_GetCursorHwndForProject(projectPath) {
    matchSegments := Bridge_ExtractProjectMatchSegments(projectPath)
    DetectHiddenWindows true
    try {
        for segment in matchSegments {
            for hwnd in WinGetList("ahk_exe Cursor.exe") {
                try {
                    winTitle := WinGetTitle("ahk_id " hwnd)
                    if (InStr(StrLower(winTitle), "preview"))
                        continue
                    if (InStr(winTitle, segment)) {
                        DetectHiddenWindows false
                        return hwnd
                    }
                } catch {
                    continue
                }
            }
        }
    } catch {
    }
    DetectHiddenWindows false
    return 0
}

; Find and activate Cursor window for project path; return hwnd or 0. Include hidden/minimized.
; Only considers windows matching the most specific segment that has any match (avoids wrong project).
Bridge_FindAndActivateCursorWindow(projectPath) {
    matchSegments := Bridge_ExtractProjectMatchSegments(projectPath)
    cursorWindows := []
    DetectHiddenWindows true
    try {
        for segment in matchSegments {
            cursorWindows := []
            for hwnd in WinGetList("ahk_exe Cursor.exe") {
                try {
                    winTitle := WinGetTitle("ahk_id " hwnd)
                    if (InStr(StrLower(winTitle), "preview"))
                        continue
                    if (InStr(winTitle, segment)) {
                        cursorWindows.Push({ hwnd: hwnd, title: winTitle })
                    }
                } catch {
                    continue
                }
            }
            if (cursorWindows.Length > 0)
                break
        }
    } catch {
        DetectHiddenWindows false
        return 0
    }
    DetectHiddenWindows false
    ; #region agent log
    pathNorm := RTrim(projectPath, "\")
    pathParts := StrSplit(pathNorm, "\")
    lastSeg := pathParts.Length ? pathParts[pathParts.Length] : ""
    segUsed := cursorWindows.Length ? "has_match" : "none"
    titlesJson := ""
    for i, w in cursorWindows {
        t := SubStr(StrReplace(StrReplace(w.title, "\", " "), '"', "'"), 1, 45)
        titlesJson .= (i > 1 ? "|" : "") . t
    }
    Bridge_Log("GeminiToCursorBridge.ahk:FindAndActivate", "match result", '{"pathLast":"' . lastSeg . '","matchCount":' . cursorWindows.Length . ',"segmentUsed":"' . segUsed . '","titles":"' . titlesJson . '"}', "H1")
    ; #endregion
    if (cursorWindows.Length = 0)
        return 0
    try {
        activeHwnd := WinGetID("A")
        for window in cursorWindows {
            if (window.hwnd = activeHwnd) {
                WinActivate("ahk_id " window.hwnd)
                Bridge_MoveMouseToCenter(window.hwnd)
                ; #region agent log
                Bridge_Log("GeminiToCursorBridge.ahk:FindAndActivate", "picked active", '{"hwnd":' . window.hwnd . ',"titleStart":"' . SubStr(StrReplace(window.title, '"', "'"), 1, 50) . '"}', "H2")
                ; #endregion
                return window.hwnd
            }
        }
    } catch {
    }
    targetWindow := cursorWindows[1]
    ; #region agent log
    Bridge_Log("GeminiToCursorBridge.ahk:FindAndActivate", "picked first", '{"hwnd":' . targetWindow.hwnd . ',"titleStart":"' . SubStr(StrReplace(targetWindow.title, '"', "'"), 1, 60) . '"}', "H2")
    ; #endregion
    try {
        WinActivate("ahk_id " targetWindow.hwnd)
        WinWaitActive("ahk_id " targetWindow.hwnd, , 2)
        Bridge_MoveMouseToCenter(targetWindow.hwnd)
        return targetWindow.hwnd
    } catch {
        return 0
    }
}

; Focus Cursor AI text field (keyboard-only; copied)
Bridge_FocusCursorAITextField(targetHwnd := 0) {
    try {
        if (targetHwnd) {
            WinActivate("ahk_id " targetHwnd)
            WinWaitActive("ahk_id " targetHwnd, , 2)
        } else {
            targetHwnd := WinExist("ahk_exe Cursor.exe")
            if (!targetHwnd)
                return false
            WinWaitActive("ahk_id " targetHwnd, , 2)
        }
        Sleep 200
        Send "^i"
        Sleep 1200
        Send "{Tab 2}"
        Sleep 100
        return true
    } catch {
        return false
    }
}

; Copy Gemini last message to clipboard. Returns { ok: true } or { ok: false, reason: "..." }. Copied logic.
Bridge_CopyGeminiLastMessageToClipboard() {
    clipBefore := A_Clipboard
    prevMatch := A_TitleMatchMode
    SetTitleMatchMode 2
    DetectHiddenWindows true
    geminiScriptHwnd := 0
    for hwnd in WinGetList("ahk_exe AutoHotkey64.exe") {
        try {
            if (InStr(WinGetTitle("ahk_id " hwnd), "Gemini.ahk")) {
                geminiScriptHwnd := hwnd
                break
            }
        } catch {
            continue
        }
    }
    if (!geminiScriptHwnd) {
        for hwnd in WinGetList("ahk_exe AutoHotkey32.exe") {
            try {
                if (InStr(WinGetTitle("ahk_id " hwnd), "Gemini.ahk")) {
                    geminiScriptHwnd := hwnd
                    break
                }
            } catch {
                continue
            }
        }
    }
    DetectHiddenWindows false
    SetTitleMatchMode prevMatch
    ; #region agent log
    Bridge_Log("GeminiToCursorBridge.ahk:Bridge_CopyGemini", "geminiScriptHwnd found", '{"hwnd":' . geminiScriptHwnd . '}', "H1")
    ; #endregion
    if (!geminiScriptHwnd)
        return { ok: false, reason: "no_script" }

    ; Find and activate Gemini browser window
    geminiBrowserHwnd := 0
    try {
        for hwnd in WinGetList("ahk_exe chrome.exe") {
            try {
                if (InStr(WinGetTitle("ahk_id " hwnd), "gemini", false)) {
                    geminiBrowserHwnd := hwnd
                    break
                }
            } catch {
                continue
            }
        }
    } catch {
    }
    if (!geminiBrowserHwnd)
        return { ok: false, reason: "no_gemini_window" }
    activated := false
    loop 2 {
        try {
            WinActivate("ahk_id " geminiBrowserHwnd)
            if (WinWaitActive("ahk_id " geminiBrowserHwnd, , 3)) {
                activated := true
                break
            }
        } catch {
        }
        if (A_Index = 1)
            Sleep 300
    }
    if (!activated)
        return { ok: false, reason: "gemini_activate_failed" }
    Sleep 200

    ; Guarantee layer: reset result file so we only accept a copy that Gemini.ahk confirms (same as #!+p path).
    try {
        if (FileExist(BRIDGE_GeminiCopyResultPath))
            FileDelete(BRIDGE_GeminiCopyResultPath)
        FileAppend("0", BRIDGE_GeminiCopyResultPath)
    } catch {
    }

    ; Re-find Gemini script window with DetectHiddenWindows so PostMessage target is findable (avoids "Target window not found").
    DetectHiddenWindows true
    SetTitleMatchMode 2
    postTargetHwnd := 0
    for hwnd in WinGetList("ahk_exe AutoHotkey64.exe") {
        try {
            if (InStr(WinGetTitle("ahk_id " hwnd), "Gemini.ahk")) {
                postTargetHwnd := hwnd
                break
            }
        } catch {
            continue
        }
    }
    if (!postTargetHwnd) {
        for hwnd in WinGetList("ahk_exe AutoHotkey32.exe") {
            try {
                if (InStr(WinGetTitle("ahk_id " hwnd), "Gemini.ahk")) {
                    postTargetHwnd := hwnd
                    break
                }
            } catch {
                continue
            }
        }
    }
    ; #region agent log
    Bridge_Log("GeminiToCursorBridge.ahk:Bridge_CopyGemini", "before PostMessage", '{"hwnd":' . postTargetHwnd . '}', "H2")
    ; #endregion
    if (!postTargetHwnd) {
        DetectHiddenWindows false
        return { ok: false, reason: "no_script" }
    }
    try {
        PostMessage(WM_COPY_LAST_GEMINI, 0, 0, , "ahk_id " postTargetHwnd)
    } catch as err {
        ; #region agent log
        em := err.HasProp("Message") ? StrReplace(StrReplace(err.Message, "\", " "), '"', "'") : ""
        ew := err.HasProp("What") ? err.What : ""
        Bridge_Log("GeminiToCursorBridge.ahk:Bridge_CopyGemini", "PostMessage catch", '{"message":"' . em . '","what":"' . ew . '"}', "H3")
        ; #endregion
        DetectHiddenWindows false
        return { ok: false, reason: "send_failed" }
    }
    DetectHiddenWindows false

    copyWaitMax := 120
    copyWaitMs := 150
    copyDone := false
    attempt := 1
    loop 2 {
        loop copyWaitMax {
            Sleep copyWaitMs
            if (A_Clipboard != "" && A_Clipboard != clipBefore) {
                copyDone := true
                break 2
            }
        }
        if (copyDone)
            break
        if (attempt = 1) {
            attempt := 2
            DetectHiddenWindows true
            postTargetHwnd := 0
            for hwnd in WinGetList("ahk_exe AutoHotkey64.exe") {
                try {
                    if (InStr(WinGetTitle("ahk_id " hwnd), "Gemini.ahk")) {
                        postTargetHwnd := hwnd
                        break
                    }
                } catch {
                    continue
                }
            }
            if (!postTargetHwnd) {
                for hwnd in WinGetList("ahk_exe AutoHotkey32.exe") {
                    try {
                        if (InStr(WinGetTitle("ahk_id " hwnd), "Gemini.ahk")) {
                            postTargetHwnd := hwnd
                            break
                        }
                    } catch {
                        continue
                    }
                }
            }
            try {
                if (postTargetHwnd)
                    PostMessage(WM_COPY_LAST_GEMINI, 0, 0, , "ahk_id " postTargetHwnd)
            } catch {
                DetectHiddenWindows false
                return { ok: false, reason: "timeout" }
            }
            DetectHiddenWindows false
        } else {
            return { ok: false, reason: "timeout" }
        }
    }
    if (!copyDone)
        return { ok: false, reason: "timeout" }
    clipNow := A_Clipboard
    if (clipNow = "" || clipNow = clipBefore)
        return { ok: false, reason: "validation_failed" }
    if (Trim(clipNow) = "")
        return { ok: false, reason: "validation_failed" }
    ; Guarantee: require minimum length so we have a real Gemini response, not a stray copy.
    if (StrLen(Trim(clipNow)) < BRIDGE_MinClipboardLength)
        return { ok: false, reason: "validation_failed" }
    ; Guarantee: require Gemini.ahk to have confirmed copy success (same code path as #!+p).
    try {
        if (FileExist(BRIDGE_GeminiCopyResultPath)) {
            resultContent := Trim(FileRead(BRIDGE_GeminiCopyResultPath))
            if (resultContent != "1")
                return { ok: false, reason: "validation_failed" }
        } else {
            return { ok: false, reason: "validation_failed" }
        }
    } catch {
        return { ok: false, reason: "validation_failed" }
    }
    ; Verification layer: re-read clipboard and result file once more before returning (guards against race).
    clipStable := A_Clipboard
    if (clipStable = "" || clipStable = clipBefore || StrLen(Trim(clipStable)) < BRIDGE_MinClipboardLength) {
        ; #region agent log
        Bridge_Log("GeminiToCursorBridge.ahk:Bridge_CopyGemini", "verify failed clipboard", '{"len":' . StrLen(clipStable) . '}', "V1")
        ; #endregion
        return { ok: false, reason: "validation_failed" }
    }
    try {
        if (!FileExist(BRIDGE_GeminiCopyResultPath) || Trim(FileRead(BRIDGE_GeminiCopyResultPath)) != "1") {
            ; #region agent log
            Bridge_Log("GeminiToCursorBridge.ahk:Bridge_CopyGemini", "verify failed result file", "{}", "V1")
            ; #endregion
            return { ok: false, reason: "validation_failed" }
        }
    } catch {
        return { ok: false, reason: "validation_failed" }
    }
    ; #region agent log
    Bridge_Log("GeminiToCursorBridge.ahk:Bridge_CopyGemini", "verify passed", '{"clipLen":' . StrLen(clipStable) . '}', "V1")
    ; #endregion
    return { ok: true }
}

; Activate any Cursor window and focus AI field (go back to Cursor and paste). Returns target hwnd on success, 0 on failure.
; If no Cursor window exists and projectPath is valid, launches Cursor with that folder.
Bridge_ActivateCursorProject(projectPath, isWorkEnvironment) {
    ; #region agent log
    Bridge_Log("GeminiToCursorBridge.ahk:ActivateProject", "entry", '{"pathLen":' . StrLen(projectPath) . '}', "H3")
    ; #endregion
    ; Just find any Cursor window (no project-path matching).
    targetHwnd := Bridge_GetAnyCursorHwnd()
    ; #region agent log
    Bridge_Log("GeminiToCursorBridge.ahk:ActivateProject", "after GetAny", '{"targetHwnd":' . targetHwnd . '}', "H3")
    ; #endregion
    if (!targetHwnd && projectPath != "" && DirExist(projectPath)) {
        cursorPath := isWorkEnvironment ? BRIDGE_CursorExeWork : BRIDGE_CursorExePersonal
        try {
            Run cursorPath . ' "' . projectPath . '"'
        } catch {
            return 0
        }
        loop 30 {
            Sleep 200
            targetHwnd := Bridge_GetAnyCursorHwnd()
            if (targetHwnd)
                break
        }
        ; #region agent log
        Bridge_Log("GeminiToCursorBridge.ahk:ActivateProject", "after launch wait", '{"targetHwnd":' . targetHwnd . '}', "H4")
        ; #endregion
        if (!targetHwnd)
            return 0
    }
    if (!targetHwnd)
        return 0
    try {
        WinActivate("ahk_id " targetHwnd)
        WinWaitActive("ahk_id " targetHwnd, , 3)
    } catch {
        return 0
    }
    ; Verification: active window must be our target.
    try {
        if (WinGetID("A") != targetHwnd) {
            ; #region agent log
            Bridge_Log("GeminiToCursorBridge.ahk:ActivateProject", "verify active failed", '{"expected":' . targetHwnd . ',"got":' . WinGetID("A") . '}', "V2")
            ; #endregion
            WinActivate("ahk_id " targetHwnd)
            WinWaitActive("ahk_id " targetHwnd, , 2)
            if (WinGetID("A") != targetHwnd) {
                ; #region agent log
                Bridge_Log("GeminiToCursorBridge.ahk:ActivateProject", "verify active retry failed", "{}", "V2")
                ; #endregion
                return 0
            }
        }
    } catch {
        return 0
    }
    Sleep 300
    if (!Bridge_FocusCursorAITextField(targetHwnd))
        return 0
    ; Verification: after focus, confirm we're still on target (focus can steal).
    try {
        if (WinGetID("A") != targetHwnd) {
            ; #region agent log
            Bridge_Log("GeminiToCursorBridge.ahk:ActivateProject", "verify after focus failed", '{"expected":' . targetHwnd . ',"got":' . WinGetID("A") . '}', "V2")
            ; #endregion
            WinActivate("ahk_id " targetHwnd)
            WinWaitActive("ahk_id " targetHwnd, , 2)
        }
    } catch {
    }
    ; #region agent log
    tTitle := ""
    try
        tTitle := SubStr(StrReplace(WinGetTitle("ahk_id " targetHwnd), Chr(34), "'"), 1, 60)
    catch
        tTitle := ""
    Bridge_Log("GeminiToCursorBridge.ahk:ActivateProject", "return hwnd", '{"hwnd":' . targetHwnd . ',"titleStart":"' . tTitle . '"}', "H3")
    ; #endregion
    return targetHwnd
}

; =============================================================================
; Entry point: copy Gemini last message, go back to Cursor window, paste, send.
; Finds any Cursor window (no project matching); if none, launches Cursor with projectPath.
; projectPath: used only when no Cursor window exists (launch folder).
; isWorkEnvironment: true = work Cursor exe, false = personal.
; Returns: { ok: true } or { ok: false, reason: "..." }
; =============================================================================
CopyFromGeminiToCursor(projectPath, isWorkEnvironment := false) {
    ; #region agent log
    Bridge_Log("GeminiToCursorBridge.ahk:CopyFromGemini", "entry", '{"pathLen":' . StrLen(projectPath) . '}', "H5")
    ; #endregion
    copyResult := Bridge_CopyGeminiLastMessageToClipboard()
    if (!copyResult.ok)
        return copyResult
    ; Verification: clipboard still has content before we switch away (nothing overwrote it).
    if (StrLen(Trim(A_Clipboard)) < BRIDGE_MinClipboardLength) {
        ; #region agent log
        Bridge_Log("GeminiToCursorBridge.ahk:CopyFromGemini", "verify clipboard before activate", '{"len":' . StrLen(A_Clipboard) . '}', "V3")
        ; #endregion
        return { ok: false, reason: "validation_failed" }
    }
    targetHwnd := Bridge_ActivateCursorProject(projectPath, isWorkEnvironment)
    if (!targetHwnd)
        return { ok: false, reason: "cursor_activate_failed" }
    ; Re-activate the selected project window immediately before paste so paste always goes there (e.g. when triggered from external window).
    try {
        WinActivate("ahk_id " targetHwnd)
        WinWaitActive("ahk_id " targetHwnd, , 2)
        Sleep 100
    } catch {
        return { ok: false, reason: "cursor_activate_failed" }
    }
    ; Verification layer: active window must be target and clipboard still valid before paste.
    try {
        activeNow := WinGetID("A")
        if (activeNow != targetHwnd) {
            ; #region agent log
            Bridge_Log("GeminiToCursorBridge.ahk:CopyFromGemini", "verify before paste active failed", '{"expected":' . targetHwnd . ',"got":' . activeNow . '}', "V3")
            ; #endregion
            WinActivate("ahk_id " targetHwnd)
            WinWaitActive("ahk_id " targetHwnd, , 2)
            if (WinGetID("A") != targetHwnd) {
                ; #region agent log
                Bridge_Log("GeminiToCursorBridge.ahk:CopyFromGemini", "verify before paste retry failed", "{}", "V3")
                ; #endregion
                return { ok: false, reason: "cursor_activate_failed" }
            }
        }
    } catch {
        return { ok: false, reason: "cursor_activate_failed" }
    }
    if (StrLen(Trim(A_Clipboard)) < BRIDGE_MinClipboardLength) {
        ; #region agent log
        Bridge_Log("GeminiToCursorBridge.ahk:CopyFromGemini", "verify clipboard before paste", '{"len":' . StrLen(A_Clipboard) . '}', "V3")
        ; #endregion
        return { ok: false, reason: "validation_failed" }
    }
    ; #region agent log
    pasteTitle := ""
    try
        pasteTitle := SubStr(StrReplace(WinGetTitle("ahk_id " targetHwnd), Chr(34), "'"), 1, 60)
    catch
        pasteTitle := ""
    Bridge_Log("GeminiToCursorBridge.ahk:CopyFromGemini", "before paste verify ok", '{"hwnd":' . targetHwnd . ',"titleStart":"' . pasteTitle . '"}', "V3")
    ; #endregion
    Send "^v"
    Send "{Enter}"
    return { ok: true }
}
