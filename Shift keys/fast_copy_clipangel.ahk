; =============================================================================
; Shift keys module: fast_copy_clipangel.ahk
; Clip Angel fast copy mode and #!+1/#!+J
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

; =============================================================================
; Send Top List Item from Clip Angel
; Hotkey: Win+Alt+Shift+1
; =============================================================================
#!+1::
{
    ClipAngel_SendTopListItem()
}

; =============================================================================
; Clip Angel: Fast Copy Mode + sequential paste (multiple clips in order)
; Hotkey: Win+Alt+Shift+J — while mode off: tap starts mode; hold 700ms+ repeats last paste count.
;         While mode on: press finishes and pastes N clips (Ctrl+C / PrtSc / Alt+PrtSc counted).
; =============================================================================
global FAST_COPY_HOLD_REPEAT_MS := 700
; Set true to write FastCopyMode_DebugLog NDJSON (development only).
global FAST_COPY_DEBUG := false
global FAST_COPY_CLIPBOARD_READ_CYCLE_MS := 1200
global FAST_COPY_GEMINI_UPLOAD_IDLE_MS := 5000
global gFastCopyModeActive := false
global gFastCopyCount := 0
global gFastCopyPasteTargetHwnd := 0
global gFastCopyLastSuccessfulCount := 0
global gFastCopyScreenshotQueue := []
global gFastCopyLastScreenshotQueue := []

FastCopyMode_DebugLog(hypothesisId, location, message, data := "") {
    ; #region agent log
    global FAST_COPY_DEBUG
    if (!FAST_COPY_DEBUG)
        return
    ; Writes NDJSON to debug-1bed80.log (Debug session: 1bed80)
    try {
        runId := "pre-fix"
        logPath := "C:\Users\fie7ca\Documents\scripts\debug-1bed80.log"
        ; Keep data small and non-sensitive; accept either a string or a Map-like object.
        dataJson := "{}"
        if (IsObject(data)) {
            parts := []
            for k, v in data {
                try parts.Push('"' FastCopyMode_JsonEscape(k) '":"' FastCopyMode_JsonEscape(v) '"')
            }
            joined := ""
            if (parts.Length) {
                for i, p in parts {
                    joined .= (i = 1 ? "" : ",") p
                }
            }
            dataJson := "{" joined "}"
        } else if (data != "") {
            dataJson := '{"value":"' FastCopyMode_JsonEscape(data) '"}'
        }
        line := '{'
            . '"sessionId":"1bed80",'
            . '"timestamp":' A_TickCount + 0 ','
            . '"runId":"' runId '",'
            . '"hypothesisId":"' FastCopyMode_JsonEscape(hypothesisId) '",'
            . '"location":"' FastCopyMode_JsonEscape(location) '",'
            . '"message":"' FastCopyMode_JsonEscape(message) '",'
            . '"data":' dataJson
            . '}'
        FileAppend(line "`n", logPath, "UTF-8")
    } catch {
        ; never break user flow
    }
    ; #endregion agent log
}

FastCopyMode_JsonEscape(s) {
    ; #region agent log
    try {
        if (s = "")
            return ""
    } catch {
        return ""
    }
    s := "" s
    s := StrReplace(s, "\", "\\")
    s := StrReplace(s, '"', '\"')
    s := StrReplace(s, "`r", "\r")
    s := StrReplace(s, "`n", "\n")
    return s
    ; #endregion agent log
}

FastCopyMode_ClipboardHasImage() {
    ; #region agent log
    ; CF_DIB=8, CF_DIBV5=17, CF_BITMAP=2
    try {
        return !!(DllCall("IsClipboardFormatAvailable", "UInt", 8, "Int")
        || DllCall("IsClipboardFormatAvailable", "UInt", 17, "Int")
        || DllCall("IsClipboardFormatAvailable", "UInt", 2, "Int"))
    } catch {
        return false
    }
    ; #endregion agent log
}

FastCopyMode_WaitForClipboardImage(timeoutMs := 1200) {
    ; #region agent log
    start := A_TickCount
    while ((A_TickCount - start) < timeoutMs) {
        if (FastCopyMode_ClipboardHasImage())
            return true
        Sleep 15
    }
    return false
    ; #endregion agent log
}

FastCopyMode_CanOpenClipboardNow() {
    ; #region agent log
    ; Returns true if clipboard can be opened immediately (no long lock).
    ok := false
    try {
        if (DllCall("OpenClipboard", "ptr", 0, "int")) {
            ok := true
            DllCall("CloseClipboard")
        }
    } catch {
        ok := false
    }
    return ok
    ; #endregion agent log
}

FastCopyMode_WaitForClipboardUnlocked(timeoutMs := 2000) {
    ; #region agent log
    start := A_TickCount
    while ((A_TickCount - start) < timeoutMs) {
        if (FastCopyMode_CanOpenClipboardNow())
            return true
        Sleep 15
    }
    return false
    ; #endregion agent log
}

FastCopyMode_WaitForClipboardReadCycle(timeoutMs := 1200) {
    ; #region agent log
    ; Wait until we observe some other process reading/locking the clipboard (OpenClipboard fails)
    ; and then wait until it becomes unlocked again. This reduces the risk of overwriting the
    ; clipboard before the target app has actually consumed the bitmap.
    start := A_TickCount
    sawLock := false
    failCount := 0
    okCount := 0

    ; Phase 1: short window with Sleep 1 (avoid busy-spinning the hotkey thread).
    sampleUntil := start + 120
    while (A_TickCount < sampleUntil) {
        if (!FastCopyMode_CanOpenClipboardNow()) {
            failCount += 1
            sawLock := true
            break
        }
        okCount += 1
        Sleep 1
    }

    ; Phase 2: regular polling until timeout budget (handles longer locks).
    if (!sawLock) {
        while ((A_TickCount - start) < timeoutMs) {
            if (!FastCopyMode_CanOpenClipboardNow()) {
                failCount += 1
                sawLock := true
                break
            }
            okCount += 1
            Sleep 10
        }
    }
    if (sawLock) {
        unlockMs := Min(1500, Max(100, timeoutMs - (A_TickCount - start)))
        if (FastCopyMode_WaitForClipboardUnlocked(unlockMs))
            return "lock_then_unlock(fails=" failCount ",oks=" okCount ")"
        return "lock_timeout(fails=" failCount ",oks=" okCount ")"
    }
    ; Never saw the clipboard get locked (target might not read immediately or at all).
    return "no_lock_seen(fails=" failCount ",oks=" okCount ")"
    ; #endregion agent log
}

FastCopyMode_SendPasteAndWaitForReadCycle(isGeminiSession := false) {
    ; #region agent log
    Send "^v"
    if (isGeminiSession) {
        Sleep 200
        return "gemini_delay"
    } else {
        Sleep 400
        return "baseline_delay"
    }
    ; #endregion agent log
}

FastCopyMode_IsGeminiActiveInChrome() {
    ; #region agent log
    try {
        return (WinActive("ahk_exe chrome.exe") && InStr(WinGetTitle("A"), "gemini", false))
    } catch {
        return false
    }
    ; #endregion agent log
}

FastCopyMode_GetActiveChromeHwnd() {
    ; #region agent log
    try {
        hwnd := WinGetID("A")
        if (hwnd && WinExist("ahk_id " hwnd) && WinActive("ahk_id " hwnd))
            return hwnd
    } catch {
    }
    return 0
    ; #endregion agent log
}

FastCopyMode_UiaForActiveChrome() {
    ; #region agent log
    hwnd := FastCopyMode_GetActiveChromeHwnd()
    if (!hwnd)
        throw Error("no_active_hwnd")
    return UIA_Browser("ahk_id " hwnd)
    ; #endregion agent log
}

FastCopyMode_GetGeminiSearchRoot(uia) {
    ; #region agent log
    ; Same pattern as GetGeminiSearchRoot in Gemini.ahk — smaller UIA subtree than full document.
    try {
        root := uia.GetCurrentMainPaneElement()
        if (root)
            return root
    } catch {
    }
    return uia
    ; #endregion agent log
}

FastCopyMode_FocusGeminiPromptField(uia := "") {
    ; #region agent log
    ; Optional uia: reuse cached UIA_Browser for the whole paste loop (efficiency).
    if (uia = "") {
        try {
            uia := FastCopyMode_UiaForActiveChrome()
        } catch Error as e {
            FastCopyMode_DebugLog("H6", "Shift keys.ahk:FastCopyMode_FocusGeminiPromptField", "uia_attach_error", Map(
                "msg", SubStr(e.Message, 1, 120)
            ))
            return false
        }
    }

    try {
        ; Use shared helper from Utils.ahk (Gemini.ahk uses the same).
        promptField := FindGeminiPromptField(uia)

        if (promptField && IsObject(promptField)) {
            try {
                promptField.SetFocus()
            } catch {
                try promptField.Click()
                catch {
                }
            }
            Sleep 80
            try {
                if (promptField.HasKeyboardFocus)
                    return true
            } catch {
                return true
            }
        }
    } catch Error as e {
        FastCopyMode_DebugLog("H6", "Shift keys.ahk:FastCopyMode_FocusGeminiPromptField", "focus_error", Map(
            "msg", SubStr(e.Message, 1, 120)
        ))
    }
    return false
    ; #endregion agent log
}

FastCopyMode_GeminiIsUploadingImage(uia) {
    ; #region agent log
    ; Heuristic: look for common uploading labels in Gemini UI (English + PT-BR).
    ; Scoped to main pane only; do not treat every ProgressBar as upload (false positives).
    try {
        root := FastCopyMode_GetGeminiSearchRoot(uia)
        texts := root.FindAll({ Type: 50020 }) ; Text
        for t in texts {
            name := t.Name
            if (!name)
                continue
            low := StrLower(name)
            if (InStr(low, "open upload file menu"))
                continue
            if (InStr(low, "upload") || InStr(low, "sending") || InStr(low, "carreg") || InStr(low, "enviando")) {
                return true
            }
        }
    } catch {
    }
    return false
    ; #endregion agent log
}

FastCopyMode_WaitForGeminiUploadIdle(uia := "", timeoutMs := 0) {
    ; #region agent log
    global FAST_COPY_GEMINI_UPLOAD_IDLE_MS
    if (timeoutMs <= 0)
        timeoutMs := FAST_COPY_GEMINI_UPLOAD_IDLE_MS
    ; Wait until Gemini is no longer showing an upload-in-progress indicator.
    start := A_TickCount
    if (uia = "") {
        try {
            uia := FastCopyMode_UiaForActiveChrome()
        } catch Error as e {
            FastCopyMode_DebugLog("H6", "Shift keys.ahk:FastCopyMode_WaitForGeminiUploadIdle", "uia_browser_error", Map(
                "msg", SubStr(e.Message, 1, 120)
            ))
            return "uia_fail"
        }
    }

    while ((A_TickCount - start) < timeoutMs) {
        if (!FastCopyMode_GeminiIsUploadingImage(uia))
            return "idle"
        Sleep 150
    }
    return "timeout"
    ; #endregion agent log
}

; --- Gemini + Clip Angel: sequential paste with upload-aware pacing (same prompt focus as Gemini.ahk #!+i) ---

FastCopyMode_IsGeminiHwnd(hwnd) {
    try {
        if (!hwnd)
            return false
        proc := WinGetProcessName("ahk_id " hwnd)
        if (proc != "chrome.exe")
            return false
        title := ""
        try title := WinGetTitle("ahk_id " hwnd)
        if (InStr(title, "gemini", false))
            return true
        ; Fallback: title can be generic; verify by finding the Gemini prompt field via UIA.
        try {
            uia := UIA_Browser("ahk_id " hwnd)
            pf := FindGeminiPromptField(uia)
            if (pf)
                return true
        } catch {
        }
    } catch {
    }
    return false
}

FastCopyMode_IsGeminiForeground() {
    try {
        return FastCopyMode_IsGeminiHwnd(WinGetID("A"))
    } catch {
        return false
    }
}

Gemini_GetUiaForActiveGeminiChrome() {
    try {
        hwnd := WinGetID("A")
        if (!FastCopyMode_IsGeminiHwnd(hwnd))
            return ""
        if (!hwnd)
            return ""
        return UIA_Browser("ahk_id " hwnd)
    } catch {
        return ""
    }
}

; Gemini_FocusPromptSameAsOpenHotkey — see Utils.ahk (shared with Gemini.ahk #!+i).

; After each Clip Angel / screenshot paste: bounded wait until upload UI clears; refocus prompt while uploading.
; timeoutMs: max wait (default FAST_COPY_GEMINI_UPLOAD_IDLE_MS). minNoIndicatorMs: if we never see "uploading",
; wait at least this long before returning idle (approximates legacy fixed tail when no indicator appears).
Gemini_WaitForUploadIdleWithRefocus(uia, timeoutMs := 0, minNoIndicatorMs := 500) {
    global FAST_COPY_GEMINI_UPLOAD_IDLE_MS
    if (!IsObject(uia))
        return "uia_fail"
    if (timeoutMs <= 0)
        timeoutMs := FAST_COPY_GEMINI_UPLOAD_IDLE_MS
    tStart := A_TickCount
    sawUploading := false
    while ((A_TickCount - tStart) < timeoutMs) {
        up := FastCopyMode_GeminiIsUploadingImage(uia)
        if (up)
            sawUploading := true
        if (!up) {
            if (sawUploading || (A_TickCount - tStart) >= minNoIndicatorMs)
                return "idle"
        } else
            FastCopyMode_FocusGeminiPromptField(uia)
        Sleep 150
    }
    return "timeout"
}

; One Clip Angel item per iteration: open once, ^!b with gaps; upload wait between images.
Gemini_PasteFromClipAngelSequential(count, uia := "") {
    if (!IsInteger(count))
        return
    n := Integer(count)
    if (n < 1)
        return
    if (uia = "") {
        uia := Gemini_GetUiaForActiveGeminiChrome()
        if (!IsObject(uia))
            return
    }
    if !ClipAngel_TryAcquireAutomationLock()
        return
    priorHwnd := 0
    try priorHwnd := WinGetID("A")
    catch {
    }
    try {
        StandardLoadingBar_Show("⏳ Pasting clips in Gemini…", BANNER_ACCENT_INTERMEDIATE, { passive: false,
            centerOnHwnd: 0,
            fontSize: 17 })
    } catch {
    }
    try {
        ClipAngel_WaitChordModifiersReleased()
        ClipAngel_ReleaseChordModifiersForSend()
        loop n {
            try StandardLoadingBar_Update("⏳ Pasting clip " A_Index " / " n " …")
            catch {
            }
            Gemini_FocusPromptSameAsOpenHotkey(uia, false)
            if (A_Index = 1) {
                ClipAngel_SendNativeTopItemKeys(priorHwnd)
            } else {
                Sleep(CLIPANGEL_SEQUENTIAL_PASTE_GAP_MS)
                ClipAngel_ReleaseChordModifiersForSend()
                SendInput "^!b"
            }
            ; Brief settle after paste, then condition-based wait for upload UI (efficiency-canon: bounded
            ; wait vs fixed 2.6s). minNoIndicatorMs 2600 preserves ~legacy tail when no upload indicator.
            Sleep 400
            try FastCopyMode_FocusGeminiPromptField(uia)
            try Gemini_WaitForUploadIdleWithRefocus(uia, 4000, 2600)
            try FastCopyMode_FocusGeminiPromptField(uia)
        }
    } finally {
        EnsureClipAngelClosed()
        ClipAngel_RestorePriorFocus(priorHwnd)
        ClipAngel_ReleaseAutomationLock()
        try StandardLoadingBar_Hide(0)
        try {
            if (FastCopyMode_IsGeminiForeground()) {
                if (IsObject(uia)) {
                    try FastCopyMode_FocusGeminiPromptField(uia)
                } else {
                    aw := WinGetID("A")
                    if (aw)
                        FocusGeminiAskFieldForHwnd(aw, false)
                }
            }
        } catch {
        }
    }
}

FastCopyMode_ReleaseHotkeyModifiers() {
    ; Ensure the Win+Alt+Shift hotkey modifiers can't leak into paste keys.
    ; Releasing modifiers does not activate or focus any other window.
    Send "{LWin up}{RWin up}{Alt up}{Shift up}{Ctrl up}"
    Sleep 30
}

FastCopyMode_CaptureScreenshotToQueue(clipboardAlreadyHasImage := false) {
    global gFastCopyScreenshotQueue

    FastCopyMode_DebugLog("H1", "Shift keys.ahk:FastCopyMode_CaptureScreenshotToQueue", "capture_enter", Map(
        "queueLenBefore", gFastCopyScreenshotQueue.Length,
        "hasImageBefore", FastCopyMode_ClipboardHasImage() ? "1" : "0"
    ))

    ; Give the OS a moment to push the screenshot into the clipboard.
    ; Alt+PrintScreen updates the clipboard with an image; rapid captures can overwrite each other
    ; unless we snapshot the clipboard right away.
    ; Wait for a real *image* to appear (not just "clipboard has something").
    ; Some callers (e.g. Win+Shift+S) already waited for the image and just want to snapshot.
    if (!clipboardAlreadyHasImage) {
        try A_Clipboard := ""
        ok := FastCopyMode_WaitForClipboardImage(1500)
    } else {
        ok := FastCopyMode_ClipboardHasImage()
    }
    FastCopyMode_DebugLog("H2", "Shift keys.ahk:FastCopyMode_CaptureScreenshotToQueue", "wait_image_done", Map(
        "ok", ok ? "1" : "0",
        "hasImageAfter", FastCopyMode_ClipboardHasImage() ? "1" : "0"
    ))
    if (!ok)
        return

    try {
        snap := ClipboardAll()
        gFastCopyScreenshotQueue.Push(snap)
        snapSize := ""
        try snapSize := snap.Size
        FastCopyMode_DebugLog("H1", "Shift keys.ahk:FastCopyMode_CaptureScreenshotToQueue", "capture_pushed", Map(
            "queueLenAfter", gFastCopyScreenshotQueue.Length,
            "hasImageAfter", FastCopyMode_ClipboardHasImage() ? "1" : "0",
            "snapType", Type(snap),
            "snapSize", snapSize
        ))
    } catch {
        ; If clipboard snapshot fails, just skip (count still increments).
        FastCopyMode_DebugLog("H2", "Shift keys.ahk:FastCopyMode_CaptureScreenshotToQueue", "clipboardall_failed", Map(
            "queueLenAfter", gFastCopyScreenshotQueue.Length,
            "hasImageAfter", FastCopyMode_ClipboardHasImage() ? "1" : "0"
        ))
    }
}

FastCopyMode_OnWinShiftS() {
    ; Win+Shift+S region snip: count only on successful clipboard image.
    try A_Clipboard := ""
    Send "#+s"
    ok := FastCopyMode_WaitForClipboardImage(30000)
    if (!ok)
        return
    FastCopyMode_OnCopy()
    FastCopyMode_CaptureScreenshotToQueue(true)
}

FastCopyMode_PasteScreenshotQueue(queue) {
    if (!IsObject(queue) || queue.Length < 1)
        return

    total := queue.Length
    try {
        StandardLoadingBar_Show("⏳ Pasting screenshots…", BANNER_ACCENT_INTERMEDIATE, { passive: false, centerOnHwnd: 0,
            fontSize: 17 })
    } catch {
    }

    clipSave := ""
    try clipSave := ClipboardAll()

    try {
        try {
            proc := WinGetProcessName("A")
            cls := WinGetClass("A")
            title := WinGetTitle("A")
        } catch {
            proc := ""
            cls := ""
            title := ""
        }
        isGeminiSession := (proc = "chrome.exe" && InStr(title, "gemini", false))
        cachedGeminiUia := ""
        if (isGeminiSession) {
            try {
                hwndGem := FastCopyMode_GetActiveChromeHwnd()
                if (hwndGem)
                    cachedGeminiUia := UIA_Browser("ahk_id " hwndGem)
            } catch {
                cachedGeminiUia := ""
            }
        }
        FastCopyMode_DebugLog("H3", "Shift keys.ahk:FastCopyMode_PasteScreenshotQueue", "paste_enter", Map(
            "queueLen", queue.Length,
            "hasImageAtEnter", FastCopyMode_ClipboardHasImage() ? "1" : "0",
            "proc", proc,
            "class", cls,
            "title", SubStr(title, 1, 120),
            "isGeminiSession", isGeminiSession ? "1" : "0"
        ))
        for idx, snap in queue {
            try {
                try StandardLoadingBar_Update("⏳ Pasting image " idx " / " total " …")
                catch {
                }
                ; Gemini: ensure prompt is focused BEFORE each paste.
                if (isGeminiSession) {
                    try {
                        focusedPre := FastCopyMode_FocusGeminiPromptField(cachedGeminiUia) ? "1" : "0"
                    } catch {
                        focusedPre := "0"
                    }
                    FastCopyMode_DebugLog("H6", "Shift keys.ahk:FastCopyMode_PasteScreenshotQueue",
                        "gemini_before_paste", Map(
                            "idx", idx,
                            "focusedPrompt", focusedPre
                        ))
                }

                A_Clipboard := snap
                ClipWait 0.6, 1
                snapSize := ""
                try snapSize := snap.Size
                FastCopyMode_DebugLog("H3", "Shift keys.ahk:FastCopyMode_PasteScreenshotQueue", "paste_iter_ready", Map(
                    "idx", idx,
                    "snapType", Type(snap),
                    "snapSize", snapSize,
                    "hasImageNow", FastCopyMode_ClipboardHasImage() ? "1" : "0"
                ))
                cycle := FastCopyMode_SendPasteAndWaitForReadCycle(isGeminiSession)
                FastCopyMode_DebugLog("H4", "Shift keys.ahk:FastCopyMode_PasteScreenshotQueue",
                    "paste_iter_clipboard_cycle", Map(
                        "idx", idx,
                        "cycle", cycle
                    ))

                ; Gemini: condition-based upload wait (early exit when UI idle) vs fixed 2.6s sleep.
                if (isGeminiSession) {
                    Sleep 400
                    try FastCopyMode_FocusGeminiPromptField(cachedGeminiUia)
                    idleStatus := IsObject(cachedGeminiUia) ? Gemini_WaitForUploadIdleWithRefocus(cachedGeminiUia, 5000,
                        800) : "no_uia"
                    if (idleStatus = "no_uia") {
                        Sleep 2600
                        idleStatus := "fallback_fixed_delay"
                    }
                    try FastCopyMode_FocusGeminiPromptField(cachedGeminiUia)
                    FastCopyMode_DebugLog("H6", "Shift keys.ahk:FastCopyMode_PasteScreenshotQueue",
                        "gemini_after_paste", Map(
                            "idx", idx,
                            "uploadIdle", idleStatus
                        ))
                }
            } catch {
                ; continue to next screenshot
                FastCopyMode_DebugLog("H4", "Shift keys.ahk:FastCopyMode_PasteScreenshotQueue", "paste_iter_failed",
                    Map(
                        "idx", idx
                    ))
            }
        }
    } finally {
        if (clipSave != "")
            try A_Clipboard := clipSave
        try StandardLoadingBar_Hide(0)
        FastCopyMode_DebugLog("H3", "Shift keys.ahk:FastCopyMode_PasteScreenshotQueue", "paste_exit", Map(
            "restoredClipboard", clipSave != "" ? "1" : "0"
        ))
    }
}

ExecuteSequentialPaste(actionCount) {
    if (!IsInteger(actionCount))
        return
    n := Integer(actionCount)
    if (n < 1)
        return
    global gFastCopyPasteTargetHwnd
    if !ClipAngel_TryAcquireAutomationLock()
        return
    priorHwnd := gFastCopyPasteTargetHwnd
    if (!priorHwnd || !WinExist("ahk_id " priorHwnd))
        priorHwnd := ClipAngel_ResolvePriorHwnd(0)
    try {
        StandardLoadingBar_Show("⏳ Pasting from clipboard…", BANNER_ACCENT_INTERMEDIATE, { passive: false, centerOnHwnd: 0,
            fontSize: 17 })
    } catch {
    }
    try {
        ClipAngel_WaitChordModifiersReleased()
        ClipAngel_ReleaseChordModifiersForSend()
        loop n {
            try StandardLoadingBar_Update("⏳ Pasting clip " A_Index " / " n " …")
            catch {
            }
            if (A_Index = 1) {
                ClipAngel_SendNativeTopItemKeys(priorHwnd)
            } else {
                Sleep(CLIPANGEL_SEQUENTIAL_PASTE_GAP_MS)
                ClipAngel_ReleaseChordModifiersForSend()
                SendInput "^!b"
            }
            Sleep(CLIPANGEL_INCREMENTAL_PASTE_SETTLE_MS)
        }
    } finally {
        EnsureClipAngelClosed()
        ClipAngel_RestorePriorFocus(priorHwnd)
        ClipAngel_ReleaseAutomationLock()
        try StandardLoadingBar_Hide(0)
    }
}

FastCopyMode_IsActive() {
    global gFastCopyModeActive
    return gFastCopyModeActive
}

FastCopyMode_OnCopy() {
    global gFastCopyCount
    gFastCopyCount += 1
    FastCopyModeBanner_Update(gFastCopyCount)
}

FastCopyMode_PlayCueSound(fileName) {
    if (!IsSoundEnabled())
        return
    path := A_ScriptDir "\sounds\" fileName
    if (!FileExist(path))
        return
    try {
        ScriptSoundPlay(path)
    } catch {
    }
}

FastCopyMode_Start() {
    global gFastCopyModeActive, gFastCopyCount, gFastCopyPasteTargetHwnd, gFastCopyScreenshotQueue
    try {
        gFastCopyPasteTargetHwnd := WinGetID("A")
    } catch {
        gFastCopyPasteTargetHwnd := 0
    }
    gFastCopyCount := 0
    gFastCopyScreenshotQueue := []
    gFastCopyModeActive := true
    try {
        FastCopyModeBanner_Show()
        FastCopyMode_PlayCueSound("fastcopy-start.mp3")
    } catch Error {
        gFastCopyModeActive := false
        ShowCenteredOverlay_Utils("❌ Could not start Fast Copy Mode", 2000, BANNER_ACCENT_ERROR)
    }
}

FastCopyMode_Finish() {
    global gFastCopyModeActive, gFastCopyCount, gFastCopyPasteTargetHwnd
    global gFastCopyScreenshotQueue, gFastCopyLastScreenshotQueue
    count := gFastCopyCount
    shotCount := IsObject(gFastCopyScreenshotQueue) ? gFastCopyScreenshotQueue.Length : 0
    try {
        FastCopyModeBanner_Hide()
    } finally {
        gFastCopyModeActive := false
        gFastCopyCount := 0
    }
    try {
        ; Paste exclusively into the *currently active* window without activating anything else.
        FastCopyMode_ReleaseHotkeyModifiers()
        if (count > 0) {
            FastCopyMode_PlayCueSound("fastcopy-finish.mp3")
            FastCopyMode_DebugLog("H5", "Shift keys.ahk:FastCopyMode_Finish", "finish_paste_start", Map(
                "count", count,
                "shotCount", shotCount
            ))
            if (shotCount > 0) {
                FastCopyMode_PasteScreenshotQueue(gFastCopyScreenshotQueue)
                ; Save for hold-to-repeat behavior.
                gFastCopyLastScreenshotQueue := gFastCopyScreenshotQueue.Clone()
            } else {
                gFastCopyLastScreenshotQueue := []
            }

            remaining := count - shotCount
            FastCopyMode_DebugLog("H5", "Shift keys.ahk:FastCopyMode_Finish", "finish_remaining", Map(
                "remaining", remaining
            ))
            if (remaining > 0) {
                ; Non-screenshot copies: Clip Angel sequential paste (Gemini uses upload-aware loop).
                if (FastCopyMode_IsGeminiForeground())
                    Gemini_PasteFromClipAngelSequential(remaining)
                else
                    ExecuteSequentialPaste(remaining)
            }

            global gFastCopyLastSuccessfulCount
            gFastCopyLastSuccessfulCount := count
        } else
            ShowCenteredOverlay_Utils("⚠ No copies recorded — nothing to paste", 2500, BANNER_ACCENT_INTERMEDIATE)
    } catch Error as e {
        ShowCenteredOverlay_Utils("❌ Fast Copy Mode: " SubStr(e.Message, 1, 80), 2500, BANNER_ACCENT_ERROR)
    }
}

FastCopyMode_RepeatLastPaste() {
    global gFastCopyLastSuccessfulCount
    global gFastCopyLastScreenshotQueue
    if (gFastCopyLastSuccessfulCount < 1) {
        ShowCenteredOverlay_Utils("⚠ No previous Fast Copy paste to repeat", 2500, BANNER_ACCENT_INTERMEDIATE)
        return
    }
    try {
        ; Repeat paste into the *currently active* window without activating anything else.
        FastCopyMode_ReleaseHotkeyModifiers()
        FastCopyMode_PlayCueSound("fastcopy-finish.mp3")
        if (IsObject(gFastCopyLastScreenshotQueue) && gFastCopyLastScreenshotQueue.Length > 0) {
            FastCopyMode_PasteScreenshotQueue(gFastCopyLastScreenshotQueue)
            remaining := gFastCopyLastSuccessfulCount - gFastCopyLastScreenshotQueue.Length
            if (remaining > 0) {
                if (FastCopyMode_IsGeminiForeground())
                    Gemini_PasteFromClipAngelSequential(remaining)
                else
                    ExecuteSequentialPaste(remaining)
            }
        } else {
            if (FastCopyMode_IsGeminiForeground())
                Gemini_PasteFromClipAngelSequential(gFastCopyLastSuccessfulCount)
            else
                ExecuteSequentialPaste(gFastCopyLastSuccessfulCount)
        }
    } catch Error as e {
        ShowCenteredOverlay_Utils("❌ Repeat paste: " SubStr(e.Message, 1, 80), 2500, BANNER_ACCENT_ERROR)
    }
}

#!+j:: {
    global gFastCopyModeActive, FAST_COPY_HOLD_REPEAT_MS
    if (gFastCopyModeActive) {
        FastCopyMode_Finish()
        return
    }
    pressTime := A_TickCount
    KeyWait "j", "T1"
    holdTime := A_TickCount - pressTime
    if (holdTime >= FAST_COPY_HOLD_REPEAT_MS)
        FastCopyMode_RepeatLastPaste()
    else
        FastCopyMode_Start()
}

; Win+Alt+Shift+L — Outlook Copilot shortcut modal (1–9). Global: works from any app; actions activate Outlook.
#!+l:: {
    SelectOutlookCopilotShortcut()
}

#HotIf FastCopyMode_IsActive()
~^c:: FastCopyMode_OnCopy()
~PrintScreen:: FastCopyMode_OnCopy()
~!PrintScreen:: {
    FastCopyMode_OnCopy()
    FastCopyMode_CaptureScreenshotToQueue()
}
$#+s:: FastCopyMode_OnWinShiftS()
#HotIf
