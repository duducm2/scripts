; =============================================================================
; Utils module: standard_loading_bar.ahk
; Standard loading bar show/update/hide lifecycle
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; Standard loading bar (monitor-aware, show/update/hide lifecycle)
; Use for long-running shortcuts; replace ad-hoc banners/overlays with this.
; Supports passive (text-only) mode and ShowWithKeys for letter-keystroke commands.
; Semantic accent colors (colorblind accessibility): border only; background stays dark.
; =============================================================================
global BANNER_ACCENT_SUCCESS := "27AE60"      ; Dark green: positive / success
global BANNER_ACCENT_ERROR := "C0392B"        ; Red: negative / error
global BANNER_ACCENT_INTERMEDIATE := "F1C40F" ; Yellow: loading, actionable, neutral
global BANNER_ACCENT_INFO := "2980B9"         ; Blue: distinct from green/yellow for color vision (info / alternate mode)
; Monitor blackout countdown (Study Topic / focus dwell): unmistakable vs generic loading banners
global BANNER_BLACKOUT_PANEL := "4A148C"       ; Purple panel fill
global BANNER_BLACKOUT_BORDER := "FF9800"    ; Orange border + progress strip
global g_StandardLoadingBarGui := 0
global g_StandardLoadingBarValue := 0
global g_StandardLoadingBarIsKeysOverlay := false
global g_StandardLoadingBarKeysHotkeys := []
global g_StandardLoadingBarKeysTimeoutTimer := ""
global g_StandardLoadingBarBorderGui := 0
global g_StandardLoadingBarTrackTimer := ""
global g_StandardLoadingBarLastForegroundMonitorIdx := 0
global g_StandardLoadingBarTimedProgressTimer := ""
global g_StandardLoadingBarTimedProgressStartTick := 0
global g_StandardLoadingBarTimedProgressDurationMs := 0
global D2C_SUBMIT_MENU_TIMEOUT_MS := 5000
; Keys overlay: same escape stack as Handy AI model selector (#!+C) — *Escape alone misses when the bar is NA/no focus.
global g_StandardLoadingBarKeysEscapeUserCb := ""
global g_StandardLoadingBarKeysEscapeActive := false
global g_StandardLoadingBarEscPollPrev := false
global g_StandardLoadingBarKeysPollTimer := ""
global g_StandardLoadingBarKeysPollPrev := Map()
global g_StandardLoadingBarKeysPollCallbacks := Map()
; Replaceable delayed-hide + hard-max watchdog (anonymous SetTimer fat-arrows do not replace each other).
global g_StandardLoadingBarHideTimerArmed := false
global g_StandardLoadingBarForceHideTimerArmed := false
global STANDARD_LOADING_BAR_FORCE_HIDE_MS := 5000

; Return work area { left, top, right, bottom } for the monitor containing hwnd, or "" on failure.
GetWorkAreaForWindow_StandardBar(hwnd) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return ""
    try {
        WinGetPos(&winX, &winY, &winW, &winH, "ahk_id " hwnd)
        centerX := winX + winW / 2
        centerY := winY + winH / 2
        n := MonitorGetCount()
        loop n {
            MonitorGet(A_Index, &L, &T, &R, &B)
            if (centerX >= L && centerX < R && centerY >= T && centerY < B) {
                MonitorGetWorkArea(A_Index, &wLeft, &wTop, &wRight, &wBottom)
                return { left: wLeft, top: wTop, right: wRight, bottom: wBottom }
            }
        }
    } catch {
    }
    return ""
}

GetActiveMonitorWorkArea_StandardBar(&left, &top, &right, &bottom) {
    left := top := 0
    right := A_ScreenWidth
    bottom := A_ScreenHeight
    activeWin := 0
    try activeWin := WinGetID("A")
    catch
        activeWin := 0
    MonitorGetWorkArea(1, &monitorLeft, &monitorTop, &monitorRight, &monitorBottom)
    mLeft := monitorLeft
    mTop := monitorTop
    mRight := monitorRight
    mBottom := monitorBottom
    if (activeWin && activeWin != 0) {
        rect := Buffer(16, 0)
        if (DllCall("GetWindowRect", "ptr", activeWin, "ptr", rect)) {
            winLeft := NumGet(rect, 0, "int")
            winTop := NumGet(rect, 4, "int")
            winRight := NumGet(rect, 8, "int")
            winBottom := NumGet(rect, 12, "int")
            centerX := winLeft + (winRight - winLeft) // 2
            centerY := winTop + (winBottom - winTop) // 2
            monitorCount := MonitorGetCount()
            loop monitorCount {
                idx := A_Index
                MonitorGetWorkArea(idx, &l, &t, &r, &b)
                if (centerX >= l && centerX <= r && centerY >= t && centerY <= b) {
                    mLeft := l
                    mTop := t
                    mRight := r
                    mBottom := b
                    break
                }
            }
        }
    }
    left := mLeft
    top := mTop
    right := mRight
    bottom := mBottom
}

; 1-based monitor index for the monitor containing the center of the foreground window; 1 if unknown.
GetMonitorIndexForForeground_StandardBar() {
    activeWin := 0
    try activeWin := WinGetID("A")
    catch
        activeWin := 0
    if (!activeWin)
        return 1
    rect := Buffer(16, 0)
    if (!DllCall("GetWindowRect", "ptr", activeWin, "ptr", rect))
        return 1
    winLeft := NumGet(rect, 0, "int")
    winTop := NumGet(rect, 4, "int")
    winRight := NumGet(rect, 8, "int")
    winBottom := NumGet(rect, 12, "int")
    centerX := winLeft + (winRight - winLeft) // 2
    centerY := winTop + (winBottom - winTop) // 2
    monitorCount := MonitorGetCount()
    loop monitorCount {
        idx := A_Index
        MonitorGetWorkArea(idx, &l, &t, &r, &b)
        if (centerX >= l && centerX <= r && centerY >= t && centerY <= b)
            return idx
    }
    return 1
}

StandardLoadingBar_StopActiveMonitorTracking() {
    global g_StandardLoadingBarTrackTimer, g_StandardLoadingBarLastForegroundMonitorIdx
    try SetTimer(g_StandardLoadingBarTrackTimer, 0)
    catch {
    }
    g_StandardLoadingBarTrackTimer := ""
    g_StandardLoadingBarLastForegroundMonitorIdx := 0
}

StandardLoadingBar_RepositionToActiveMonitor() {
    global g_StandardLoadingBarGui, g_StandardLoadingBarBorderGui
    if !IsObject(g_StandardLoadingBarGui)
        return
    GetActiveMonitorWorkArea_StandardBar(&ml, &mt, &mr, &mb)
    monitorWidth := mr - ml
    try {
        g_StandardLoadingBarGui.GetPos(, , &gw, &gh)
    } catch {
        return
    }
    guiX := Round(ml + (monitorWidth - gw) / 2)
    if (guiX < ml)
        guiX := ml
    if (guiX + gw > mr)
        guiX := mr - gw
    guiY := mt + 40
    if (IsObject(g_StandardLoadingBarBorderGui)) {
        borderWidth := 6
        try {
            g_StandardLoadingBarBorderGui.Move(guiX - borderWidth, guiY - borderWidth, gw + 2 * borderWidth, gh + 2 *
                borderWidth)
        } catch {
        }
    }
    try {
        g_StandardLoadingBarGui.Move(guiX, guiY)
        hwnd := g_StandardLoadingBarGui.Hwnd
        if (hwnd)
            DllCall("SetWindowPos", "Ptr", hwnd, "Ptr", 0, "Int", guiX, "Int", guiY, "Int", 0, "Int", 0, "UInt", 0x0015
            )
    } catch {
    }
}

StandardLoadingBar_TrackActiveMonitorTick() {
    global g_StandardLoadingBarGui, g_StandardLoadingBarLastForegroundMonitorIdx
    if !IsObject(g_StandardLoadingBarGui) {
        StandardLoadingBar_StopActiveMonitorTracking()
        return
    }
    newIdx := GetMonitorIndexForForeground_StandardBar()
    if (newIdx != g_StandardLoadingBarLastForegroundMonitorIdx) {
        g_StandardLoadingBarLastForegroundMonitorIdx := newIdx
        StandardLoadingBar_RepositionToActiveMonitor()
    }
}

StandardLoadingBar_Show(state := "Working...", barColor := BANNER_ACCENT_INTERMEDIATE, options := "") {
    global g_StandardLoadingBarGui, g_StandardLoadingBarValue, g_StandardLoadingBarIsKeysOverlay,
        g_StandardLoadingBarBorderGui
    ; Capture prior HWNDs so a failed Hide cannot leave an AlwaysOnTop orphan under a new banner.
    oldGuiHwnd := 0
    oldBorderHwnd := 0
    try {
        if IsObject(g_StandardLoadingBarGui)
            oldGuiHwnd := Integer(g_StandardLoadingBarGui.Hwnd)
    } catch {
    }
    try {
        if IsObject(g_StandardLoadingBarBorderGui)
            oldBorderHwnd := Integer(g_StandardLoadingBarBorderGui.Hwnd)
    } catch {
    }
    try StandardLoadingBar_CloseKeysOverlay()
    ; Hide(0) also cancels replaceable hide / force-hide watchdogs from a prior banner.
    try StandardLoadingBar_Hide(0)
    for hwndKill in [oldGuiHwnd, oldBorderHwnd] {
        if (hwndKill && WinExist("ahk_id " hwndKill)) {
            try WinHide("ahk_id " hwndKill)
            catch {
            }
            try DllCall("user32\DestroyWindow", "ptr", hwndKill)
            catch {
            }
        }
    }
    g_StandardLoadingBarGui := 0
    g_StandardLoadingBarBorderGui := 0
    passive := options && options.HasProp("passive") && options.passive
    centerOnHwnd := options && options.HasProp("centerOnHwnd") ? options.centerOnHwnd : 0
    textWidth := options && options.HasProp("textWidth") ? options.textWidth : 0
    fontSize := options && options.HasProp("fontSize") ? options.fontSize : 17
    alpha := options && options.HasProp("alpha") ? options.alpha : 235
    passiveBgColor := options && options.HasProp("passiveBgColor") ? options.passiveBgColor : ""
    noBorder := options && options.HasProp("noBorder") ? options.noBorder : false
    promptKeys := options && options.HasProp("promptKeys") ? options.promptKeys : ""
    trackActiveMonitor := options && options.HasProp("trackActiveMonitor") && options.trackActiveMonitor
    manualProgress := options && options.HasProp("manualProgress") && options.manualProgress
    overlayBgColor := options && options.HasProp("overlayBgColor") && options.overlayBgColor != "" ? options.overlayBgColor :
        "1E1E2E"

    if (centerOnHwnd) {
        workArea := GetWorkAreaForWindow_StandardBar(centerOnHwnd)
        if (workArea != "") {
            ml := workArea.left
            mt := workArea.top
            mr := workArea.right
            mb := workArea.bottom
        } else
            GetActiveMonitorWorkArea_StandardBar(&ml, &mt, &mr, &mb)
    } else
        GetActiveMonitorWorkArea_StandardBar(&ml, &mt, &mr, &mb)
    monitorWidth := mr - ml
    monitorHeight := mb - mt
    barWidth := textWidth > 0 ? textWidth : Min(900, Max(360, Floor(monitorWidth * 0.6)))
    overlayGui := Gui("+AlwaysOnTop -Caption +ToolWindow -DPIScale")
    overlayGui.BackColor := overlayBgColor
    overlayGui.MarginX := 16
    overlayGui.MarginY := 10
    overlayGui.SetFont("s" . fontSize . " cFFFFFF", "Segoe UI")
    overlayGui.Add("Text", "w" . barWidth . (passive ? " Wrap Center" : " Center"), state)
    if (promptKeys != "") {
        overlayGui.SetFont("s" . fontSize . " cFFFFFF", "Segoe UI")
        ; Wrap so long key strips (e.g. Gemini Copy response? + [F]) are not clipped at fixed width.
        overlayGui.Add("Text", "xm w" . barWidth . " Center Wrap", promptKeys)
    }
    if (!passive) {
        progressOpts := "w" . barWidth . " h10 c" . barColor . " Background45475A Smooth vOverlayProg"
        overlayGui.Add("Progress", progressOpts, 0)
    }
    overlayGui.Show("AutoSize Hide")
    overlayGui.GetPos(, , &gw, &gh)
    guiX := Round(ml + (monitorWidth - gw) / 2)
    if (guiX < ml)
        guiX := ml
    if (guiX + gw > mr)
        guiX := mr - gw
    guiY := mt + 40

    ; Create border frame behind the overlay for visibility (optional; skip when noBorder to show a single banner). Accent color when passiveBgColor set, else yellow.
    if (!noBorder) {
        borderWidth := 6
        try {
            if IsObject(g_StandardLoadingBarBorderGui)
                g_StandardLoadingBarBorderGui.Destroy()
        } catch {
        }
        borderGui := Gui("+AlwaysOnTop -Caption +ToolWindow -DPIScale")
        borderGui.BackColor := (passiveBgColor != "") ? passiveBgColor : BANNER_ACCENT_INTERMEDIATE
        borderGui.Show("NA x" . (guiX - borderWidth) . " y" . (guiY - borderWidth) . " w" . (gw + 2 * borderWidth) .
        " h" .
        (gh + 2 * borderWidth))
        g_StandardLoadingBarBorderGui := borderGui
    } else {
        try {
            if IsObject(g_StandardLoadingBarBorderGui)
                g_StandardLoadingBarBorderGui.Destroy()
        } catch {
        }
        g_StandardLoadingBarBorderGui := 0
    }
    overlayGui.Show("x" . guiX . " y" . guiY . " NA")
    try {
        hwnd := overlayGui.Hwnd
        if (hwnd)
            DllCall("SetWindowPos", "Ptr", hwnd, "Ptr", 0, "Int", guiX, "Int", guiY, "Int", 0, "Int", 0, "UInt", 0x0015
            )
    }
    WinSetTransparent(alpha, overlayGui)
    g_StandardLoadingBarGui := overlayGui
    g_StandardLoadingBarValue := 0
    if (!passive && !manualProgress)
        SetTimer(StandardLoadingBar_Tick, 40)
    if (trackActiveMonitor) {
        StandardLoadingBar_StopActiveMonitorTracking()
        g_StandardLoadingBarLastForegroundMonitorIdx := GetMonitorIndexForForeground_StandardBar()
        g_StandardLoadingBarTrackTimer := SetTimer(StandardLoadingBar_TrackActiveMonitorTick, 115)
    }
}

StandardLoadingBar_Tick() {
    global g_StandardLoadingBarGui, g_StandardLoadingBarValue
    if !IsObject(g_StandardLoadingBarGui) {
        SetTimer(StandardLoadingBar_Tick, 0)
        return
    }
    try {
        g_StandardLoadingBarValue += 4
        if (g_StandardLoadingBarValue > 100)
            g_StandardLoadingBarValue := 0
        g_StandardLoadingBarGui["OverlayProg"].Value := g_StandardLoadingBarValue
    } catch {
        SetTimer(StandardLoadingBar_Tick, 0)
    }
}

StandardLoadingBar_StopTimedProgress() {
    global g_StandardLoadingBarTimedProgressTimer, g_StandardLoadingBarTimedProgressStartTick,
        g_StandardLoadingBarTimedProgressDurationMs
    try SetTimer(g_StandardLoadingBarTimedProgressTimer, 0)
    catch {
    }
    g_StandardLoadingBarTimedProgressTimer := ""
    g_StandardLoadingBarTimedProgressStartTick := 0
    g_StandardLoadingBarTimedProgressDurationMs := 0
}

StandardLoadingBar_SetProgressValue(value) {
    global g_StandardLoadingBarGui, g_StandardLoadingBarValue
    if !IsObject(g_StandardLoadingBarGui)
        return
    clamped := Max(0, Min(100, Round(value)))
    g_StandardLoadingBarValue := clamped
    try g_StandardLoadingBarGui["OverlayProg"].Value := clamped
    catch {
    }
}

StandardLoadingBar_StartTimedProgress(durationMs) {
    global g_StandardLoadingBarTimedProgressTimer, g_StandardLoadingBarTimedProgressStartTick,
        g_StandardLoadingBarTimedProgressDurationMs
    StandardLoadingBar_StopTimedProgress()
    SetTimer(StandardLoadingBar_Tick, 0)
    if (durationMs <= 0) {
        StandardLoadingBar_SetProgressValue(100)
        return
    }
    g_StandardLoadingBarTimedProgressStartTick := A_TickCount
    g_StandardLoadingBarTimedProgressDurationMs := durationMs
    StandardLoadingBar_SetProgressValue(0)
    g_StandardLoadingBarTimedProgressTimer := SetTimer(StandardLoadingBar_TimedProgressTick, 40)
}

StandardLoadingBar_TimedProgressTick() {
    global g_StandardLoadingBarGui, g_StandardLoadingBarTimedProgressStartTick,
        g_StandardLoadingBarTimedProgressDurationMs
    if !IsObject(g_StandardLoadingBarGui) {
        StandardLoadingBar_StopTimedProgress()
        return
    }
    durationMs := g_StandardLoadingBarTimedProgressDurationMs
    if (durationMs <= 0) {
        StandardLoadingBar_SetProgressValue(100)
        StandardLoadingBar_StopTimedProgress()
        return
    }
    elapsedMs := A_TickCount - g_StandardLoadingBarTimedProgressStartTick
    if (elapsedMs >= durationMs) {
        StandardLoadingBar_SetProgressValue(100)
        StandardLoadingBar_StopTimedProgress()
        return
    }
    StandardLoadingBar_SetProgressValue((elapsedMs * 100.0) / durationMs)
}

StandardLoadingBar_Update(state := "", barColor := "") {
    global g_StandardLoadingBarGui
    if !IsObject(g_StandardLoadingBarGui)
        return
    try {
        if (state != "" && g_StandardLoadingBarGui.Controls.Length > 0)
            g_StandardLoadingBarGui.Controls[1].Text := state
    } catch {
    }
    if (barColor != "") {
        try
            g_StandardLoadingBarGui["OverlayProg"].Opt("c" . barColor)
        catch {
        }
    }
}

StandardLoadingBar_CancelHideTimers() {
    global g_StandardLoadingBarHideTimerArmed, g_StandardLoadingBarForceHideTimerArmed
    try SetTimer(StandardLoadingBar_HideNow, 0)
    catch {
    }
    try SetTimer(StandardLoadingBar_ForceHideWatchdog, 0)
    catch {
    }
    g_StandardLoadingBarHideTimerArmed := false
    g_StandardLoadingBarForceHideTimerArmed := false
}

; Named timer target so delayed Hide replaces prior schedules (not fat-arrow one-shots).
StandardLoadingBar_HideNow(*) {
    global g_StandardLoadingBarHideTimerArmed
    g_StandardLoadingBarHideTimerArmed := false
    StandardLoadingBar_Hide(0)
}

; Hard-max dismiss for any banner that missed its normal hide (passive or keys).
StandardLoadingBar_ForceHideWatchdog(*) {
    global g_StandardLoadingBarForceHideTimerArmed, g_StandardLoadingBarGui, g_StandardLoadingBarIsKeysOverlay
    g_StandardLoadingBarForceHideTimerArmed := false
    if (g_StandardLoadingBarIsKeysOverlay) {
        try StandardLoadingBar_CloseKeysOverlay()
        catch {
        }
        return
    }
    if (!IsObject(g_StandardLoadingBarGui))
        return
    StandardLoadingBar_Hide(0)
}

; Arm (or replace) a force-hide watchdog; does not shorten a shorter normal Hide(duration).
StandardLoadingBar_ArmForceHide(maxMs := 0) {
    global g_StandardLoadingBarForceHideTimerArmed, STANDARD_LOADING_BAR_FORCE_HIDE_MS
    if (maxMs < 1)
        maxMs := STANDARD_LOADING_BAR_FORCE_HIDE_MS
    try SetTimer(StandardLoadingBar_ForceHideWatchdog, 0)
    catch {
    }
    SetTimer(StandardLoadingBar_ForceHideWatchdog, -maxMs)
    g_StandardLoadingBarForceHideTimerArmed := true
}

StandardLoadingBar_Hide(delayMs := 0) {
    global g_StandardLoadingBarGui, g_StandardLoadingBarValue, g_StandardLoadingBarIsKeysOverlay,
        g_StandardLoadingBarBorderGui, g_StandardLoadingBarHideTimerArmed, STANDARD_LOADING_BAR_FORCE_HIDE_MS
    if (delayMs > 0) {
        try SetTimer(StandardLoadingBar_HideNow, 0)
        catch {
        }
        SetTimer(StandardLoadingBar_HideNow, -delayMs)
        g_StandardLoadingBarHideTimerArmed := true
        ; Hard ceiling so a missed HideNow / failed Destroy cannot leave the banner for minutes.
        forceMs := delayMs + 2000
        if (forceMs < STANDARD_LOADING_BAR_FORCE_HIDE_MS)
            forceMs := STANDARD_LOADING_BAR_FORCE_HIDE_MS
        StandardLoadingBar_ArmForceHide(forceMs)
        return
    }
    ; Cancel HideNow only first — keep force-hide until destroy succeeds (re-armed on failure).
    try SetTimer(StandardLoadingBar_HideNow, 0)
    catch {
    }
    g_StandardLoadingBarHideTimerArmed := false
    StandardLoadingBar_StopActiveMonitorTracking()
    StandardLoadingBar_StopTimedProgress()
    if (g_StandardLoadingBarIsKeysOverlay) {
        StandardLoadingBar_CloseKeysOverlay()
        return
    }
    SetTimer(StandardLoadingBar_Tick, 0)
    guiObj := g_StandardLoadingBarGui
    borderObj := g_StandardLoadingBarBorderGui
    guiHwnd := 0
    borderHwnd := 0
    try {
        if IsObject(guiObj)
            guiHwnd := Integer(guiObj.Hwnd)
    } catch {
    }
    try {
        if IsObject(borderObj)
            borderHwnd := Integer(borderObj.Hwnd)
    } catch {
    }
    try {
        if IsObject(guiObj)
            guiObj.Destroy()
    } catch {
    }
    try {
        if IsObject(borderObj)
            borderObj.Destroy()
    } catch {
    }
    ; Fallback if Gui.Destroy left an AlwaysOnTop window alive.
    if (guiHwnd && WinExist("ahk_id " guiHwnd)) {
        try WinHide("ahk_id " guiHwnd)
        catch {
        }
        try DllCall("user32\DestroyWindow", "ptr", guiHwnd)
        catch {
        }
    }
    if (borderHwnd && WinExist("ahk_id " borderHwnd)) {
        try WinHide("ahk_id " borderHwnd)
        catch {
        }
        try DllCall("user32\DestroyWindow", "ptr", borderHwnd)
        catch {
        }
    }
    stillLive := (guiHwnd && WinExist("ahk_id " guiHwnd))
    if (stillLive) {
        ; Keep references and retry — do not orphan an AlwaysOnTop window.
        g_StandardLoadingBarGui := guiObj
        g_StandardLoadingBarBorderGui := borderObj
        StandardLoadingBar_ArmForceHide(1000)
        return
    }
    try SetTimer(StandardLoadingBar_ForceHideWatchdog, 0)
    catch {
    }
    global g_StandardLoadingBarForceHideTimerArmed
    g_StandardLoadingBarForceHideTimerArmed := false
    g_StandardLoadingBarGui := 0
    g_StandardLoadingBarValue := 0
    g_StandardLoadingBarBorderGui := 0
}

; Unregister keys and timeout timer for the "ShowWithKeys" overlay, then hide. Idempotent.
StandardLoadingBar_CloseKeysOverlay() {
    global g_StandardLoadingBarKeysHotkeys, g_StandardLoadingBarKeysTimeoutTimer
    global g_StandardLoadingBarGui, g_StandardLoadingBarValue, g_StandardLoadingBarIsKeysOverlay,
        g_StandardLoadingBarBorderGui
    global g_StandardLoadingBarKeysEscapeActive, g_OnEscapePressed, g_StandardLoadingBarKeysEscapeUserCb
    StandardLoadingBar_CancelHideTimers()
    g_StandardLoadingBarIsKeysOverlay := false
    StandardLoadingBar_StopKeysSelectionPoll()
    hadEscStack := g_StandardLoadingBarKeysEscapeActive
    g_StandardLoadingBarKeysEscapeActive := false
    g_StandardLoadingBarKeysEscapeUserCb := ""
    g_StandardLoadingBarEscPollPrev := false
    try SetTimer(StandardLoadingBar_KeysEscapePoll, 0)
    catch {
    }
    try SetTimer(g_StandardLoadingBarKeysTimeoutTimer, 0)
    catch {
    }
    g_StandardLoadingBarKeysTimeoutTimer := ""
    StandardLoadingBar_StopActiveMonitorTracking()
    StandardLoadingBar_StopTimedProgress()
    for key in g_StandardLoadingBarKeysHotkeys {
        try Hotkey(key, "Off")
        catch {
        }
    }
    g_StandardLoadingBarKeysHotkeys := []
    if (hadEscStack) {
        g_OnEscapePressed := ""
        Utils_EnsureGlobalEscapeHotkey()
    }
    SetTimer(StandardLoadingBar_Tick, 0)
    try {
        if IsObject(g_StandardLoadingBarGui)
            g_StandardLoadingBarGui.Destroy()
    } catch {
    }
    g_StandardLoadingBarGui := 0
    g_StandardLoadingBarValue := 0
    try {
        if IsObject(g_StandardLoadingBarBorderGui)
            g_StandardLoadingBarBorderGui.Destroy()
    } catch {
    }
    g_StandardLoadingBarBorderGui := 0
}

; Return the Escape cancel callback from keyCallbacks, if any (used for robust Esc handling).
StandardLoadingBar_EscapeCallbackFromKeyCallbacks(keyCallbacks) {
    if (keyCallbacks is Map) {
        if keyCallbacks.Has("*Escape")
            return keyCallbacks["*Escape"]
        if keyCallbacks.Has("Escape")
            return keyCallbacks["Escape"]
        return ""
    }
    try {
        for kn, cb in keyCallbacks {
            if (!cb)
                continue
            k := StrLower(Trim(kn))
            if (k = "*escape" || k = "escape")
                return cb
        }
    } catch {
    }
    return ""
}

StandardLoadingBar_KeysSelectionModifiersDown() {
    try {
        ; Chord modifiers only (not Shift): WaitForTriggerKeyRelease already waits for Shift; including Shift here
        ; blocked poll edges while Shift was still down after #!+… chords.
        return GetKeyState("LWin", "P") || GetKeyState("RWin", "P") || GetKeyState("Ctrl", "P") || GetKeyState("Alt",
            "P")
    } catch {
        return false
    }
}

; AHK v2: do not use keyName >= "0" on letters — throws "Expected a Number but got a String".
StandardLoadingBar_IsDigitKey(keyName) {
    if (StrLen(keyName) != 1)
        return false
    o := Ord(keyName)
    return o >= Ord("0") && o <= Ord("9")
}

StandardLoadingBar_KeysSelectionKeyDown(keyName) {
    try {
        if (StandardLoadingBar_IsDigitKey(keyName))
            return GetKeyState(keyName, "P") || GetKeyState("Numpad" . keyName, "P")
        return GetKeyState(keyName, "P")
    } catch {
        return false
    }
}

StandardLoadingBar_StopKeysSelectionPoll() {
    global g_StandardLoadingBarKeysPollTimer, g_StandardLoadingBarKeysPollPrev, g_StandardLoadingBarKeysPollCallbacks
    try SetTimer(g_StandardLoadingBarKeysPollTimer, 0)
    catch {
    }
    g_StandardLoadingBarKeysPollTimer := ""
    g_StandardLoadingBarKeysPollPrev := Map()
    g_StandardLoadingBarKeysPollCallbacks := Map()
}

StandardLoadingBar_StartKeysSelectionPoll(keyCallbacks) {
    global g_StandardLoadingBarKeysPollTimer, g_StandardLoadingBarKeysPollPrev, g_StandardLoadingBarKeysPollCallbacks
    StandardLoadingBar_StopKeysSelectionPoll()
    g_StandardLoadingBarKeysPollCallbacks := Map()
    g_StandardLoadingBarKeysPollPrev := Map()
    try {
        for keyName, cb in keyCallbacks {
            if (!cb)
                continue
            knL := StrLower(Trim(keyName))
            if (knL = "escape" || knL = "*escape")
                continue
            g_StandardLoadingBarKeysPollCallbacks[keyName] := cb
            g_StandardLoadingBarKeysPollPrev[keyName] := StandardLoadingBar_KeysSelectionKeyDown(keyName)
        }
    } catch {
    }
    if (g_StandardLoadingBarKeysPollCallbacks.Count = 0)
        return
    g_StandardLoadingBarKeysPollTimer := SetTimer(StandardLoadingBar_KeysSelectionPoll, 50)
}

; Edge-triggered poll: survives Hotkey("1") conflicts from other scripts (see CursorTransfer selector loop).
StandardLoadingBar_KeysSelectionPoll() {
    global g_StandardLoadingBarIsKeysOverlay, g_StandardLoadingBarKeysPollPrev, g_StandardLoadingBarKeysPollCallbacks
    if (!g_StandardLoadingBarIsKeysOverlay) {
        StandardLoadingBar_StopKeysSelectionPoll()
        return
    }
    if (StandardLoadingBar_KeysSelectionModifiersDown()) {
        for keyName, cb in g_StandardLoadingBarKeysPollCallbacks {
            if (!cb)
                continue
            g_StandardLoadingBarKeysPollPrev[keyName] := StandardLoadingBar_KeysSelectionKeyDown(keyName)
        }
        return
    }
    for keyName, cb in g_StandardLoadingBarKeysPollCallbacks {
        if (!cb)
            continue
        isDown := StandardLoadingBar_KeysSelectionKeyDown(keyName)
        wasDown := g_StandardLoadingBarKeysPollPrev.Has(keyName) ? g_StandardLoadingBarKeysPollPrev[keyName] : false
        g_StandardLoadingBarKeysPollPrev[keyName] := isDown
        if (isDown && !wasDown)
            StandardLoadingBar_KeyWrapper(keyName, cb)
    }
}

; After chord release, wait until digit selection keys are up so poll arming does not treat a held key as wasDown.
StandardLoadingBar_WaitForSelectionKeysRelease(keyCallbacks) {
    try {
        for keyName, cb in keyCallbacks {
            if (!cb)
                continue
            knL := StrLower(Trim(keyName))
            if (knL = "escape" || knL = "*escape")
                continue
            if (StandardLoadingBar_IsDigitKey(keyName)) {
                while StandardLoadingBar_KeysSelectionKeyDown(keyName)
                    KeyWait keyName
            }
        }
    } catch {
    }
}

; After a chord hotkey (e.g. #!+w), wait until Win/Ctrl/Alt/Shift and the trigger key are released
; so *1 / *2 selection hotkeys are not swallowed while modifiers are still held.
StandardLoadingBar_WaitForTriggerKeyRelease() {
    try {
        if (A_ThisHotkey = "")
            return
        th := A_ThisHotkey
        if InStr(th, "#") {
            try KeyWait "LWin"
            try KeyWait "RWin"
        }
        if InStr(th, "^")
            try KeyWait "Ctrl"
        if InStr(th, "!")
            try KeyWait "Alt"
        if InStr(th, "+")
            try KeyWait "Shift"
        hk := th
        hk := StrReplace(hk, "+", "")
        hk := StrReplace(hk, "^", "")
        hk := StrReplace(hk, "!", "")
        hk := StrReplace(hk, "#", "")
        if (StrLen(hk) = 1)
            KeyWait hk
    } catch {
    }
}

; Poll Esc — fallback when $*Escape / g_OnEscapePressed miss (same idea as ShowAiModelSelector escape poll).
StandardLoadingBar_KeysEscapePoll() {
    global g_StandardLoadingBarKeysEscapeActive, g_StandardLoadingBarIsKeysOverlay, g_StandardLoadingBarEscPollPrev
    if (!g_StandardLoadingBarKeysEscapeActive || !g_StandardLoadingBarIsKeysOverlay) {
        try SetTimer(StandardLoadingBar_KeysEscapePoll, 0)
        catch {
        }
        return
    }
    escDown := GetKeyState("Escape", "P") || (DllCall("user32\GetAsyncKeyState", "int", 0x1B) & 0x8000)
    if escDown {
        if !g_StandardLoadingBarEscPollPrev {
            g_StandardLoadingBarEscPollPrev := true
            StandardLoadingBar_KeysEscapeDismiss()
        }
    } else
        g_StandardLoadingBarEscPollPrev := false
}

; $*Escape, I10 g_OnEscapePressed, Gui Escape, and poll all route here (align with #!+C Handy AI model selector).
StandardLoadingBar_KeysEscapeDismiss(*) {
    global g_StandardLoadingBarIsKeysOverlay, g_StandardLoadingBarKeysEscapeUserCb, g_OnEscapePressed
    if !g_StandardLoadingBarIsKeysOverlay {
        g_OnEscapePressed := ""
        return false
    }
    cb := g_StandardLoadingBarKeysEscapeUserCb
    if cb {
        try cb.Call()
        catch {
        }
    }
    StandardLoadingBar_CloseKeysOverlay()
    return true
}

; Show overlay and register hotkeys; optional timeout. keyCallbacks: Map/object key -> callback (e.g. "N" -> fn, "R" -> fn).
; timeoutCallback: called when timeout fires (can be empty). Registers both upper and lower case for letter keys.
; passiveBgColor: optional; when set, used as border color. Prefer BANNER_ACCENT_SUCCESS / BANNER_ACCENT_ERROR / BANNER_ACCENT_INTERMEDIATE. Overlay background stays dark.
; noBorder: when true, do not create the yellow border (single banner only).
; promptKeys: optional; fixed bottom strip text (e.g. "[Y] Confirm  [N] Cancel"). Shown in uniform position below main message.
; trackActiveMonitor: when true, reposition the bar to follow the foreground window's monitor while visible (dictation/Gemini flows).
; showProgress: when true, show a single timed 0-100 progress fill while waiting for keys.
; preserveUserFocus: when true, keep the current active window focused (do not activate overlay GUI).
; overlayBgColor: optional main banner panel color (default dark 1E1E2E); use for themed banners e.g. blackout countdown.
; skipEscapeDismiss: when true, do not register $*Escape / poll (fragile UIs e.g. Command Palette bookmark prompt).
StandardLoadingBar_ShowWithKeys(state, keyCallbacks, timeoutMs := 0, centerOnHwnd := 0, timeoutCallback := "", barColor :=
    BANNER_ACCENT_INTERMEDIATE, textWidth := 500, fontSize := 17, passiveBgColor := "", noBorder := false, promptKeys :=
    "", trackActiveMonitor := false, showProgress := false, preserveUserFocus := false, overlayBgColor := "",
    skipEscapeDismiss := false) {
    global g_StandardLoadingBarIsKeysOverlay, g_StandardLoadingBarKeysHotkeys, g_StandardLoadingBarKeysTimeoutTimer
    global g_StandardLoadingBarGui, g_StandardLoadingBarKeysEscapeUserCb, g_StandardLoadingBarKeysEscapeActive,
        g_StandardLoadingBarEscPollPrev, g_OnEscapePressed
    opts := { passive: !showProgress, centerOnHwnd: centerOnHwnd, textWidth: textWidth, fontSize: fontSize }
    if (showProgress)
        opts.manualProgress := true
    if (overlayBgColor != "")
        opts.overlayBgColor := overlayBgColor
    if (passiveBgColor != "")
        opts.passiveBgColor := passiveBgColor
    if (noBorder)
        opts.noBorder := true
    if (promptKeys != "")
        opts.promptKeys := promptKeys
    if (trackActiveMonitor)
        opts.trackActiveMonitor := true
    StandardLoadingBar_Show(state, barColor, opts)
    if (showProgress)
        StandardLoadingBar_StartTimedProgress(timeoutMs)
    g_StandardLoadingBarIsKeysOverlay := true
    g_StandardLoadingBarKeysHotkeys := []
    escCb := StandardLoadingBar_EscapeCallbackFromKeyCallbacks(keyCallbacks)
    g_StandardLoadingBarKeysEscapeActive := false

    StandardLoadingBar_WaitForTriggerKeyRelease()
    StandardLoadingBar_WaitForSelectionKeysRelease(keyCallbacks)

    ; Register selection hotkeys as GLOBAL while the overlay is open.
    ; Critical: the overlay may fail to activate immediately, and we still need the keys (e.g. "N") to be captured
    ; instead of falling through to the underlying app. Also, clear any caller #HotIf context before registering.
    try HotIf()
    catch {
    }

    ; Register primary and case-variant keys.
    for keyName, cb in keyCallbacks {
        if (!cb)
            continue
        if (!skipEscapeDismiss) {
            knL := StrLower(Trim(keyName))
            if (knL = "*escape" || knL = "escape")
                continue
        }
        StandardLoadingBar_RegisterKeyHandler(keyName, cb)
        if (StrLen(keyName) = 1) {
            o := Ord(keyName)
            alt := ""
            if (o >= Ord("a") && o <= Ord("z"))
                alt := StrUpper(keyName)
            else if (o >= Ord("A") && o <= Ord("Z"))
                alt := StrLower(keyName)
            if (alt != "" && alt != keyName)
                StandardLoadingBar_RegisterKeyHandler(alt, cb)
            if (StandardLoadingBar_IsDigitKey(keyName))
                StandardLoadingBar_RegisterKeyHandler("Numpad" . keyName, cb)
        }
    }

    if (!skipEscapeDismiss) {
        g_StandardLoadingBarKeysEscapeUserCb := escCb ? escCb : ""
        try {
            Hotkey("$*Escape", StandardLoadingBar_KeysEscapeDismiss, "On")
            g_StandardLoadingBarKeysHotkeys.Push("$*Escape")
        } catch {
        }
        g_OnEscapePressed := StandardLoadingBar_KeysEscapeDismiss
        Utils_EnsureGlobalEscapeHotkey()
        try {
            if IsObject(g_StandardLoadingBarGui)
                g_StandardLoadingBarGui.OnEvent("Escape", StandardLoadingBar_KeysEscapeDismiss)
        } catch {
        }
        g_StandardLoadingBarEscPollPrev := GetKeyState("Escape", "P") || (DllCall("user32\GetAsyncKeyState", "int",
            0x1B) &
        0x8000)
        SetTimer(StandardLoadingBar_KeysEscapePoll, 50)
        g_StandardLoadingBarKeysEscapeActive := true
    }

    ; Default behavior keeps key capture reliable by activating the overlay.
    ; Some flows (e.g. dictation E/V paste target) must preserve the user's current text-field focus.
    if (!preserveUserFocus) {
        try {
            if IsObject(g_StandardLoadingBarGui) && g_StandardLoadingBarGui.Hwnd
                WinActivate(g_StandardLoadingBarGui.Hwnd)
        } catch {
        }
    }

    ; Reset any HotIf context so we don't leak it to unrelated hotkeys.
    try HotIf()
    catch {
    }

    if (timeoutMs > 0) {
        g_StandardLoadingBarKeysTimeoutTimer := SetTimer(StandardLoadingBar_KeysTimeoutFired.Bind(timeoutCallback), -
        timeoutMs)
    }

    StandardLoadingBar_StartKeysSelectionPoll(keyCallbacks)
}

StandardLoadingBar_RegisterKeyHandler(key, cb) {
    global g_StandardLoadingBarKeysHotkeys
    if (!cb)
        return
    ; $* prefix: hook hotkey, ignore sent keys; * allows extra modifiers (Shift) still held.
    keyToReg := key
    if (StrLen(key) = 1 || RegExMatch(key, "i)^Numpad[0-9]$"))
        keyToReg := "$*" . key
    fn := StandardLoadingBar_KeyWrapper.Bind(key, cb)
    try {
        #InputLevel 10
        Hotkey(keyToReg, fn, "On")
        #InputLevel 0
        g_StandardLoadingBarKeysHotkeys.Push(keyToReg)
    } catch as err {
    }
}

StandardLoadingBar_KeyWrapper(key, cb, *) {
    global g_StandardLoadingBarIsKeysOverlay
    if (!g_StandardLoadingBarIsKeysOverlay)
        return
    ; Run callback first so it can close the overlay (avoids destroying GUI from hotkey context before callback runs).
    if (cb) {
        try {
            cb.Call()
        }
        catch {
        }
    }
    ; Callback may have closed the keys overlay and started a loading bar — do not destroy the replacement GUI.
    if (g_StandardLoadingBarIsKeysOverlay)
        StandardLoadingBar_CloseKeysOverlay()
}

StandardLoadingBar_KeysTimeoutFired(timeoutCallback) {
    global g_StandardLoadingBarIsKeysOverlay
    ; Nested Func refs from hotkey closures can fail a plain `if (timeoutCallback)` truth test in v2; use HasMethod.
    cbCallable := false
    try {
        if (IsObject(timeoutCallback))
            cbCallable := HasMethod(timeoutCallback, "Call")
    } catch {
        cbCallable := false
    }
    ; Only run timeout callback if overlay was not already dismissed (e.g. user pressed N); avoids copy when timer fires after cancel.
    if (g_StandardLoadingBarIsKeysOverlay && cbCallable) {
        StandardLoadingBar_SetProgressValue(100)
        try timeoutCallback.Call()
        catch {
        }
    }
    StandardLoadingBar_CloseKeysOverlay()
}

; =============================================================================
; All-monitors busy Loading Indication (nestable Begin/End)
; Visual only — does not block input. Separate from the single-monitor bar.
; =============================================================================
global g_BusyAllMonitorsDepth := 0
global g_BusyAllMonitorsOverlays := []   ; [{ overlay: Gui, border: Gui|0 }, ...]
global g_BusyAllMonitorsValue := 0
global g_BusyAllMonitorsTickTimer := ""
global g_BusyAllMonitorsForceTimerArmed := false
global STANDARD_BUSY_ALL_MONITORS_FORCE_MS := 5000

StandardLoadingBar_BusyAllMonitors_Begin(state := "⏳ Arranging window...") {
    global g_BusyAllMonitorsDepth, g_BusyAllMonitorsOverlays, g_BusyAllMonitorsValue
    global g_BusyAllMonitorsTickTimer, g_BusyAllMonitorsForceTimerArmed, g_StandardLoadingBarIsKeysOverlay

    g_BusyAllMonitorsDepth += 1
    if (g_BusyAllMonitorsDepth > 1)
        return

    ; Do not fight an interactive keys overlay (e.g. undo modal).
    if (g_StandardLoadingBarIsKeysOverlay)
        return

    StandardLoadingBar_BusyAllMonitors_DestroyOverlays()
    g_BusyAllMonitorsOverlays := []
    g_BusyAllMonitorsValue := 0

    barColor := BANNER_ACCENT_INTERMEDIATE
    fontSize := 17
    alpha := 235
    borderWidth := 6
    n := 0
    try n := MonitorGetCount()
    catch
        n := 0
    loop n {
        monIdx := A_Index
        try {
            MonitorGetWorkArea(monIdx, &ml, &mt, &mr, &mb)
        } catch {
            continue
        }
        monitorWidth := mr - ml
        barWidth := Min(900, Max(360, Floor(monitorWidth * 0.6)))
        overlayGui := Gui("+AlwaysOnTop -Caption +ToolWindow -DPIScale")
        overlayGui.BackColor := "1E1E2E"
        overlayGui.MarginX := 16
        overlayGui.MarginY := 10
        overlayGui.SetFont("s" . fontSize . " cFFFFFF", "Segoe UI")
        overlayGui.Add("Text", "w" . barWidth . " Center", state)
        progressOpts := "w" . barWidth . " h10 c" . barColor . " Background45475A Smooth vBusyProg"
        overlayGui.Add("Progress", progressOpts, 0)
        overlayGui.Show("AutoSize Hide")
        overlayGui.GetPos(, , &gw, &gh)
        guiX := Round(ml + (monitorWidth - gw) / 2)
        if (guiX < ml)
            guiX := ml
        if (guiX + gw > mr)
            guiX := mr - gw
        guiY := mt + 40

        borderGui := Gui("+AlwaysOnTop -Caption +ToolWindow -DPIScale")
        borderGui.BackColor := barColor
        borderGui.Show("NA x" . (guiX - borderWidth) . " y" . (guiY - borderWidth) . " w" . (gw + 2 * borderWidth) .
        " h" . (gh + 2 * borderWidth))
        overlayGui.Show("x" . guiX . " y" . guiY . " NA")
        try WinSetTransparent(alpha, overlayGui)
        catch {
        }
        g_BusyAllMonitorsOverlays.Push({ overlay: overlayGui, border: borderGui })
    }

    if (g_BusyAllMonitorsOverlays.Length > 0) {
        g_BusyAllMonitorsTickTimer := SetTimer(StandardLoadingBar_BusyAllMonitors_Tick, 40)
        if (!g_BusyAllMonitorsForceTimerArmed) {
            g_BusyAllMonitorsForceTimerArmed := true
            SetTimer(StandardLoadingBar_BusyAllMonitors_ForceClear, -STANDARD_BUSY_ALL_MONITORS_FORCE_MS)
        }
    }
}

StandardLoadingBar_BusyAllMonitors_End() {
    global g_BusyAllMonitorsDepth
    if (g_BusyAllMonitorsDepth <= 0) {
        g_BusyAllMonitorsDepth := 0
        return
    }
    g_BusyAllMonitorsDepth -= 1
    if (g_BusyAllMonitorsDepth > 0)
        return
    StandardLoadingBar_BusyAllMonitors_Clear()
}

StandardLoadingBar_BusyAllMonitors_Clear() {
    global g_BusyAllMonitorsDepth, g_BusyAllMonitorsTickTimer, g_BusyAllMonitorsForceTimerArmed
    g_BusyAllMonitorsDepth := 0
    try SetTimer(StandardLoadingBar_BusyAllMonitors_Tick, 0)
    catch {
    }
    g_BusyAllMonitorsTickTimer := ""
    try SetTimer(StandardLoadingBar_BusyAllMonitors_ForceClear, 0)
    catch {
    }
    g_BusyAllMonitorsForceTimerArmed := false
    StandardLoadingBar_BusyAllMonitors_DestroyOverlays()
}

StandardLoadingBar_BusyAllMonitors_ForceClear(*) {
    global g_BusyAllMonitorsForceTimerArmed
    g_BusyAllMonitorsForceTimerArmed := false
    StandardLoadingBar_BusyAllMonitors_Clear()
}

StandardLoadingBar_BusyAllMonitors_DestroyOverlays() {
    global g_BusyAllMonitorsOverlays
    if (!IsObject(g_BusyAllMonitorsOverlays))
        g_BusyAllMonitorsOverlays := []
    for item in g_BusyAllMonitorsOverlays {
        try {
            if (IsObject(item) && IsObject(item.overlay))
                item.overlay.Destroy()
        } catch {
        }
        try {
            if (IsObject(item) && IsObject(item.border))
                item.border.Destroy()
        } catch {
        }
    }
    g_BusyAllMonitorsOverlays := []
}

StandardLoadingBar_BusyAllMonitors_Tick() {
    global g_BusyAllMonitorsOverlays, g_BusyAllMonitorsValue, g_BusyAllMonitorsDepth
    if (g_BusyAllMonitorsDepth <= 0 || !IsObject(g_BusyAllMonitorsOverlays) || g_BusyAllMonitorsOverlays.Length = 0) {
        try SetTimer(StandardLoadingBar_BusyAllMonitors_Tick, 0)
        catch {
        }
        return
    }
    try {
        g_BusyAllMonitorsValue += 4
        if (g_BusyAllMonitorsValue > 100)
            g_BusyAllMonitorsValue := 0
        for item in g_BusyAllMonitorsOverlays {
            try item.overlay["BusyProg"].Value := g_BusyAllMonitorsValue
            catch {
            }
        }
    } catch {
        try SetTimer(StandardLoadingBar_BusyAllMonitors_Tick, 0)
        catch {
        }
    }
}
