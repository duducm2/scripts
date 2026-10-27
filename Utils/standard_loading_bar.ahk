; =============================================================================
; Utils module: standard_loading_bar.ahk
; Standard loading bar show/update/hide lifecycle
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
;
; Thin-include consumer: Spotify.ahk includes this file alone (not full Utils.ahk)
; and provides a no-op Utils_EnsureGlobalEscapeHotkey() stub. Keep any NEW external
; symbol references shim-able the same way so standalone scripts do not need to
; pull in print_screen_escape / hotstrings / utility_shortcuts.
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
; Optional. Gemini sets this so Win+Alt+Shift+P during the copy menu upgrades to code copy.
global g_StandardLoadingBarChordKeyCb := ""
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
    guiOpt := "+AlwaysOnTop -Caption +ToolWindow -DPIScale"
    if (options && options.HasProp("noActivate") && options.noActivate)
        guiOpt .= " +E0x08000000"
    overlayGui := Gui(guiOpt)
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
    if (!IsObject(g_StandardLoadingBarGui) && !IsObject(g_StandardLoadingBarBorderGui))
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
    ; Hide the accent frame before the text window. The frame is a solid rectangle
    ; behind the text; if the text window closes first, Esc leaves a yellow block.
    if (borderHwnd)
        StandardLoadingBar_ForceDestroyHwnd(borderHwnd)
    if (guiHwnd)
        StandardLoadingBar_ForceDestroyHwnd(guiHwnd)
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
    stillLive := (guiHwnd && WinExist("ahk_id " guiHwnd)) || (borderHwnd && WinExist("ahk_id " borderHwnd))
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
    if (IsObject(g_StandardLoadingBarKeysTimeoutTimer)) {
        try SetTimer(g_StandardLoadingBarKeysTimeoutTimer, 0)
        catch {
        }
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
    ; Accent frame first (see Hide). Gui.Destroy on the RichEdit pane can stall
    ; after that pane is already gone, which leaves the frame as a solid block.
    if (borderHwnd)
        StandardLoadingBar_ForceDestroyHwnd(borderHwnd)
    if (guiHwnd)
        StandardLoadingBar_ForceDestroyHwnd(guiHwnd)
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
    stillLive := (guiHwnd && WinExist("ahk_id " guiHwnd)) || (borderHwnd && WinExist("ahk_id " borderHwnd))
    if (stillLive) {
        g_StandardLoadingBarGui := guiObj
        g_StandardLoadingBarBorderGui := borderObj
        g_StandardLoadingBarIsKeysOverlay := false
        StandardLoadingBar_ArmForceHide(400)
        return
    }
    g_StandardLoadingBarGui := 0
    g_StandardLoadingBarValue := 0
    g_StandardLoadingBarBorderGui := 0
}

; Hide, then destroy. Returns true when the window is gone.
StandardLoadingBar_ForceDestroyHwnd(hwnd) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return true
    try WinHide("ahk_id " hwnd)
    catch {
    }
    try DllCall("user32\DestroyWindow", "ptr", hwnd)
    catch {
    }
    return !WinExist("ahk_id " hwnd)
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
        ; Chord modifiers only (not Shift). Shift stays out so a held Shift after #!+… does not block picks.
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

; Opener letter still physically down after a chord menu opens (e.g. W from #!+W).
; Swallowed until it is released so key-repeat does not activate that entry.
; Lives here so Act and Spotify, which load this file and not utility_shortcuts.ahk, can dismiss too.
global g_UtilitySelectorSwallowKey := ""

; Drop the opener chord without waiting. A KeyWait per held key blocked the menu while the chord was down.
; {vkE8} while Win is down keeps the later Win release from opening Start.
; Modifier ups are inlined so this does not depend on CopilotWeb / utility_shortcuts.
; Defined here: Act and Spotify load this file and not utility_shortcuts.ahk.
StandardLoadingBar_ReleaseStuckModifiers() {
    try SendInput "{Blind}{vk11 up}{vkA2 up}{vkA3 up}{vk12 up}{vkA4 up}{vkA5 up}{vk10 up}{vkA0 up}{vkA1 up}{vk5B up}{vk5C up}"
    catch {
    }
    for vk in [0x11, 0xA2, 0xA3, 0x12, 0xA4, 0xA5, 0x10, 0xA0, 0xA1, 0x5B, 0x5C]
        DllCall("keybd_event", "UChar", vk, "UChar", 0, "UInt", 2, "UPtr", 0)
}

UtilityShortcuts_DismissOpenerChord() {
    try SendInput "{Blind}{vkE8}"
    catch {
    }
    StandardLoadingBar_ReleaseStuckModifiers()
    UtilityShortcuts_ArmOpenerKeySwallow()
}

UtilityShortcuts_ArmOpenerKeySwallow() {
    global g_UtilitySelectorSwallowKey
    g_UtilitySelectorSwallowKey := ""
    SetTimer(UtilityShortcuts_PollOpenerKeyUp, 0)
    th := ""
    try th := A_ThisHotkey
    catch {
        th := ""
    }
    if (th = "")
        return
    hk := RegExReplace(th, "^[$*~]+")
    hk := RegExReplace(hk, "[#^!+<>*]+", "")
    if (StrLen(hk) != 1)
        return
    hk := StrLower(hk)
    if !GetKeyState(hk, "P")
        return
    g_UtilitySelectorSwallowKey := hk
    SetTimer(UtilityShortcuts_PollOpenerKeyUp, 15)
}

UtilityShortcuts_PollOpenerKeyUp(*) {
    global g_UtilitySelectorSwallowKey
    if (g_UtilitySelectorSwallowKey = "" || !GetKeyState(g_UtilitySelectorSwallowKey, "P")) {
        g_UtilitySelectorSwallowKey := ""
        SetTimer(UtilityShortcuts_PollOpenerKeyUp, 0)
    }
}

UtilityShortcuts_ClearOpenerKeySwallow() {
    global g_UtilitySelectorSwallowKey
    g_UtilitySelectorSwallowKey := ""
    SetTimer(UtilityShortcuts_PollOpenerKeyUp, 0)
}

; True while the key that opened the menu is still held (including auto-repeat).
UtilityShortcuts_ShouldSwallowOpenerKey(char) {
    global g_UtilitySelectorSwallowKey
    if (g_UtilitySelectorSwallowKey = "" || StrLower(char) != g_UtilitySelectorSwallowKey)
        return false
    if GetKeyState(g_UtilitySelectorSwallowKey, "P")
        return true
    UtilityShortcuts_ClearOpenerKeySwallow()
    return false
}

; Chord hotkeys only. Plain-key and timer prompts must not release modifiers.
StandardLoadingBar_DismissChordHotkey() {
    th := ""
    try th := A_ThisHotkey
    catch {
        th := ""
    }
    if (th = "" || !RegExMatch(th, "[#^!+]"))
        return
    UtilityShortcuts_DismissOpenerChord()
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
; Chord hotkeys drop modifiers immediately (no KeyWait). Show GUI first (Show calls CloseKeysOverlay — must not run after arming).
StandardLoadingBar_ShowWithKeys(state, keyCallbacks, timeoutMs := 0, centerOnHwnd := 0, timeoutCallback := "", barColor :=
    BANNER_ACCENT_INTERMEDIATE, textWidth := 500, fontSize := 17, passiveBgColor := "", noBorder := false, promptKeys :=
    "", trackActiveMonitor := false, showProgress := false, preserveUserFocus := false, overlayBgColor := "",
    skipEscapeDismiss := false) {
    global g_StandardLoadingBarIsKeysOverlay, g_StandardLoadingBarKeysHotkeys, g_StandardLoadingBarKeysTimeoutTimer
    global g_StandardLoadingBarGui, g_StandardLoadingBarKeysEscapeUserCb, g_StandardLoadingBarKeysEscapeActive,
        g_StandardLoadingBarEscPollPrev, g_OnEscapePressed

    ; Drop Win/Ctrl/Alt/Shift before the banner takes focus. Plain-key and timer prompts skip this.
    StandardLoadingBar_DismissChordHotkey()

    ; Snapshot physical down-state at entry (keys already held must not spuriously fire).
    entryDown := Map()
    try {
        for keyName, cb in keyCallbacks {
            if (!cb)
                continue
            knL := StrLower(Trim(keyName))
            if (knL = "escape" || knL = "*escape")
                continue
            entryDown[keyName] := StandardLoadingBar_KeysSelectionKeyDown(keyName)
        }
    } catch {
    }

    ; Show first: CloseKeysOverlay inside Show clears any prior overlay; no keys registered yet.
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
    if (preserveUserFocus)
        opts.noActivate := true
    StandardLoadingBar_Show(state, barColor, opts)
    if (showProgress)
        StandardLoadingBar_StartTimedProgress(timeoutMs)

    g_StandardLoadingBarIsKeysOverlay := true
    g_StandardLoadingBarKeysHotkeys := []
    escCb := StandardLoadingBar_EscapeCallbackFromKeyCallbacks(keyCallbacks)
    g_StandardLoadingBarKeysEscapeActive := false

    ; Register selection hotkeys as GLOBAL before wait-for-release so rapid keys are not lost.
    ; Critical: the overlay may fail to activate immediately, and we still need the keys (e.g. "N") to be captured
    ; instead of falling through to the underlying app. Also, clear any caller #HotIf context before registering.
    try HotIf()
    catch {
    }

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

    StandardLoadingBar_StartKeysSelectionPoll(keyCallbacks)

    ; Consume mid-setup presses: down now but was up at entry — user pressed during GUI create / arming.
    if (!StandardLoadingBar_KeysSelectionModifiersDown()) {
        try {
            for keyName, cb in keyCallbacks {
                if (!cb || !g_StandardLoadingBarIsKeysOverlay)
                    continue
                knL := StrLower(Trim(keyName))
                if (knL = "escape" || knL = "*escape")
                    continue
                isDown := StandardLoadingBar_KeysSelectionKeyDown(keyName)
                wasAtEntry := entryDown.Has(keyName) ? entryDown[keyName] : false
                if (isDown && !wasAtEntry) {
                    StandardLoadingBar_KeyWrapper(keyName, cb)
                    if (!g_StandardLoadingBarIsKeysOverlay)
                        return
                }
            }
        } catch {
        }
    }

    ; Default behavior keeps key capture reliable by activating the overlay.
    ; preserveUserFocus (Send dictation) must not pull the user back to the window where the take started.
    ; V pastes into whatever field is focused when it is pressed.
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
        ; Store the BoundFunc itself — SetTimer's return value is empty in v2, so
        ; CloseKeysOverlay must cancel this same object (logs: stale 6s timeout killed the next preview).
        if (IsObject(g_StandardLoadingBarKeysTimeoutTimer)) {
            try SetTimer(g_StandardLoadingBarKeysTimeoutTimer, 0)
            catch {
            }
        }
        g_StandardLoadingBarKeysTimeoutTimer := StandardLoadingBar_KeysTimeoutFired.Bind(timeoutCallback)
        SetTimer(g_StandardLoadingBarKeysTimeoutTimer, -timeoutMs)
    }
}

StandardLoadingBar_RegisterKeyHandler(key, cb) {
    global g_StandardLoadingBarKeysHotkeys
    if (!cb)
        return
    ; $* prefix: hook hotkey, ignore sent keys; * allows extra modifiers (Shift) still held.
    ; Enter / NumpadEnter need the hook too, or a focused RichEdit swallows them.
    keyToReg := key
    if (StrLen(key) = 1 || RegExMatch(key, "i)^(Numpad[0-9]|Enter|NumpadEnter)$"))
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
    global g_StandardLoadingBarIsKeysOverlay, g_StandardLoadingBarChordKeyCb
    if (!g_StandardLoadingBarIsKeysOverlay)
        return
    ; A chord such as Win+Alt+Shift+P is the double-tap, not a menu choice.
    ; Do not release those modifiers and then accept the letter (that starts whole-message copy).
    ; The banner's $*P is a higher input level than #!+p, so the second tap lands here.
    ; Hand it back so that tap can upgrade to code copy instead of being dropped.
    chordDown := false
    try chordDown := GetKeyState("LWin", "P") || GetKeyState("RWin", "P") || GetKeyState("Alt", "P") || GetKeyState(
        "Shift", "P")
    ; Only P is the double-tap upgrade. D/Y/F/… must still run while that chord is held,
    ; or Win+Alt+Shift+P then D does nothing (the hook already ate the letter).
    if (chordDown && StrLower(key) = "p") {
        chordCb := g_StandardLoadingBarChordKeyCb
        if (chordCb != "") {
            try chordCb.Call(key)
            return
        }
    }
    ; Ignore while Ctrl is still physically down. A stuck modifier (not still held) is released
    ; so V is not delivered as Clip Angel's Ctrl+Alt+V.
    ; Win/Alt/Shift still held from the opener do not drop other letters.
    if (StandardLoadingBar_KeysSelectionModifiersDown()) {
        StandardLoadingBar_ReleaseStuckModifiers()
        if (StandardLoadingBar_KeysSelectionModifiersDown() && (GetKeyState("Ctrl", "P") || StrLower(key) = "p"))
            return
    }
    if UtilityShortcuts_ShouldSwallowOpenerKey(key)
        return
    if RegExMatch(key, "i)^Numpad([0-9])$", &numpadDigit) && UtilityShortcuts_ShouldSwallowOpenerKey(numpadDigit[1])
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
; Intended rearrange indicator (BusyAllMonitors). Keep false — do not disable for
; “speed”; Place latency is settled via eligibility retries, not by hiding banners.
global STANDARD_BUSY_ALL_MONITORS_DISABLED := false

StandardLoadingBar_BusyAllMonitors_Begin(state := "⏳ Arranging window...") {
    global g_BusyAllMonitorsDepth, g_BusyAllMonitorsOverlays, g_BusyAllMonitorsValue
    global g_BusyAllMonitorsTickTimer, g_BusyAllMonitorsForceTimerArmed, g_StandardLoadingBarIsKeysOverlay
    global STANDARD_BUSY_ALL_MONITORS_DISABLED

    if (STANDARD_BUSY_ALL_MONITORS_DISABLED)
        return

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
    global g_BusyAllMonitorsDepth, STANDARD_BUSY_ALL_MONITORS_DISABLED
    if (STANDARD_BUSY_ALL_MONITORS_DISABLED)
        return
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

; =============================================================================
; All-monitors passive Information Only banner (auto-dismiss; no progress bar)
; Separate from BusyAllMonitors and the single-monitor StandardLoadingBar.
; =============================================================================
global g_PassiveAllMonitorsOverlays := []   ; [{ overlay: Gui, border: Gui|0 }, ...]
global g_PassiveAllMonitorsHideArmed := false

StandardLoadingBar_ShowPassiveAllMonitors(text, durationMs := 3000, accentColor := "") {
    global g_PassiveAllMonitorsOverlays, g_PassiveAllMonitorsHideArmed
    if (text = "")
        return false
    if (accentColor = "")
        accentColor := BANNER_ACCENT_INFO
    if (durationMs < 1)
        durationMs := 1

    ; Clear a stuck keys overlay so push results are never silently skipped.
    try StandardLoadingBar_CloseKeysOverlay()
    catch {
    }

    StandardLoadingBar_PassiveAllMonitors_Clear()
    g_PassiveAllMonitorsOverlays := []

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
        overlayGui.Add("Text", "w" . barWidth . " Wrap Center", text)
        overlayGui.Show("AutoSize Hide")
        overlayGui.GetPos(, , &gw, &gh)
        guiX := Round(ml + (monitorWidth - gw) / 2)
        if (guiX < ml)
            guiX := ml
        if (guiX + gw > mr)
            guiX := mr - gw
        guiY := mt + 40

        borderGui := Gui("+AlwaysOnTop -Caption +ToolWindow -DPIScale")
        borderGui.BackColor := accentColor
        borderGui.Show("NA x" . (guiX - borderWidth) . " y" . (guiY - borderWidth) . " w" . (gw + 2 * borderWidth) .
        " h" . (gh + 2 * borderWidth))
        overlayGui.Show("x" . guiX . " y" . guiY . " NA")
        ; Raise overlay above other AlwaysOnTop windows (border stays under it).
        try {
            hwnd := overlayGui.Hwnd
            if (hwnd)
                DllCall("SetWindowPos", "Ptr", hwnd, "Ptr", -1, "Int", 0, "Int", 0, "Int", 0, "Int", 0,
                    "UInt", 0x0013)
        } catch {
        }
        try WinSetTransparent(alpha, overlayGui)
        catch {
        }
        g_PassiveAllMonitorsOverlays.Push({ overlay: overlayGui, border: borderGui })
    }

    if (g_PassiveAllMonitorsOverlays.Length = 0)
        return false

    try SetTimer(StandardLoadingBar_PassiveAllMonitors_HideNow, 0)
    catch {
    }
    SetTimer(StandardLoadingBar_PassiveAllMonitors_HideNow, -durationMs)
    g_PassiveAllMonitorsHideArmed := true
    ; Hard ceiling so a missed hide cannot leave banners stuck.
    forceMs := durationMs + 2000
    if (forceMs < 5000)
        forceMs := 5000
    SetTimer(StandardLoadingBar_PassiveAllMonitors_ForceClear, -forceMs)
    return true
}

StandardLoadingBar_PassiveAllMonitors_HideNow(*) {
    global g_PassiveAllMonitorsHideArmed
    g_PassiveAllMonitorsHideArmed := false
    try SetTimer(StandardLoadingBar_PassiveAllMonitors_ForceClear, 0)
    catch {
    }
    StandardLoadingBar_PassiveAllMonitors_Clear()
}

StandardLoadingBar_PassiveAllMonitors_ForceClear(*) {
    StandardLoadingBar_PassiveAllMonitors_Clear()
}

StandardLoadingBar_PassiveAllMonitors_Clear() {
    global g_PassiveAllMonitorsOverlays, g_PassiveAllMonitorsHideArmed
    g_PassiveAllMonitorsHideArmed := false
    try SetTimer(StandardLoadingBar_PassiveAllMonitors_HideNow, 0)
    catch {
    }
    try SetTimer(StandardLoadingBar_PassiveAllMonitors_ForceClear, 0)
    catch {
    }
    if (!IsObject(g_PassiveAllMonitorsOverlays))
        g_PassiveAllMonitorsOverlays := []
    for item in g_PassiveAllMonitorsOverlays {
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
    g_PassiveAllMonitorsOverlays := []
}

; Full AI reply on the monitor under the foreground window.
; The pane is landscape on a wide screen and never taller than that monitor's work area.
; Markdown (**bold**, headings, lists, code) is drawn as formatting, not raw asterisks.
; No auto-dismiss — user closes with Enter, E, or Escape.
ShowCopiedResponseBanner(text, anchorHwnd := 0) {
    global g_ShowCopiedResponseRichDll
    if (!text || StrLen(Trim(text)) = 0) {
        ShowCopiedResponse_LegacyBanner("Nothing to display")
        return
    }
    ml := mt := mr := mb := 0
    placed := false
    if (anchorHwnd) {
        work := GetWorkAreaForWindow_StandardBar(anchorHwnd)
        if (work != "") {
            ml := work.left
            mt := work.top
            mr := work.right
            mb := work.bottom
            placed := true
        }
    }
    if (!placed)
        GetActiveMonitorWorkArea_StandardBar(&ml, &mt, &mr, &mb)
    ; Drop a prior bar before building this one so its teardown cannot delete the new GUI.
    try {
        if (g_StandardLoadingBarIsKeysOverlay)
            StandardLoadingBar_CloseKeysOverlay()
        else
            StandardLoadingBar_Hide(0)
    } catch {
    }
    try {
        if (!g_ShowCopiedResponseRichDll)
            g_ShowCopiedResponseRichDll := DllCall("LoadLibrary", "str", "msftedit.dll", "ptr")
        if (!g_ShowCopiedResponseRichDll)
            throw Error("msftedit.dll unavailable")
        ShowCopiedResponse_ShowReadingPane(text, ml, mt, mr, mb)
    } catch {
        ShowCopiedResponse_LegacyBanner(text)
    }
}

ShowCopiedResponse_LegacyBanner(text) {
    state := "ℹ " . text
    closeKeys := Map(
        "Enter", ShowCopiedResponse_DismissNoOp,
        "Escape", ShowCopiedResponse_DismissNoOp,
        "E", ShowCopiedResponse_DismissNoOp)
    StandardLoadingBar_ShowWithKeys(state, closeKeys, 0, 0, "",
        BANNER_ACCENT_INTERMEDIATE, 600,
        17, "", false,
        "[Enter] [E] [Esc] Close",
        true)
}

ShowCopiedResponse_DismissNoOp(*) {
}

; Punch out the center so the accent window is a frame, not a solid block.
ShowCopiedResponse_FrameBorder(hwnd, w, h, inset) {
    if (!hwnd || w <= inset * 2 || h <= inset * 2)
        return
    outer := DllCall("gdi32\CreateRectRgn", "int", 0, "int", 0, "int", w, "int", h, "ptr")
    inner := DllCall("gdi32\CreateRectRgn", "int", inset, "int", inset, "int", w - inset, "int", h - inset, "ptr")
    if (!outer || !inner) {
        if (outer)
            DllCall("gdi32\DeleteObject", "ptr", outer)
        if (inner)
            DllCall("gdi32\DeleteObject", "ptr", inner)
        return
    }
    DllCall("gdi32\CombineRgn", "ptr", outer, "ptr", outer, "ptr", inner, "int", 4)
    DllCall("gdi32\DeleteObject", "ptr", inner)
    ; SetWindowRgn takes ownership of the region when it succeeds.
    if !DllCall("user32\SetWindowRgn", "ptr", hwnd, "ptr", outer, "int", 1)
        DllCall("gdi32\DeleteObject", "ptr", outer)
}

ShowCopiedResponse_ShowReadingPane(text, ml, mt, mr, mb) {
    global g_StandardLoadingBarGui, g_StandardLoadingBarBorderGui, g_StandardLoadingBarValue
    global g_ShowCopiedResponseRichDll
    if (!g_ShowCopiedResponseRichDll)
        g_ShowCopiedResponseRichDll := DllCall("LoadLibrary", "str", "msftedit.dll", "ptr")
    blocks := ShowCopiedResponse_ParseBlocks(text)
    dpi := ShowCopiedResponse_DpiAt(ml, mt, mr, mb)
    layout := ShowCopiedResponse_Layout(blocks, ml, mt, mr, mb, dpi)

    border := 6
    guiOpt := "+AlwaysOnTop -Caption +ToolWindow -DPIScale"
    borderGui := Gui(guiOpt)
    borderGui.BackColor := BANNER_ACCENT_INTERMEDIATE
    overlayGui := Gui(guiOpt)
    overlayGui.BackColor := "1E1E2E"
    overlayGui.MarginX := 0
    overlayGui.MarginY := 0

    editOpts := "ClassRICHEDIT50W x" . layout.marginX . " y" . layout.marginY
        . " w" . layout.editW . " h" . layout.editH . " +0x44 +VScroll -E0x200"
    edit := overlayGui.Add("Custom", editOpts)
    footerY := layout.marginY + layout.editH + 8
    overlayGui.SetFont("s13 cA6ADC8", "Segoe UI")
    overlayGui.Add("Text", "x" . layout.marginX . " y" . footerY . " w" . layout.editW . " h" . layout.footerH .
        " Center", "[Enter]    [E]    [Esc]    Close")

    g_StandardLoadingBarBorderGui := borderGui
    g_StandardLoadingBarGui := overlayGui
    g_StandardLoadingBarValue := 0

    hwnd := edit.Hwnd
    ShowCopiedResponse_ThemeEdit(hwnd)
    bg := ShowCopiedResponse_ColorRef("1E1E2E")
    SendMessage(0x435, 0, 0x7FFFFFFE, hwnd)
    SendMessage(0x443, 0, bg, hwnd)

    guiX := layout.guiX
    guiY := layout.guiY
    guiW := layout.guiW
    guiH := layout.guiH
    borderW := guiW + 2 * border
    borderH := guiH + 2 * border
    borderGui.Show("NA x" . (guiX - border) . " y" . (guiY - border) . " w" . borderW . " h" . borderH)
    ; Hollow frame, so a missed destroy cannot leave a solid yellow slab.
    ShowCopiedResponse_FrameBorder(borderGui.Hwnd, borderW, borderH, border)
    overlayGui.Show("x" . guiX . " y" . guiY . " w" . guiW . " h" . guiH . " NA")
    try {
        if (hwnd)
            DllCall("SetWindowPos", "Ptr", overlayGui.Hwnd, "Ptr", 0, "Int", guiX, "Int", guiY, "Int", 0, "Int", 0,
                "UInt", 0x0015)
    } catch {
    }
    ; Arm before the long render so Enter/E/Esc work if drawing a huge reply takes a while.
    ShowCopiedResponse_ArmDismissKeys()
    try StandardLoadingBar_CancelHideTimers()
    catch {
    }
    try {
        ShowCopiedResponse_Render(hwnd, blocks, layout.bodyPt)
    } catch {
        try ShowCopiedResponse_SetPlain(hwnd, text)
        catch {
        }
    }
    SendMessage(0x443, 0, bg, hwnd)
    SendMessage(0xCF, 1, 0, hwnd)
    try edit.Focus()
    catch {
    }
    ; Collapse the formatting selection and start at the top. Do this after focus,
    ; or the first focus paints the Windows highlight over the reply.
    SendMessage(0xB1, 0, 0, hwnd)
    scrollPt := Buffer(8, 0)
    SendMessage(0x4DD, 0, scrollPt.Ptr, hwnd)
}

ShowCopiedResponse_Render(hwnd, blocks, bodyPt) {
    doc := ShowCopiedResponse_BuildDocument(blocks, bodyPt)
    try SendMessage(0x0B, 0, 0, hwnd)
    try {
        ShowCopiedResponse_SetPlain(hwnd, doc.plain)
        for sp in doc.spans {
            if (sp.len <= 0)
                continue
            ShowCopiedResponse_Select(hwnd, sp.start, sp.start + sp.len)
            cf := ShowCopiedResponse_CharFormat(sp.face, sp.pt, sp.color, sp.bold, sp.italic, sp.underline)
            SendMessage(0x444, 1, cf.Ptr, hwnd)
        }
        for para in doc.paras {
            if (para.len <= 0)
                continue
            ShowCopiedResponse_Select(hwnd, para.start, para.start + Max(para.len, 1))
            pf := ShowCopiedResponse_ParaFormat(para.indent, para.offset, para.before, para.after)
            SendMessage(0x447, 0, pf.Ptr, hwnd)
        }
        ShowCopiedResponse_Select(hwnd, 0, 0)
    } finally {
        try SendMessage(0x0B, 1, 0, hwnd)
    }
    try DllCall("RedrawWindow", "ptr", hwnd, "ptr", 0, "ptr", 0, "uint", 0x105)
}

ShowCopiedResponse_BuildDocument(blocks, bodyPt) {
    doc := { plain: "", spans: [], paras: [] }
    u16 := 0
    codePt := Max(12, bodyPt - 1)
    first := true
    body := { bold: false, italic: false, code: false, underline: false, pt: bodyPt, color: "CDD6F4", face: "Segoe UI" }
    for b in blocks {
        paraStart := u16
        kind := b.kind
        indent := 0
        offset := 0
        before := 40
        after := 80
        if (kind = "h") {
            level := b.level
            add := level = 1 ? 8 : level = 2 ? 5 : level = 3 ? 2 : 1
            head := { bold: true, italic: false, code: false, underline: false, pt: bodyPt + add, color: "CBA6F7",
                face: "Segoe UI" }
            ShowCopiedResponse_AppendRun(doc, &u16, ShowCopiedResponse_PlainInline(b.text), head)
            before := first ? 40 : (level <= 2 ? 220 : 140)
            after := 70
        } else if (kind = "hr") {
            muted := { bold: false, italic: false, code: false, underline: false, pt: bodyPt, color: "A6ADC8",
                face: "Segoe UI" }
            ShowCopiedResponse_AppendRun(doc, &u16, "────────────────────────────────", muted)
            before := 50
            after := 50
        } else if (kind = "ul" || kind = "ol") {
            prefix := kind = "ul" ? "•  " : (b.n . ".  ")
            ShowCopiedResponse_AppendRun(doc, &u16, prefix, body)
            ShowCopiedResponse_AppendInline(doc, &u16, b.text, bodyPt)
            before := 10
            after := 36
            indent := 420
            offset := kind = "ul" ? -200 : -260
        } else if (kind = "quote") {
            ShowCopiedResponse_AppendInline(doc, &u16, b.text, bodyPt, true, "89B4FA")
            before := 30
            after := 70
            indent := 280
        } else if (kind = "code") {
            codeStyle := { bold: false, italic: false, code: true, underline: false, pt: codePt, color: "A6E3A1",
                face: "Consolas" }
            ShowCopiedResponse_AppendRun(doc, &u16, b.text, codeStyle)
            before := 50
            after := 90
            indent := 200
        } else if (kind = "table") {
            tableStyle := { bold: false, italic: false, code: true, underline: false, pt: codePt, color: "CDD6F4",
                face: "Consolas" }
            ShowCopiedResponse_AppendRun(doc, &u16, ShowCopiedResponse_FormatTable(b.text), tableStyle)
            before := 30
            after := 70
            indent := 80
        } else {
            ShowCopiedResponse_AppendInline(doc, &u16, b.text, bodyPt)
            before := first ? 16 : 30
            after := 110
        }
        doc.paras.Push({
            start: paraStart,
            len: u16 - paraStart,
            indent: indent,
            offset: offset,
            before: before,
            after: after
        })
        doc.plain .= "`r"
        u16++
        first := false
    }
    return doc
}

ShowCopiedResponse_AppendInline(doc, &u16, text, bodyPt, forceItalic := false, color := "CDD6F4") {
    for r in ShowCopiedResponse_InlineRuns(text) {
        face := r.code ? "Consolas" : "Segoe UI"
        pt := r.code ? Max(12, bodyPt - 1) : bodyPt
        col := r.code ? "A6E3A1" : (r.underline ? "89B4FA" : color)
        ShowCopiedResponse_AppendRun(doc, &u16, r.text, {
            bold: r.bold,
            italic: r.italic || (forceItalic && !r.code),
            code: r.code,
            underline: r.underline,
            pt: pt,
            color: col,
            face: face
        })
    }
}

ShowCopiedResponse_AppendRun(doc, &u16, text, style) {
    if (text = "")
        return
    text := StrReplace(text, "`r`n", "`n")
    text := StrReplace(text, "`n", "`r")
    len := ShowCopiedResponse_Utf16Len(text)
    if (len <= 0)
        return
    doc.spans.Push({
        start: u16,
        len: len,
        bold: style.bold,
        italic: style.italic,
        code: style.code,
        underline: style.underline,
        pt: style.pt,
        color: style.color,
        face: style.face
    })
    doc.plain .= text
    u16 += len
}

ShowCopiedResponse_Utf16Len(s) {
    if (s = "")
        return 0
    ; StrPut without a buffer returns the size in bytes, including the terminating null.
    return (StrPut(s, "UTF-16") // 2) - 1
}

ShowCopiedResponse_SetPlain(hwnd, plain) {
    settextex := Buffer(8, 0)
    NumPut("uint", 8, settextex, 0)
    NumPut("uint", 1200, settextex, 4)
    bytes := StrPut(plain, "UTF-16")
    textBuf := Buffer(bytes, 0)
    StrPut(plain, textBuf, "UTF-16")
    SendMessage(0x461, settextex.Ptr, textBuf.Ptr, hwnd)
}

ShowCopiedResponse_Select(hwnd, cpMin, cpMax) {
    SendMessage(0xB1, cpMin, cpMax, hwnd)
}

ShowCopiedResponse_CharFormat(face, pt, colorHex, bold, italic, underline) {
    cf := Buffer(116, 0)
    NumPut("uint", 116, cf, 0)
    mask := 0x1 | 0x2 | 0x4 | 0x40000000 | 0x80000000 | 0x20000000
    effects := (bold ? 0x1 : 0) | (italic ? 0x2 : 0) | (underline ? 0x4 : 0)
    NumPut("uint", mask, cf, 4)
    NumPut("uint", effects, cf, 8)
    NumPut("int", Round(pt * 20), cf, 12)
    NumPut("uint", ShowCopiedResponse_ColorRef(colorHex), cf, 20)
    NumPut("uchar", 1, cf, 24)
    StrPut(face, cf.Ptr + 26, 32, "UTF-16")
    return cf
}

ShowCopiedResponse_ParaFormat(startIndent, offset, spaceBefore, spaceAfter) {
    pf := Buffer(188, 0)
    NumPut("uint", 188, pf, 0)
    ; Start indent, hanging offset, space before and after. Leave line spacing at the control default
    ; so wrapped lines do not stack (a twip-sized rule collapses a paragraph onto itself).
    NumPut("uint", 0x1 | 0x4 | 0x40 | 0x80, pf, 4)
    NumPut("int", startIndent, pf, 12)
    NumPut("int", offset, pf, 20)
    NumPut("int", spaceBefore, pf, 156)
    NumPut("int", spaceAfter, pf, 160)
    return pf
}

ShowCopiedResponse_ThemeEdit(hwnd) {
    try DllCall("uxtheme\SetWindowTheme", "ptr", hwnd, "wstr", "", "wstr", "")
    catch {
    }
    ; ECOOP_OR + ECO_AUTOVSCROLL so the wheel and Page Down move through a long reply.
    try SendMessage(0x44D, 2, 0x40, hwnd)
    catch {
    }
}

ShowCopiedResponse_DpiAt(ml, mt, mr, mb) {
    cx := ml + (mr - ml) // 2
    cy := mt + (mb - mt) // 2
    rect := Buffer(16, 0)
    NumPut("int", cx, rect, 0)
    NumPut("int", cy, rect, 4)
    NumPut("int", cx + 1, rect, 8)
    NumPut("int", cy + 1, rect, 12)
    hmon := 0
    try hmon := DllCall("user32\MonitorFromRect", "ptr", rect, "uint", 2, "ptr")
    catch
        hmon := 0
    dx := 0
    dy := 0
    if (hmon) {
        try {
            if (DllCall("shcore\GetDpiForMonitor", "ptr", hmon, "int", 0, "uint*", &dx, "uint*", &dy) = 0 && dx >= 96)
                return dx
        } catch {
        }
    }
    dpi := 96
    try dpi := A_ScreenDPI
    catch
        dpi := 96
    return dpi >= 96 ? dpi : 96
}

; Wide monitors get a landscape pane (wider than it is tall). Height never leaves the work area.
ShowCopiedResponse_Layout(blocks, ml, mt, mr, mb, dpi) {
    workW := Max(1, mr - ml)
    workH := Max(1, mb - mt)
    border := 6
    outer := Max(12, Floor(Min(workW, workH) * 0.03))
    marginX := 26
    marginY := 18
    footerH := 28
    gap := 8
    maxGuiW := Max(480, workW - 2 * outer - 2 * border)
    maxGuiH := Max(240, workH - 2 * outer - 2 * border)
    maxEditW := Max(400, maxGuiW - marginX * 2)
    maxEditH := Max(140, maxGuiH - marginY * 2 - footerH - gap)

    scale := dpi / 96
    bodyPt := workH >= 1400 ? 18 : workH >= 1050 ? 17 : 15
    charPx := Max(7, Round(bodyPt * 0.5 * scale))
    linePx := Max(22, Round(bodyPt * 1.85 * scale))
    aspect := workW / workH

    if (aspect >= 1.35) {
        ; Landscape card: at least 1.4× as wide as the tallest pane that still fits, and a solid share of the width.
        editHWant := maxEditH
        editW := Min(maxEditW, Max(Floor(editHWant * 1.4), Floor(workW * 0.62)))
    } else {
        editW := Min(maxEditW, Max(480, Floor(workW * 0.92)))
    }
    ; Grow toward the full width when a narrower column would run past the bottom of this monitor.
    guard := 0
    loop {
        needH := ShowCopiedResponse_EstimateLines(blocks, editW, charPx) * linePx + 28
        if (needH <= maxEditH || editW >= maxEditW || guard >= 8)
            break
        guard++
        editW := Min(maxEditW, editW + Max(120, Floor((maxEditW - editW) / 2)))
    }
    needH := ShowCopiedResponse_EstimateLines(blocks, editW, charPx) * linePx + 28
    ; A short note stays a wide card, but not a near-empty slab.
    if (needH < maxEditH * 0.38) {
        shorterW := Min(editW, Max(680, Floor(workW * (aspect >= 1.35 ? 0.5 : 0.86))))
        shortNeed := ShowCopiedResponse_EstimateLines(blocks, shorterW, charPx) * linePx + 28
        if (shortNeed <= maxEditH) {
            editW := shorterW
            needH := shortNeed
        }
    }
    editH := Min(maxEditH, Max(150, needH))
    editW := Min(editW, maxEditW)

    guiW := editW + marginX * 2
    guiH := marginY + editH + gap + footerH + marginY
    if (guiW + 2 * border > workW) {
        guiW := workW - 2 * border
        editW := Max(200, guiW - marginX * 2)
    }
    if (guiH + 2 * border > workH) {
        guiH := workH - 2 * border
        editH := Max(80, guiH - marginY * 2 - gap - footerH)
    }
    guiX := ml + Floor((workW - guiW) / 2)
    guiY := mt + Floor((workH - guiH) / 2)
    if (guiX - border < ml)
        guiX := ml + border
    if (guiY - border < mt)
        guiY := mt + border
    if (guiX + guiW + border > mr)
        guiX := mr - border - guiW
    if (guiY + guiH + border > mb)
        guiY := mb - border - guiH
    return {
        editW: editW,
        editH: editH,
        guiW: guiW,
        guiH: guiH,
        guiX: guiX,
        guiY: guiY,
        marginX: marginX,
        marginY: marginY,
        footerH: footerH,
        bodyPt: bodyPt
    }
}

ShowCopiedResponse_EstimateLines(blocks, contentW, charPx) {
    cpl := Max(20, Floor(contentW / Max(charPx, 1)))
    lines := 0
    for b in blocks {
        kind := b.kind
        if (kind = "hr") {
            lines += 1
            continue
        }
        text := b.text
        if (kind = "code" || kind = "table") {
            for ln in StrSplit(text, "`n")
                lines += Max(1, Ceil(Max(StrLen(ln), 1) / cpl))
            lines += 1
            continue
        }
        vis := RegExReplace(text, "[*_``#]+", "")
        prefix := kind = "ol" ? 4 : kind = "ul" ? 2 : 0
        n := Max(StrLen(vis) + prefix, 1)
        lines += Max(1, Ceil(n / cpl))
        if (kind = "h")
            lines += 1
    }
    return Max(1, lines)
}

ShowCopiedResponse_ParseBlocks(text) {
    src := StrReplace(text, "`r`n", "`n")
    src := StrReplace(src, "`r", "`n")
    src := Trim(src, " `t`n")
    raw := StrSplit(src, "`n")
    blocks := []
    i := 1
    while (i <= raw.Length) {
        line := raw[i]
        trimmed := Trim(line, " `t")
        if (RegExMatch(trimmed, "^``````")) {
            i++
            code := ""
            while (i <= raw.Length && !RegExMatch(Trim(raw[i], " `t"), "^``````")) {
                code .= (code = "" ? "" : "`n") . raw[i]
                i++
            }
            if (i <= raw.Length)
                i++
            blocks.Push({ kind: "code", text: code, level: 0, n: "" })
            continue
        }
        if (trimmed = "") {
            i++
            continue
        }
        if (RegExMatch(trimmed, "^(#{1,6})\s+(.*)$", &hm)) {
            blocks.Push({ kind: "h", text: hm[2], level: StrLen(hm[1]), n: "" })
            i++
            continue
        }
        if (RegExMatch(trimmed, "^(-{3,}|\*{3,}|_{3,})$")) {
            blocks.Push({ kind: "hr", text: "", level: 0, n: "" })
            i++
            continue
        }
        if (ShowCopiedResponse_IsTableLine(trimmed)) {
            rows := trimmed
            i++
            while (i <= raw.Length && ShowCopiedResponse_IsTableLine(Trim(raw[i], " `t"))) {
                rows .= "`n" . Trim(raw[i], " `t")
                i++
            }
            blocks.Push({ kind: "table", text: rows, level: 0, n: "" })
            continue
        }
        if (RegExMatch(trimmed, "^>\s?(.*)$", &qm)) {
            quote := qm[1]
            i++
            while (i <= raw.Length) {
                nxt := Trim(raw[i], " `t")
                if (nxt = "" || !RegExMatch(nxt, "^>\s?(.*)$", &qm2))
                    break
                quote .= " " . qm2[1]
                i++
            }
            blocks.Push({ kind: "quote", text: quote, level: 0, n: "" })
            continue
        }
        if (RegExMatch(trimmed, "^(\d+)\.\s+(.*)$", &nm)) {
            item := nm[2]
            i++
            item := ShowCopiedResponse_TakeContinuations(raw, &i, item)
            blocks.Push({ kind: "ol", text: item, level: 0, n: nm[1] })
            continue
        }
        if (RegExMatch(trimmed, "^[-*+]\s+(.*)$", &bm)) {
            item := bm[1]
            i++
            item := ShowCopiedResponse_TakeContinuations(raw, &i, item)
            blocks.Push({ kind: "ul", text: item, level: 0, n: "" })
            continue
        }
        para := trimmed
        i++
        para := ShowCopiedResponse_TakeContinuations(raw, &i, para)
        blocks.Push({ kind: "p", text: para, level: 0, n: "" })
    }
    if (blocks.Length = 0)
        blocks.Push({ kind: "p", text: Trim(text), level: 0, n: "" })
    return blocks
}

ShowCopiedResponse_TakeContinuations(raw, &i, text) {
    while (i <= raw.Length) {
        nxt := Trim(raw[i], " `t")
        if (nxt = "" || ShowCopiedResponse_IsBlockStart(nxt))
            break
        text .= " " . nxt
        i++
    }
    return text
}

ShowCopiedResponse_IsBlockStart(trimmed) {
    if (trimmed = "")
        return false
    if (RegExMatch(trimmed, "^(#{1,6})\s+\S"))
        return true
    if (RegExMatch(trimmed, "^``````"))
        return true
    if (RegExMatch(trimmed, "^>\s?"))
        return true
    if (RegExMatch(trimmed, "^(-{3,}|\*{3,}|_{3,})$"))
        return true
    if (RegExMatch(trimmed, "^(\d+)\.\s+\S"))
        return true
    if (RegExMatch(trimmed, "^[-*+]\s+\S"))
        return true
    if (ShowCopiedResponse_IsTableLine(trimmed))
        return true
    return false
}

ShowCopiedResponse_IsTableLine(trimmed) {
    return !!RegExMatch(trimmed, "\|.*\|")
}

ShowCopiedResponse_FormatTable(text) {
    parsed := []
    widths := []
    for row in StrSplit(text, "`n") {
        if (RegExMatch(Trim(row), "^\|?\s*:?-{2,}.*\|"))
            continue
        parts := StrSplit(row, "|")
        cells := []
        for p in parts {
            cell := ShowCopiedResponse_PlainInline(Trim(p))
            if (cell = "" && (A_Index = 1 || A_Index = parts.Length))
                continue
            cells.Push(cell)
        }
        if (cells.Length = 0)
            continue
        parsed.Push(cells)
        for cell in cells {
            idx := A_Index
            w := StrLen(cell)
            if (idx > widths.Length || w > widths[idx]) {
                while (widths.Length < idx)
                    widths.Push(0)
                widths[idx] := w
            }
        }
    }
    out := ""
    for cells in parsed {
        line := ""
        for cell in cells {
            idx := A_Index
            pad := idx < cells.Length ? Max(2, widths[idx] - StrLen(cell) + 3) : 0
            gap := ""
            loop pad
                gap .= " "
            line .= cell . gap
        }
        out .= (out = "" ? "" : "`n") . RTrim(line)
    }
    return out
}

ShowCopiedResponse_PlainInline(s) {
    s := RegExReplace(s, "!\[([^\]]*)\]\([^)]*\)", "$1")
    s := RegExReplace(s, "\[([^\]]+)\]\([^)]*\)", "$1")
    s := RegExReplace(s, "``([^``]*)``", "$1")
    s := RegExReplace(s, "\*\*\*([^*]+)\*\*\*", "$1")
    s := RegExReplace(s, "\*\*([^*]+)\*\*", "$1")
    s := RegExReplace(s, "__([^_]+)__", "$1")
    s := RegExReplace(s, "(?<![\w*])\*([^*\n]+)\*(?!\*)", "$1")
    s := RegExReplace(s, "(?<!\w)_([^_\n]+)_(?!\w)", "$1")
    return s
}

ShowCopiedResponse_InlineRuns(s, bold := false, italic := false) {
    runs := []
    if (s = "")
        return runs
    buf := ""
    i := 1
    n := StrLen(s)
    while (i <= n) {
        ch := SubStr(s, i, 1)
        if (ch = "\" && i < n) {
            nxt := SubStr(s, i + 1, 1)
            if (InStr("*_``[]()#!", nxt)) {
                buf .= nxt
                i += 2
                continue
            }
        }
        if (ch = "``") {
            close := InStr(s, "``", , i + 1)
            if (close > i) {
                ShowCopiedResponse_FlushRun(runs, &buf, bold, italic, false, false)
                inner := SubStr(s, i + 1, close - i - 1)
                if (inner != "")
                    runs.Push(ShowCopiedResponse_Run(inner, false, false, true, false))
                i := close + 1
                continue
            }
        }
        if (SubStr(s, i, 3) = "***") {
            close := InStr(s, "***", , i + 3)
            if (close) {
                ShowCopiedResponse_FlushRun(runs, &buf, bold, italic, false, false)
                inner := SubStr(s, i + 3, close - (i + 3))
                for r in ShowCopiedResponse_InlineRuns(inner, true, true)
                    runs.Push(r)
                i := close + 3
                continue
            }
        }
        marker2 := SubStr(s, i, 2)
        if (marker2 = "**" || marker2 = "__") {
            close := InStr(s, marker2, , i + 2)
            if (close) {
                ShowCopiedResponse_FlushRun(runs, &buf, bold, italic, false, false)
                inner := SubStr(s, i + 2, close - (i + 2))
                for r in ShowCopiedResponse_InlineRuns(inner, true, italic)
                    runs.Push(r)
                i := close + 2
                continue
            }
        }
        if (ch = "*" || (ch = "_" && ShowCopiedResponse_UnderscoreOpens(s, i))) {
            close := ShowCopiedResponse_FindSingleMarker(s, ch, i + 1)
            if (close) {
                inner := SubStr(s, i + 1, close - i - 1)
                if (Trim(inner) != "") {
                    ShowCopiedResponse_FlushRun(runs, &buf, bold, italic, false, false)
                    for r in ShowCopiedResponse_InlineRuns(inner, bold, true)
                        runs.Push(r)
                    i := close + 1
                    continue
                }
            }
        }
        if ((ch = "!" && SubStr(s, i + 1, 1) = "[") || ch = "[") {
            bracketPos := ch = "!" ? i + 1 : i
            link := ShowCopiedResponse_ReadLink(s, bracketPos)
            if (IsObject(link)) {
                ShowCopiedResponse_FlushRun(runs, &buf, bold, italic, false, false)
                for r in ShowCopiedResponse_InlineRuns(link.label, bold, italic) {
                    r.underline := true
                    runs.Push(r)
                }
                i := link.next
                continue
            }
        }
        buf .= ch
        i++
    }
    ShowCopiedResponse_FlushRun(runs, &buf, bold, italic, false, false)
    return runs
}

ShowCopiedResponse_FlushRun(runs, &buf, bold, italic, code, underline) {
    if (buf = "")
        return
    runs.Push(ShowCopiedResponse_Run(buf, bold, italic, code, underline))
    buf := ""
}

ShowCopiedResponse_Run(text, bold, italic, code, underline) {
    return { text: text, bold: bold, italic: italic, code: code, underline: underline }
}

ShowCopiedResponse_ReadLink(s, bracketPos) {
    rb := InStr(s, "]", , bracketPos + 1)
    if (!rb || SubStr(s, rb + 1, 1) != "(")
        return ""
    rp := InStr(s, ")", , rb + 2)
    if (!rp)
        return ""
    label := SubStr(s, bracketPos + 1, rb - bracketPos - 1)
    if (label = "")
        label := SubStr(s, rb + 2, rp - rb - 2)
    return { label: label, next: rp + 1 }
}

ShowCopiedResponse_UnderscoreOpens(s, i) {
    prev := i > 1 ? SubStr(s, i - 1, 1) : ""
    next := SubStr(s, i + 1, 1)
    if (next = "" || next = " " || next = "_")
        return false
    return !ShowCopiedResponse_IsWord(prev)
}

ShowCopiedResponse_FindSingleMarker(s, marker, start) {
    i := start
    n := StrLen(s)
    while (i <= n) {
        p := InStr(s, marker, , i)
        if (!p)
            return 0
        prev := p > 1 ? SubStr(s, p - 1, 1) : ""
        next := SubStr(s, p + 1, 1)
        if (marker = "*" && (prev = "*" || next = "*")) {
            i := p + 1
            continue
        }
        if (marker = "_" && (ShowCopiedResponse_IsWord(next) || prev = "_")) {
            i := p + 1
            continue
        }
        return p
    }
    return 0
}

ShowCopiedResponse_IsWord(ch) {
    return ch != "" && !!RegExMatch(ch, "^[A-Za-z0-9]$")
}

ShowCopiedResponse_ColorRef(hex) {
    rr := Integer("0x" . SubStr(hex, 1, 2))
    gg := Integer("0x" . SubStr(hex, 3, 2))
    bb := Integer("0x" . SubStr(hex, 5, 2))
    return (bb << 16) | (gg << 8) | rr
}

global g_ShowCopiedResponseRichDll := 0

ShowCopiedResponse_ArmDismissKeys() {
    global g_StandardLoadingBarIsKeysOverlay, g_StandardLoadingBarKeysHotkeys
    global g_StandardLoadingBarGui, g_StandardLoadingBarKeysEscapeUserCb, g_StandardLoadingBarKeysEscapeActive
    global g_StandardLoadingBarEscPollPrev, g_OnEscapePressed

    g_StandardLoadingBarIsKeysOverlay := true
    g_StandardLoadingBarKeysHotkeys := []
    keyCallbacks := Map(
        "Enter", ShowCopiedResponse_DismissNoOp,
        "NumpadEnter", ShowCopiedResponse_DismissNoOp,
        "E", ShowCopiedResponse_DismissNoOp,
        "Escape", ShowCopiedResponse_DismissNoOp)

    try HotIf()
    catch {
    }
    for keyName, cb in keyCallbacks {
        if (StrLower(Trim(keyName)) = "escape")
            continue
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
        }
    }
    g_StandardLoadingBarKeysEscapeUserCb := ShowCopiedResponse_DismissNoOp
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
    g_StandardLoadingBarEscPollPrev := GetKeyState("Escape", "P") || (DllCall("user32\GetAsyncKeyState", "int", 0x1B) &
    0x8000)
    SetTimer(StandardLoadingBar_KeysEscapePoll, 50)
    g_StandardLoadingBarKeysEscapeActive := true
    StandardLoadingBar_StartKeysSelectionPoll(keyCallbacks)
    try {
        if IsObject(g_StandardLoadingBarGui) && g_StandardLoadingBarGui.Hwnd
            WinActivate(g_StandardLoadingBarGui.Hwnd)
    } catch {
    }
    try HotIf()
    catch {
    }
}
