; =============================================================================
; Utils module: focus_mode.ahk
; Focus mode multi-monitor blackout (#!+Y)
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; Focus Mode (multi-monitor blackout)
; Hotkey: Win+Alt+Shift+Y
; =============================================================================

global g_FocusModeOn := false
global g_FocusModeActiveMonitor := 0
global g_FocusModeOverlays := []  ; array of GUI overlays (one per covered monitor)
global g_FocusModeTrackedWindow := 0  ; window handle that was active when focus mode was enabled
global g_FocusModeAnchorHwnd := 0  ; window that may relocate blackout when it moves to another monitor
global g_FocusModeMonitorTimer := false  ; timer for monitoring window focus changes
global g_FocusModeKeepMonitorFile := A_ScriptDir "\.cursor\focus_mode_keep_monitor"
global g_FocusModeDisableRequestFile := A_ScriptDir "\.cursor\focus_mode_disable_request"

FocusMode_WriteKeepMonitorState(mon) {
    global g_FocusModeKeepMonitorFile
    if (!mon)
        return
    try {
        cursorDir := A_ScriptDir "\.cursor"
        if !DirExist(cursorDir)
            DirCreate(cursorDir)
        try FileDelete(g_FocusModeKeepMonitorFile)
        FileAppend(String(mon), g_FocusModeKeepMonitorFile, "UTF-8")
    } catch {
    }
}

FocusMode_ClearKeepMonitorState() {
    global g_FocusModeKeepMonitorFile, g_FocusModeDisableRequestFile
    try FileDelete(g_FocusModeKeepMonitorFile)
    catch {
    }
    try FileDelete(g_FocusModeDisableRequestFile)
    catch {
    }
}

FocusMode_ReadKeepMonitorFromFile() {
    global g_FocusModeKeepMonitorFile
    try {
        if !FileExist(g_FocusModeKeepMonitorFile)
            return 0
        t := Trim(FileRead(g_FocusModeKeepMonitorFile, "UTF-8"))
        if (t != "" && t is Integer)
            return Integer(t)
    } catch {
    }
    return 0
}

FocusMode_RequestDisableCrossProcess() {
    global g_FocusModeDisableRequestFile
    try {
        cursorDir := A_ScriptDir "\.cursor"
        if !DirExist(cursorDir)
            DirCreate(cursorDir)
        FileAppend("", g_FocusModeDisableRequestFile, "UTF-8")
    } catch {
    }
}

FocusMode_CheckCrossProcessRequests() {
    global g_FocusModeDisableRequestFile
    try {
        if !FileExist(g_FocusModeDisableRequestFile)
            return false
        FileDelete(g_FocusModeDisableRequestFile)
    } catch {
        return false
    }
    DisableFocusMode()
    return true
}

GetActiveMonitorIndex() {
    hwnd := WinExist("A")
    if (!hwnd) {
        return 0
    }

    rect := Buffer(16, 0)
    if (!DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect)) {
        return 0
    }

    left := NumGet(rect, 0, "int")
    top := NumGet(rect, 4, "int")
    right := NumGet(rect, 8, "int")
    bottom := NumGet(rect, 12, "int")

    centerX := left + (right - left) // 2
    centerY := top + (bottom - top) // 2

    monitorCount := MonitorGetCount()
    loop monitorCount {
        MonitorGet(A_Index, &ml, &mt, &mr, &mb)
        if (centerX >= ml && centerX <= mr && centerY >= mt && centerY <= mb) {
            return A_Index
        }
    }

    return 0
}

FocusMode_DestroyOverlays() {
    global g_FocusModeOverlays

    if (!IsSet(g_FocusModeOverlays))
        g_FocusModeOverlays := []
    for overlay in g_FocusModeOverlays {
        try {
            if (IsObject(overlay) && overlay.Hwnd)
                overlay.Destroy()
        } catch {
        }
    }
    g_FocusModeOverlays := []
}

FocusMode_BuildOverlays(keepMonitorIndex) {
    global g_FocusModeOverlays

    if (!keepMonitorIndex)
        return
    FocusMode_DestroyOverlays()

    monitorCount := MonitorGetCount()
    loop monitorCount {
        i := A_Index
        if (i = keepMonitorIndex)
            continue

        MonitorGet(i, &l, &t, &r, &b)
        w := r - l
        h := b - t
        if (w <= 0 || h <= 0)
            continue

        overlay := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x20") ; WS_EX_TRANSPARENT => click-through
        overlay.Opt("-DPIScale")
        overlay.BackColor := "000000"
        overlay.Show("NA x" l " y" t " w" w " h" h)
        g_FocusModeOverlays.Push(overlay)
    }
}

FocusMode_SetKeepMonitor(keepMonitorIndex) {
    global g_FocusModeOn, g_FocusModeActiveMonitor, g_FocusModeOverlays, g_FocusModeTrackedWindow

    if (!g_FocusModeOn || !keepMonitorIndex)
        return
    try {
        if (keepMonitorIndex < 1 || keepMonitorIndex > MonitorGetCount())
            return
    } catch {
        return
    }

    if (g_FocusModeActiveMonitor = keepMonitorIndex) {
        hasLiveOverlay := false
        for overlay in g_FocusModeOverlays {
            if (IsObject(overlay) && overlay.Hwnd && WinExist("ahk_id " . overlay.Hwnd)) {
                hasLiveOverlay := true
                break
            }
        }
        if (hasLiveOverlay)
            return
    }

    g_FocusModeActiveMonitor := keepMonitorIndex
    FocusMode_BuildOverlays(keepMonitorIndex)
    FocusMode_WriteKeepMonitorState(keepMonitorIndex)
    g_FocusModeTrackedWindow := WinExist("A")
}

EnableFocusMode(keepMonitorIndex := 0) {
    global g_FocusModeOn, g_FocusModeActiveMonitor, g_FocusModeOverlays, g_FocusModeTrackedWindow,
        g_FocusModeAnchorHwnd

    ; Ensure globals are initialized (avoid "variable has not been assigned" on first call)
    if (!IsSet(g_FocusModeOverlays))
        g_FocusModeOverlays := []
    if (!IsSet(g_FocusModeOn))
        g_FocusModeOn := false

    ; Flag plus overlay refs / WinExist: trust non-empty array like ToggleFocusMode (WinExist can miss fullscreen Gui overlays).
    hasOverlayRefs := IsObject(g_FocusModeOverlays) && g_FocusModeOverlays.Length > 0
    hasLiveOverlay := false
    if (hasOverlayRefs) {
        for overlay in g_FocusModeOverlays {
            if (IsObject(overlay) && overlay.Hwnd && WinExist("ahk_id " . overlay.Hwnd)) {
                hasLiveOverlay := true
                break
            }
        }
    }

    if (g_FocusModeOn || hasOverlayRefs) {
        if (hasOverlayRefs && !g_FocusModeOn) {
            DisableFocusMode()
        } else if (g_FocusModeOn && hasLiveOverlay) {
            return
        } else if (g_FocusModeOn && hasOverlayRefs && !hasLiveOverlay) {
            DisableFocusMode()
        }
        ; g_FocusModeOn && !hasOverlayRefs: inconsistent — fall through and recreate overlays
    }

    activeMon := keepMonitorIndex
    if (!activeMon)
        activeMon := GetActiveMonitorIndex()
    if (!activeMon) {
        return
    }

    g_FocusModeActiveMonitor := activeMon
    g_FocusModeTrackedWindow := WinExist("A")
    g_FocusModeAnchorHwnd := g_FocusModeTrackedWindow
    FocusMode_WriteKeepMonitorState(activeMon)
    StartFocusModeWindowMonitor()
    FocusMode_BuildOverlays(activeMon)
    g_FocusModeOn := true
}

DisableFocusMode() {
    global g_FocusModeOn, g_FocusModeActiveMonitor, g_FocusModeOverlays, g_FocusModeTrackedWindow,
        g_FocusModeMonitorTimer, g_FocusModeAnchorHwnd

    ; Stop monitoring window focus changes
    StopFocusModeWindowMonitor()
    StopPdfFocusMonitor()
    FocusMode_ClearKeepMonitorState()

    for overlay in g_FocusModeOverlays {
        try {
            if (IsObject(overlay) && overlay.Hwnd) {
                overlay.Destroy()
            }
        } catch {
            ; Ignore
        }
    }

    g_FocusModeOverlays := []
    g_FocusModeActiveMonitor := 0
    g_FocusModeTrackedWindow := 0
    g_FocusModeAnchorHwnd := 0
    g_FocusModeOn := false
}

ToggleFocusMode() {
    global g_FocusModeOn, g_FocusModeOverlays

    ; Treat non-empty overlay array as active even if WinExist fails on some setups (Dpi/multi-monitor),
    ; so #!+Y always tears down focus-mode state instead of calling EnableFocusMode() by mistake.
    hasOverlayRefs := IsObject(g_FocusModeOverlays) && g_FocusModeOverlays.Length > 0
    actualState := g_FocusModeOn || hasOverlayRefs

    if (actualState) {
        DisableFocusMode()
    } else {
        EnableFocusMode()
        try {
            if (MonitorGetCount() > 1) {
                hwnd := WinExist("A")
                if (hwnd)
                    StartPdfFocusMonitor(hwnd, "Immediate")
            }
        } catch {
        }
    }
}

; Keep tracked HWND in sync while foreground stays on the keep-clear monitor.
FocusModeWindowMonitor(*) {
    global g_FocusModeOn, g_FocusModeActiveMonitor, g_FocusModeTrackedWindow

    if (!g_FocusModeOn)
        return

    fg := WinExist("A")
    if (!fg)
        return

    fgMon := GetActiveMonitorIndex()
    if (fgMon && g_FocusModeActiveMonitor && fgMon = g_FocusModeActiveMonitor)
        g_FocusModeTrackedWindow := fg
}

; Start monitoring window focus changes
StartFocusModeWindowMonitor() {
    global g_FocusModeMonitorTimer

    ; Stop any existing timer first
    StopFocusModeWindowMonitor()

    ; Start timer to check window focus every 200ms
    g_FocusModeMonitorTimer := SetTimer(FocusModeWindowMonitor, 200)
}

; Stop monitoring window focus changes
StopFocusModeWindowMonitor() {
    global g_FocusModeMonitorTimer

    ; First-call safety: global may be unset
    if (!IsSet(g_FocusModeMonitorTimer))
        g_FocusModeMonitorTimer := false

    if (g_FocusModeMonitorTimer) {
        try {
            SetTimer(g_FocusModeMonitorTimer, 0)  ; Disable timer
        } catch {
            ; Ignore errors
        }
        g_FocusModeMonitorTimer := false
    }
}

#!+Y::
{
    ToggleFocusMode()
}
