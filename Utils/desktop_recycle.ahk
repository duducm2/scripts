; =============================================================================
; Utils module: desktop_recycle.ahk
; Desktop to Recycle Bin macro
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; Move all Desktop items to Recycle Bin (recoverable)
; Trigger: Ctrl+Alt+Win+8
; Opens a temporary Desktop Explorer at 50% size / 50% opacity, centered on the
; active window's monitor; tracks focus for 4s and recenters if the user switches
; monitors. Y / timeout = recycle; N / Escape = cancel. Preview hwnd is marked
; with window prop DesktopToRecycleTempExclude so AutoSlot skips it (cross-process).
; =============================================================================
global g_DesktopToRecyclePath := ""
global g_DesktopToRecycleCloseHwnd := 0
global g_DesktopToRecycleWeOpenedExplorer := false
global g_DesktopToRecycleTrackTimer := ""
global g_DesktopToRecycleTrackLastMonIdx := 0
global DESKTOP_TO_RECYCLE_AUTOSLOT_PROP := "DesktopToRecycleTempExclude"
global DESKTOP_TO_RECYCLE_TRACK_INTERVAL := 115
global DESKTOP_TO_RECYCLE_PREVIEW_OPACITY := 128  ; 50% of 255
global DESKTOP_TO_RECYCLE_PREVIEW_SCALE := 0.5

DesktopToRecycle_OnConfirm(*) {
    DesktopToRecycle_StopTrack()
    DesktopToRecycle_ClosePreviewExplorer()
    PlayCleaningDesktopSound()
    DesktopToRecycle_Run()
}

DesktopToRecycle_OnCancel(*) {
    DesktopToRecycle_StopTrack()
    DesktopToRecycle_ClosePreviewExplorer()
    ShowCenteredOverlay_Utils("⚠ Desktop cleanup cancelled", 1500, BANNER_ACCENT_INTERMEDIATE)
}

DesktopToRecycle_OnTimeout(*) {
    DesktopToRecycle_StopTrack()
    DesktopToRecycle_ClosePreviewExplorer()
    DesktopToRecycle_Run()
}

DesktopToRecycle_NormalizePath(p) {
    p := RTrim(p, "\")
    try return StrLower(p)
    return p
}

; Mark hwnd so AutoSlot (WindowManagement process) skips Place/occupancy via GetProp.
DesktopToRecycle_MarkAutoSlotExclude(hwnd) {
    global DESKTOP_TO_RECYCLE_AUTOSLOT_PROP
    if (!hwnd)
        return
    try DllCall("SetPropW", "ptr", hwnd, "wstr", DESKTOP_TO_RECYCLE_AUTOSLOT_PROP, "ptr", 1)
    catch {
    }
}

DesktopToRecycle_ClearAutoSlotExclude(hwnd) {
    global DESKTOP_TO_RECYCLE_AUTOSLOT_PROP
    if (!hwnd)
        return
    try DllCall("RemovePropW", "ptr", hwnd, "wstr", DESKTOP_TO_RECYCLE_AUTOSLOT_PROP)
    catch {
    }
}

; Cross-process suppress file (Utils + WindowManagement/AutoSlot share A_ScriptDir).
DesktopToRecycle_AutoSlotSuppressPath() {
    return A_ScriptDir "\assets\data\desktop_recycle_autoslot_suppress.ini"
}

; Call BEFORE Run explorer so AutoSlot debounce cannot Place the new window.
DesktopToRecycle_BeginAutoSlotSuppress(durationMs := 12000) {
    path := DesktopToRecycle_AutoSlotSuppressPath()
    try DirCreate(A_ScriptDir "\assets\data")
    catch {
    }
    until := DllCall("GetTickCount", "UInt") + durationMs
    try {
        IniWrite(until, path, "Suppress", "Until")
        IniWrite(1, path, "Suppress", "Active")
    } catch {
    }
}

DesktopToRecycle_EndAutoSlotSuppress() {
    path := DesktopToRecycle_AutoSlotSuppressPath()
    try {
        IniWrite(0, path, "Suppress", "Until")
        IniWrite(0, path, "Suppress", "Active")
    } catch {
    }
}

; Used by AutoSlot_IsExcludedExeOrTitle (same process when WM includes Utils, and via file).
DesktopToRecycle_AutoSlotSuppressActive() {
    path := DesktopToRecycle_AutoSlotSuppressPath()
    try {
        active := Integer(IniRead(path, "Suppress", "Active", 0))
        until := Integer(IniRead(path, "Suppress", "Until", 0))
    } catch {
        return false
    }
    if (!active || until < 1)
        return false
    return DllCall("GetTickCount", "UInt") < until
}

DesktopToRecycle_IsDesktopExplorerTitle(title) {
    if (title = "")
        return false
    return InStr(title, "Desktop", false) || InStr(title, "Área de Trabalho", false)
}

; Find Explorer hwnd showing targetPath; else title Desktop / Área de Trabalho.
DesktopToRecycle_FindDesktopExplorer(targetPath) {
    if (targetPath && targetPath != "") {
        normTarget := DesktopToRecycle_NormalizePath(targetPath)
        try {
            shell := ComObject("Shell.Application")
            for window in shell.Windows {
                try {
                    if (!window || !window.hwnd)
                        continue
                    path := window.Document.Folder.Self.Path
                    if (DesktopToRecycle_NormalizePath(path) = normTarget)
                        return Integer(window.hwnd)
                } catch
                    continue
            }
        } catch {
        }
    }
    prevMode := A_TitleMatchMode
    try {
        SetTitleMatchMode 2
        hwnd := WinExist("Área de Trabalho ahk_class CabinetWClass")
        if (hwnd)
            return hwnd
        return WinExist("Desktop ahk_class CabinetWClass")
    } finally {
        SetTitleMatchMode prevMode
    }
}

; Snapshot of Desktop Explorer hwnds before we launch (to prefer a newly created one).
DesktopToRecycle_CollectDesktopExplorerHwnds(targetPath) {
    found := Map()
    if (targetPath && targetPath != "") {
        normTarget := DesktopToRecycle_NormalizePath(targetPath)
        try {
            shell := ComObject("Shell.Application")
            for window in shell.Windows {
                try {
                    if (!window || !window.hwnd)
                        continue
                    path := window.Document.Folder.Self.Path
                    if (DesktopToRecycle_NormalizePath(path) = normTarget)
                        found[Integer(window.hwnd)] := true
                } catch
                    continue
            }
        } catch {
        }
    }
    return found
}

; Center hwnd at 50% of the given work area; apply 50% opacity.
DesktopToRecycle_PlacePreviewOnWorkArea(hwnd, workLeft, workTop, workRight, workBottom) {
    global DESKTOP_TO_RECYCLE_PREVIEW_OPACITY, DESKTOP_TO_RECYCLE_PREVIEW_SCALE
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    monW := workRight - workLeft
    monH := workBottom - workTop
    if (monW < 1 || monH < 1)
        return false
    w := Max(200, Round(monW * DESKTOP_TO_RECYCLE_PREVIEW_SCALE))
    h := Max(150, Round(monH * DESKTOP_TO_RECYCLE_PREVIEW_SCALE))
    x := Round(workLeft + (monW - w) / 2)
    y := Round(workTop + (monH - h) / 2)
    try {
        if (WinGetMinMax("ahk_id " hwnd) = -1)
            WinRestore("ahk_id " hwnd)
    } catch {
    }
    try WinMove(x, y, w, h, "ahk_id " hwnd)
    catch {
        return false
    }
    try WinSetTransparent(DESKTOP_TO_RECYCLE_PREVIEW_OPACITY, "ahk_id " hwnd)
    catch {
    }
    return true
}

DesktopToRecycle_PlacePreviewOnMonitor(hwnd, monIdx) {
    if (!hwnd || monIdx < 1)
        return false
    try MonitorGetWorkArea(monIdx, &l, &t, &r, &b)
    catch {
        return false
    }
    return DesktopToRecycle_PlacePreviewOnWorkArea(hwnd, l, t, r, b)
}

DesktopToRecycle_StopTrack() {
    global g_DesktopToRecycleTrackTimer, g_DesktopToRecycleTrackLastMonIdx
    try SetTimer(DesktopToRecycle_TrackTick, 0)
    catch {
    }
    g_DesktopToRecycleTrackTimer := ""
    g_DesktopToRecycleTrackLastMonIdx := 0
}

; Follow foreground window's monitor (ignore when preview itself is active).
DesktopToRecycle_TrackTick(*) {
    global g_DesktopToRecycleCloseHwnd, g_DesktopToRecycleTrackLastMonIdx
    hwnd := g_DesktopToRecycleCloseHwnd
    if (!hwnd || !WinExist("ahk_id " hwnd)) {
        DesktopToRecycle_StopTrack()
        return
    }
    fg := 0
    try fg := WinGetID("A")
    catch {
        fg := 0
    }
    if (fg && fg = hwnd)
        return
    monIdx := GetMonitorIndexForForeground_StandardBar()
    if (monIdx < 1)
        return
    if (monIdx = g_DesktopToRecycleTrackLastMonIdx)
        return
    if (DesktopToRecycle_PlacePreviewOnMonitor(hwnd, monIdx))
        g_DesktopToRecycleTrackLastMonIdx := monIdx
}

DesktopToRecycle_StartTrack(hwnd, initialMonIdx) {
    global g_DesktopToRecycleTrackTimer, g_DesktopToRecycleTrackLastMonIdx, DESKTOP_TO_RECYCLE_TRACK_INTERVAL
    DesktopToRecycle_StopTrack()
    g_DesktopToRecycleTrackLastMonIdx := initialMonIdx
    SetTimer(DesktopToRecycle_TrackTick, DESKTOP_TO_RECYCLE_TRACK_INTERVAL)
    g_DesktopToRecycleTrackTimer := DesktopToRecycle_TrackTick
}

; Close only the temporary preview Explorer hwnd (not every Desktop Explorer).
DesktopToRecycle_ClosePreviewExplorer() {
    global g_DesktopToRecycleCloseHwnd, g_DesktopToRecycleWeOpenedExplorer
    DesktopToRecycle_StopTrack()
    hwnd := g_DesktopToRecycleCloseHwnd
    g_DesktopToRecycleCloseHwnd := 0
    g_DesktopToRecycleWeOpenedExplorer := false
    if (!hwnd)
        return
    DesktopToRecycle_ClearAutoSlotExclude(hwnd)
    if (WinExist("ahk_id " hwnd)) {
        try WinClose("ahk_id " hwnd)
        catch {
            try WinKill("ahk_id " hwnd)
            catch {
            }
        }
    }
}

; Open a new Desktop Explorer, exclude from AutoSlot, place at 50%/50% opacity. Returns hwnd or 0.
DesktopToRecycle_OpenPreviewExplorer(targetPath, workLeft, workTop, workRight, workBottom) {
    global g_DesktopToRecycleWeOpenedExplorer
    g_DesktopToRecycleWeOpenedExplorer := false
    if (!targetPath || !DirExist(targetPath))
        return 0

    before := DesktopToRecycle_CollectDesktopExplorerHwnds(targetPath)
    ; /n prefers a new window rather than reusing an existing folder view.
    try Run('explorer.exe /n,"' targetPath '"')
    catch {
        try Run('explorer.exe "' targetPath '"')
        catch {
            return 0
        }
    }

    hwnd := 0
    deadline := A_TickCount + 2500
    while (A_TickCount < deadline) {
        try {
            shell := ComObject("Shell.Application")
            for window in shell.Windows {
                try {
                    if (!window || !window.hwnd)
                        continue
                    h := Integer(window.hwnd)
                    path := window.Document.Folder.Self.Path
                    if (DesktopToRecycle_NormalizePath(path) != DesktopToRecycle_NormalizePath(targetPath))
                        continue
                    if (!before.Has(h)) {
                        hwnd := h
                        break
                    }
                } catch
                    continue
            }
        } catch {
        }
        if (hwnd)
            break
        ; Fallback: any Desktop Explorer if we cannot detect "new"
        cand := DesktopToRecycle_FindDesktopExplorer(targetPath)
        if (cand && !before.Has(cand)) {
            hwnd := cand
            break
        }
        Sleep 50
    }
    if (!hwnd)
        hwnd := DesktopToRecycle_FindDesktopExplorer(targetPath)
    ; Never adopt a pre-existing Desktop Explorer (would resize/opacity/close the user's window).
    if (!hwnd || !WinExist("ahk_id " hwnd) || before.Has(hwnd))
        return 0

    ; Exclude before place so AutoSlot SHOW debounce never Places this hwnd.
    DesktopToRecycle_MarkAutoSlotExclude(hwnd)
    g_DesktopToRecycleWeOpenedExplorer := true

    if (!DesktopToRecycle_PlacePreviewOnWorkArea(hwnd, workLeft, workTop, workRight, workBottom)) {
        DesktopToRecycle_ClearAutoSlotExclude(hwnd)
        try WinClose("ahk_id " hwnd)
        catch {
        }
        g_DesktopToRecycleWeOpenedExplorer := false
        return 0
    }
    return hwnd
}

DesktopToRecycle_Run() {
    global g_DesktopToRecyclePath
    path := g_DesktopToRecyclePath
    if (!path || path = "" || !DirExist(path))
        path := A_Desktop
    ui := "[Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs"
    rec := "[Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin"
    ps := "Add-Type -AssemblyName Microsoft.VisualBasic;$d='" . path .
        "';if(-not(Test-Path -LiteralPath $d)){exit 1};$files=@(Get-ChildItem -LiteralPath $d -Force|Where-Object{-not $_.PSIsContainer});$dirs=@(Get-ChildItem -LiteralPath $d -Force|Where-Object{$_.PSIsContainer});foreach($f in $files){try{[Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($f.FullName," .
        ui . "," . rec .
        ")}catch{}};foreach($dir in $dirs){try{[Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory($dir.FullName," .
        ui . "," . rec . ")}catch{}};exit 0"
    try {
        exitCode := RunWait('powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "' . ps . '"', "", "Hide")
        if (exitCode = 0)
            ShowCenteredOverlay_Utils("✅ Desktop items moved to Recycle Bin", 2000, BANNER_ACCENT_SUCCESS)
        else
            ShowCenteredOverlay_Utils("❌ Desktop path not found or error: " path, 3500, BANNER_ACCENT_ERROR)
    } catch as err {
        ShowCenteredOverlay_Utils("❌ Error moving to Recycle Bin", 2500, BANNER_ACCENT_ERROR)
    }
}

; Entry point for Desktop to Recycle macro (^!#8)
DesktopToRecycle_Trigger() {
    global g_DesktopToRecycleCloseHwnd, g_DesktopToRecyclePath, g_DesktopToRecycleWeOpenedExplorer
    DesktopToRecycle_StopTrack()
    DesktopToRecycle_ClosePreviewExplorer()
    g_DesktopToRecycleWeOpenedExplorer := false
    g_DesktopToRecyclePath := GetDesktopToRecyclePath()
    path := g_DesktopToRecyclePath
    if (!path || path = "" || !DirExist(path))
        path := A_Desktop

    ; Work area of the monitor with the *current* active window — before Explorer steals focus.
    GetActiveMonitorWorkArea_StandardBar(&workLeft, &workTop, &workRight, &workBottom)
    initialMonIdx := GetMonitorIndexForForeground_StandardBar()

    StandardLoadingBar_CloseKeysOverlay()
    StandardLoadingBar_Hide(0)
    Sleep 50

    hwnd := DesktopToRecycle_OpenPreviewExplorer(path, workLeft, workTop, workRight, workBottom)
    g_DesktopToRecycleCloseHwnd := hwnd ? hwnd : 0
    if (hwnd)
        DesktopToRecycle_StartTrack(hwnd, initialMonIdx)

    state := "🗑️ Move all items from:`n" . g_DesktopToRecyclePath . "`nto Recycle Bin? (4s)"
    keyCallbacks := Map(
        "Y", DesktopToRecycle_OnConfirm,
        "N", DesktopToRecycle_OnCancel,
        "Escape", DesktopToRecycle_OnCancel)
    StandardLoadingBar_ShowWithKeys(
        state,
        keyCallbacks,
        4000,
        0,
        DesktopToRecycle_OnTimeout,
        BANNER_ACCENT_INTERMEDIATE,
        0,
        17,
        "",
        false,
        "[Y] Yes  [N] Cancel",
        true,
        true,
        true)
}
