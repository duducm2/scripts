; =============================================================================
; Utils module: desktop_recycle.ahk
; Desktop to Recycle Bin macro
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; Move all Desktop items to Recycle Bin (recoverable)
; Trigger: Ctrl+Alt+Win+8 (or RegisterMacro char N when selector existed)
; Target path: OneDrive Desktop. Standard banner with 4s timeout (N = cancel, Y or timeout = run); then success/error banner.
; During the 4s confirm: open Desktop Explorer if needed (hidden if we open it), capture a 60%-opacity
; frozen snapshot cropped to DWM visible frame, centered on the active window's monitor.
; =============================================================================
global g_DesktopToRecyclePath := ""  ; Set from GetDesktopToRecyclePath() when macro runs
global g_DesktopToRecycleCloseHwnd := 0
global g_DesktopToRecycleWeOpenedExplorer := false
global g_DesktopToRecycleSnapshotGui := 0
global g_DesktopToRecycleSnapshotHbm := 0

DesktopToRecycle_OnConfirm(*) {
    DesktopToRecycle_HideSnapshot()
    PlayCleaningDesktopSound()
    DesktopToRecycle_Run()
}

DesktopToRecycle_OnCancel(*) {
    DesktopToRecycle_HideSnapshot()
    DesktopToRecycle_CloseOpenedExplorerIfNeeded()
    ShowCenteredOverlay_Utils("⚠ Desktop cleanup cancelled", 1500, BANNER_ACCENT_INTERMEDIATE)
}

DesktopToRecycle_OnTimeout(*) {
    DesktopToRecycle_HideSnapshot()
    DesktopToRecycle_Run()
}

; Normalize folder path for comparison (trim trailing backslash, lowercase on Windows)
DesktopToRecycle_NormalizePath(p) {
    p := RTrim(p, "\")
    try return StrLower(p)
    return p
}

; Find Explorer hwnd showing targetPath via Shell.Application; else title Desktop / Área de Trabalho.
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

; Close Explorer only when this macro opened it (cancel / failed capture). Clears the flag.
DesktopToRecycle_CloseOpenedExplorerIfNeeded() {
    global g_DesktopToRecycleWeOpenedExplorer, g_DesktopToRecycleCloseHwnd
    if (!g_DesktopToRecycleWeOpenedExplorer) {
        g_DesktopToRecycleWeOpenedExplorer := false
        return
    }
    hwnd := g_DesktopToRecycleCloseHwnd
    g_DesktopToRecycleWeOpenedExplorer := false
    if (hwnd && WinExist("ahk_id " hwnd)) {
        try WinClose("ahk_id " hwnd)
        catch {
        }
    }
    g_DesktopToRecycleCloseHwnd := 0
}

; Ensure Desktop Explorer is open/restored for targetPath. Returns hwnd or 0.
; If this call opens Explorer, hides it after a short paint settle (PrintWindow still works).
DesktopToRecycle_EnsureDesktopExplorer(targetPath) {
    global g_DesktopToRecycleWeOpenedExplorer
    g_DesktopToRecycleWeOpenedExplorer := false
    hwnd := DesktopToRecycle_FindDesktopExplorer(targetPath)
    weOpened := false
    if (!hwnd && targetPath && DirExist(targetPath)) {
        try Run('explorer.exe "' targetPath '"')
        catch {
            return 0
        }
        weOpened := true
        deadline := A_TickCount + 2000
        while (A_TickCount < deadline) {
            hwnd := DesktopToRecycle_FindDesktopExplorer(targetPath)
            if (hwnd)
                break
            Sleep 50
        }
    }
    if (!hwnd || !WinExist("ahk_id " hwnd)) {
        g_DesktopToRecycleWeOpenedExplorer := false
        return 0
    }
    try {
        if (WinGetMinMax("ahk_id " hwnd) = -1)
            WinRestore("ahk_id " hwnd)
    } catch {
    }
    ; Brief settle so the items view can paint before PrintWindow / hide.
    Sleep 150
    if (!WinExist("ahk_id " hwnd)) {
        g_DesktopToRecycleWeOpenedExplorer := false
        return 0
    }
    if (weOpened) {
        try WinHide("ahk_id " hwnd)
        catch {
        }
        g_DesktopToRecycleWeOpenedExplorer := true
    }
    return hwnd
}

; Visible frame via DWMWA_EXTENDED_FRAME_BOUNDS (excludes drop-shadow). Returns false on failure.
DesktopToRecycle_GetDwmExtendedFrame(hwnd, &left, &top, &right, &bottom) {
    left := top := right := bottom := 0
    rc := Buffer(16, 0)
    ; DWMWA_EXTENDED_FRAME_BOUNDS = 9
    if (DllCall("dwmapi\DwmGetWindowAttribute", "ptr", hwnd, "uint", 9, "ptr", rc, "uint", 16) != 0)
        return false
    left := NumGet(rc, 0, "int")
    top := NumGet(rc, 4, "int")
    right := NumGet(rc, 8, "int")
    bottom := NumGet(rc, 12, "int")
    return (right > left && bottom > top)
}

; Crop fullHbm (winW x winH) to the DWM visible frame offset within GetWindowRect. Returns new HBITMAP or 0.
DesktopToRecycle_CropBitmapToDwmFrame(fullHbm, hdcRef, winL, winT, winW, winH, hwnd) {
    if (!fullHbm || !hdcRef)
        return 0
    if (!DesktopToRecycle_GetDwmExtendedFrame(hwnd, &extL, &extT, &extR, &extB))
        return 0
    ox := extL - winL
    oy := extT - winT
    cw := extR - extL
    ch := extB - extT
    if (cw < 1 || ch < 1 || ox < 0 || oy < 0 || ox + cw > winW || oy + ch > winH)
        return 0
    hdcSrc := DllCall("CreateCompatibleDC", "ptr", hdcRef, "ptr")
    if (!hdcSrc)
        return 0
    hdcDst := 0
    cropHbm := 0
    try {
        oldSrc := DllCall("SelectObject", "ptr", hdcSrc, "ptr", fullHbm, "ptr")
        hdcDst := DllCall("CreateCompatibleDC", "ptr", hdcRef, "ptr")
        if (!hdcDst) {
            DllCall("SelectObject", "ptr", hdcSrc, "ptr", oldSrc, "ptr")
            return 0
        }
        cropHbm := DllCall("CreateCompatibleBitmap", "ptr", hdcRef, "int", cw, "int", ch, "ptr")
        if (!cropHbm) {
            DllCall("SelectObject", "ptr", hdcSrc, "ptr", oldSrc, "ptr")
            return 0
        }
        oldDst := DllCall("SelectObject", "ptr", hdcDst, "ptr", cropHbm, "ptr")
        ; SRCCOPY = 0x00CC0020
        DllCall("BitBlt", "ptr", hdcDst, "int", 0, "int", 0, "int", cw, "int", ch,
            "ptr", hdcSrc, "int", ox, "int", oy, "uint", 0x00CC0020)
        DllCall("SelectObject", "ptr", hdcDst, "ptr", oldDst, "ptr")
        DllCall("SelectObject", "ptr", hdcSrc, "ptr", oldSrc, "ptr")
        return cropHbm
    } finally {
        if (hdcDst)
            DllCall("DeleteDC", "ptr", hdcDst)
        DllCall("DeleteDC", "ptr", hdcSrc)
    }
}

; Capture hwnd to HBITMAP via PrintWindow; crop DWM drop-shadow when possible. Returns 0 on failure.
DesktopToRecycle_CaptureWindowBitmap(hwnd) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return 0
    rect := Buffer(16, 0)
    if (!DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect))
        return 0
    winL := NumGet(rect, 0, "int")
    winT := NumGet(rect, 4, "int")
    winR := NumGet(rect, 8, "int")
    winB := NumGet(rect, 12, "int")
    w := winR - winL
    h := winB - winT
    if (w < 1 || h < 1)
        return 0
    hdcWin := DllCall("GetWindowDC", "ptr", hwnd, "ptr")
    if (!hdcWin)
        return 0
    hdcMem := 0
    hbm := 0
    try {
        hdcMem := DllCall("CreateCompatibleDC", "ptr", hdcWin, "ptr")
        if (!hdcMem)
            return 0
        hbm := DllCall("CreateCompatibleBitmap", "ptr", hdcWin, "int", w, "int", h, "ptr")
        if (!hbm)
            return 0
        oldBm := DllCall("SelectObject", "ptr", hdcMem, "ptr", hbm, "ptr")
        ; PW_RENDERFULLCONTENT = 2 — better for DWM/Explorer chrome (works while hidden after paint)
        ok := DllCall("PrintWindow", "ptr", hwnd, "ptr", hdcMem, "uint", 2)
        DllCall("SelectObject", "ptr", hdcMem, "ptr", oldBm, "ptr")
        if (!ok) {
            DllCall("DeleteObject", "ptr", hbm)
            return 0
        }
        cropped := DesktopToRecycle_CropBitmapToDwmFrame(hbm, hdcWin, winL, winT, w, h, hwnd)
        if (cropped) {
            DllCall("DeleteObject", "ptr", hbm)
            return cropped
        }
        return hbm
    } finally {
        if (hdcMem)
            DllCall("DeleteDC", "ptr", hdcMem)
        DllCall("ReleaseDC", "ptr", hwnd, "ptr", hdcWin)
    }
}

DesktopToRecycle_HideSnapshot() {
    global g_DesktopToRecycleSnapshotGui, g_DesktopToRecycleSnapshotHbm
    try {
        if IsObject(g_DesktopToRecycleSnapshotGui)
            g_DesktopToRecycleSnapshotGui.Destroy()
    } catch {
    }
    g_DesktopToRecycleSnapshotGui := 0
    if (g_DesktopToRecycleSnapshotHbm) {
        try DllCall("DeleteObject", "ptr", g_DesktopToRecycleSnapshotHbm)
        catch {
        }
        g_DesktopToRecycleSnapshotHbm := 0
    }
}

; AlwaysOnTop snapshot Gui, centered on the given monitor work area (same helpers as StandardLoadingBar).
; Opacity 60% (WinSetTransparent 153). Returns true on success.
; workLeft/Top/Right/Bottom: capture before opening Explorer so focus steal does not move the target monitor.
DesktopToRecycle_ShowSnapshot(hwnd, workLeft, workTop, workRight, workBottom) {
    global g_DesktopToRecycleSnapshotGui, g_DesktopToRecycleSnapshotHbm
    DesktopToRecycle_HideSnapshot()
    hbm := DesktopToRecycle_CaptureWindowBitmap(hwnd)
    if (!hbm)
        return false
    ; Bitmap pixel size from object header (avoids DPI / window-rect mismatch that left a white frame).
    bm := Buffer(32, 0)  ; BITMAP
    if (!DllCall("GetObject", "ptr", hbm, "int", bm.Size, "ptr", bm)) {
        DllCall("DeleteObject", "ptr", hbm)
        return false
    }
    srcW := NumGet(bm, 4, "int")   ; bmWidth
    srcH := NumGet(bm, 8, "int")   ; bmHeight
    if (srcW < 1 || srcH < 1) {
        DllCall("DeleteObject", "ptr", hbm)
        return false
    }
    monW := workRight - workLeft
    monH := workBottom - workTop
    if (monW < 1 || monH < 1) {
        DllCall("DeleteObject", "ptr", hbm)
        return false
    }
    ; Fit inside work area; keep aspect ratio; center (same placement idea as StandardLoadingBar_Show).
    scale := Min(1, monW / srcW, monH / srcH)
    dispW := Max(1, Round(srcW * scale))
    dispH := Max(1, Round(srcH * scale))

    ; Dark BackColor + zero margins: default white Gui chrome was showing as a block around the Picture.
    snapGui := Gui("+AlwaysOnTop -Caption +ToolWindow -DPIScale")
    snapGui.BackColor := "1E1E2E"
    snapGui.MarginX := 0
    snapGui.MarginY := 0
    ; HBITMAP:* — Gui does not take ownership; we DeleteObject in HideSnapshot.
    try snapGui.Add("Picture", "x0 y0 w" dispW " h" dispH, "HBITMAP:*" hbm)
    catch {
        try snapGui.Destroy()
        catch {
        }
        DllCall("DeleteObject", "ptr", hbm)
        return false
    }
    ; AutoSize to the Picture, then center — avoids forced w/h outer/client mismatch.
    snapGui.Show("Hide AutoSize")
    snapGui.GetPos(, , &gw, &gh)
    guiX := Round(workLeft + (monW - gw) / 2)
    guiY := Round(workTop + (monH - gh) / 2)
    snapGui.Show("NA x" guiX " y" guiY)
    ; 60% opacity (0 = invisible, 255 = fully opaque)
    try WinSetTransparent(153, snapGui)
    catch {
    }
    g_DesktopToRecycleSnapshotGui := snapGui
    g_DesktopToRecycleSnapshotHbm := hbm
    return true
}

; Close any Explorer window(s) showing the given folder path (via Shell.Application)
DesktopToRecycle_CloseDesktopExplorer(targetPath) {
    global g_DesktopToRecycleWeOpenedExplorer, g_DesktopToRecycleCloseHwnd
    if (!targetPath || targetPath = "") {
        g_DesktopToRecycleWeOpenedExplorer := false
        return
    }
    normTarget := DesktopToRecycle_NormalizePath(targetPath)
    try {
        shell := ComObject("Shell.Application")
        for window in shell.Windows {
            try {
                if (!window || !window.hwnd)
                    continue
                path := window.Document.Folder.Self.Path
                if (DesktopToRecycle_NormalizePath(path) = normTarget) {
                    window.Quit()
                    g_DesktopToRecycleCloseHwnd := 0
                    g_DesktopToRecycleWeOpenedExplorer := false
                    return
                }
            } catch
                continue
        }
    } catch {
    }
    ; Fallback: close by hwnd if we had stored it at trigger time
    if (g_DesktopToRecycleCloseHwnd && WinExist("ahk_id " g_DesktopToRecycleCloseHwnd)) {
        try WinClose("ahk_id " g_DesktopToRecycleCloseHwnd)
    }
    g_DesktopToRecycleCloseHwnd := 0
    g_DesktopToRecycleWeOpenedExplorer := false
}

DesktopToRecycle_Run() {
    global g_DesktopToRecyclePath, g_DesktopToRecycleCloseHwnd, g_DesktopToRecycleWeOpenedExplorer
    ; Resolve path: use configured path; if empty or missing, fall back to A_Desktop (works on any PC)
    path := g_DesktopToRecyclePath
    if (!path || path = "" || !DirExist(path))
        path := A_Desktop
    ; Use .NET FileIO.FileSystem SendToRecycleBin (no Shell verbs); process dirs last so parent exists
    ui := "[Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs"
    rec := "[Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin"
    ps := "Add-Type -AssemblyName Microsoft.VisualBasic;$d='" . path .
        "';if(-not(Test-Path -LiteralPath $d)){exit 1};$files=@(Get-ChildItem -LiteralPath $d -Force|Where-Object{-not $_.PSIsContainer});$dirs=@(Get-ChildItem -LiteralPath $d -Force|Where-Object{$_.PSIsContainer});foreach($f in $files){try{[Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($f.FullName," .
        ui . "," . rec .
        ")}catch{}};foreach($dir in $dirs){try{[Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory($dir.FullName," .
        ui . "," . rec . ")}catch{}};exit 0"
    try {
        exitCode := RunWait('powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "' . ps . '"', "", "Hide")
        if (exitCode = 0) {
            ShowCenteredOverlay_Utils("✅ Desktop items moved to Recycle Bin", 2000, BANNER_ACCENT_SUCCESS)
            DesktopToRecycle_CloseDesktopExplorer(path)
        } else {
            ShowCenteredOverlay_Utils("❌ Desktop path not found or error: " path, 3500, BANNER_ACCENT_ERROR)
            DesktopToRecycle_CloseDesktopExplorer(path)
        }
    } catch as err {
        ShowCenteredOverlay_Utils("❌ Error moving to Recycle Bin", 2500, BANNER_ACCENT_ERROR)
        DesktopToRecycle_CloseOpenedExplorerIfNeeded()
    }
    g_DesktopToRecycleCloseHwnd := 0
    g_DesktopToRecycleWeOpenedExplorer := false
}

; Entry point for Desktop to Recycle macro (^!#8)
DesktopToRecycle_Trigger() {
    global g_DesktopToRecycleCloseHwnd, g_DesktopToRecyclePath, g_DesktopToRecycleWeOpenedExplorer
    DesktopToRecycle_HideSnapshot()
    g_DesktopToRecycleWeOpenedExplorer := false
    g_DesktopToRecyclePath := GetDesktopToRecyclePath()
    path := g_DesktopToRecyclePath
    if (!path || path = "" || !DirExist(path))
        path := A_Desktop

    ; Work area of the monitor with the *current* active window — before Explorer open steals focus.
    GetActiveMonitorWorkArea_StandardBar(&workLeft, &workTop, &workRight, &workBottom)

    StandardLoadingBar_CloseKeysOverlay()
    StandardLoadingBar_Hide(0)
    Sleep 50

    hwnd := DesktopToRecycle_EnsureDesktopExplorer(path)
    g_DesktopToRecycleCloseHwnd := hwnd ? hwnd : 0
    if (hwnd) {
        if (!DesktopToRecycle_ShowSnapshot(hwnd, workLeft, workTop, workRight, workBottom))
            DesktopToRecycle_CloseOpenedExplorerIfNeeded()
    }

    state := "🗑️ Move all items from:`n" . g_DesktopToRecyclePath . "`nto Recycle Bin? (4s)"
    keyCallbacks := Map(
        "Y", DesktopToRecycle_OnConfirm,
        "N", DesktopToRecycle_OnCancel,
        "Escape", DesktopToRecycle_OnCancel)
    ; Center on active monitor (centerOnHwnd := 0), use standard intermediate accent with border.
    ; Snapshot is shown first (AlwaysOnTop); banner stacks above and owns keys.
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
