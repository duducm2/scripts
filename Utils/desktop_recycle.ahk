; =============================================================================
; Utils module: desktop_recycle.ahk
; Desktop to Recycle Bin macro
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; Move all Desktop items to Recycle Bin (recoverable)
; Trigger: Win+Alt+Shift+U selector → letter N
; Target path: OneDrive Desktop. Standard banner with 4s timeout (N = cancel, Y or timeout = run); then success/error banner.
; =============================================================================
global g_DesktopToRecyclePath := ""  ; Set from GetDesktopToRecyclePath() when macro runs
global g_DesktopToRecycleCloseHwnd := 0

DesktopToRecycle_OnConfirm(*) {
    PlayCleaningDesktopSound()
    DesktopToRecycle_Run()
}

DesktopToRecycle_OnCancel(*) {
    ShowCenteredOverlay_Utils("⚠ Desktop cleanup cancelled", 1500, BANNER_ACCENT_INTERMEDIATE)
}

DesktopToRecycle_OnTimeout(*) {
    DesktopToRecycle_Run()
}

; Normalize folder path for comparison (trim trailing backslash, lowercase on Windows)
DesktopToRecycle_NormalizePath(p) {
    p := RTrim(p, "\")
    try return StrLower(p)
    return p
}

; Close any Explorer window(s) showing the given folder path (via Shell.Application)
DesktopToRecycle_CloseDesktopExplorer(targetPath) {
    if (!targetPath || targetPath = "")
        return
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
                    return
                }
            } catch
                continue
        }
    } catch {
    }
    ; Fallback: close by hwnd if we had stored it at trigger time
    global g_DesktopToRecycleCloseHwnd
    if (g_DesktopToRecycleCloseHwnd && WinExist("ahk_id " g_DesktopToRecycleCloseHwnd)) {
        try WinClose("ahk_id " g_DesktopToRecycleCloseHwnd)
    }
    g_DesktopToRecycleCloseHwnd := 0
}

DesktopToRecycle_Run() {
    global g_DesktopToRecyclePath, g_DesktopToRecycleCloseHwnd
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
    }
    g_DesktopToRecycleCloseHwnd := 0
}

; Entry point when "N" is pressed in Win+Alt+Shift+U selector
DesktopToRecycle_Trigger() {
    global g_DesktopToRecycleCloseHwnd, g_DesktopToRecyclePath
    g_DesktopToRecyclePath := GetDesktopToRecyclePath()
    ; Remember active window if it's Explorer showing Desktop - close it after cleaning
    hwnd := WinExist("A")
    g_DesktopToRecycleCloseHwnd := 0
    if (hwnd && WinGetProcessName("ahk_id " hwnd) = "explorer.exe") {
        try {
            if (InStr(WinGetTitle("ahk_id " hwnd), "Desktop", false))
                g_DesktopToRecycleCloseHwnd := hwnd
        } catch {
        }
    }
    StandardLoadingBar_CloseKeysOverlay()
    StandardLoadingBar_Hide(0)
    Sleep 50
    state := "🗑️ Move all items from:`n" . g_DesktopToRecyclePath . "`nto Recycle Bin? (4s)"
    keyCallbacks := Map("Y", DesktopToRecycle_OnConfirm, "N", DesktopToRecycle_OnCancel)
    ; Center on active monitor (centerOnHwnd := 0), use standard intermediate accent with border.
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
