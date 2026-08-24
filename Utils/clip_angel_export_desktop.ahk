; =============================================================================
; Utils module: clip_angel_export_desktop.ahk
; Copy Clip Angel Row 0 (newest clip) preview text to a UTF-8 .txt on Desktop.
; Utility Shortcuts: #!+U → Macros → [c]
; =============================================================================

ClipAngelExport_UniqueDesktopPath(desktopDir) {
    stamp := FormatTime(, "yyyyMMdd-HHmmss")
    base := "clipangel-last-" stamp
    path := desktopDir "\" base ".txt"
    if !FileExist(path)
        return path
    i := 2
    loop {
        path := desktopDir "\" base "-" i ".txt"
        if !FileExist(path)
            return path
        i += 1
    }
}

ClipAngel_ExportLastClipToDesktop() {
    if !ClipAngel_TryAcquireAutomationLock()
        return
    savedClip := ClipboardAll()
    try {
        StandardLoadingBar_Show("⏳ Clip Angel: exporting...", BANNER_ACCENT_INTERMEDIATE)
        ClipAngel_ActivateNativeFirstClip()
        if !ClipAngel_WaitForListReady(CLIPANGEL_FAVORITE_OPEN_READY_MS, true) {
            ShowCenteredOverlay_Utils("❌ Clip Angel did not open.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        clipHwnd := ClipAngel_MainHwnd()
        if (clipHwnd)
            ClipAngel_LeaveFavoritesFilter(clipHwnd)
        ClipAngel_SelectClipCopyThenMinimize(0)
        content := A_Clipboard
        if (Trim(content) = "") {
            ShowCenteredOverlay_Utils("❌ Clip empty or not text.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        desktopPath := AiQuickDownload_ResolveDesktopPath()
        if (!desktopPath || !DirExist(desktopPath)) {
            ShowCenteredOverlay_Utils("❌ Desktop folder not found.", 2500, BANNER_ACCENT_ERROR)
            return
        }
        outPath := ClipAngelExport_UniqueDesktopPath(desktopPath)
        WriteUtf8File(outPath, content)
        SplitPath(outPath, &name)
        if (StrLen(name) > 48)
            name := SubStr(name, 1, 45) "..."
        ShowCenteredOverlay_Utils("✅ Saved: " name, 2200, BANNER_ACCENT_SUCCESS)
        try ScriptSoundPlay(A_ScriptDir . "\assets\sounds\copy.wav")
    } catch Error as e {
        ShowCenteredOverlay_Utils("❌ Clip Angel export failed: " . e.Message, 2500, BANNER_ACCENT_ERROR)
    } finally {
        try A_Clipboard := savedClip
        catch {
        }
        try StandardLoadingBar_Hide(0)
        catch {
        }
        ClipAngel_ReleaseAutomationLock()
    }
}

RegisterMacro(ClipAngel_ExportLastClipToDesktop, "📎 ClipAngel last clip → Desktop", "c")