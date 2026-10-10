; =============================================================================
; Utils module: desktop_mtime_index.ahk
; Cached Desktop listing. Restats only when the folder mtime changes.
; =============================================================================

global g_DesktopMtimeIndex := Map()
global g_DesktopMtimeFolderKey := ""
global IMPORT_WATCHER_USE_INDEX := true

DesktopMtime_Resolve(desktop := "") {
    if (desktop != "" && DirExist(desktop))
        return RTrim(desktop, "\")
    try {
        if (IsSet(ImportWatcher_ResolveDesktopPath))
            return ImportWatcher_ResolveDesktopPath()
    }
    return RTrim(A_Desktop, "\")
}

DesktopMtime_Refresh(desktop := "") {
    global g_DesktopMtimeIndex, g_DesktopMtimeFolderKey, IMPORT_WATCHER_USE_INDEX
    desktop := DesktopMtime_Resolve(desktop)
    if (!IMPORT_WATCHER_USE_INDEX || desktop = "" || !DirExist(desktop))
        return g_DesktopMtimeIndex
    folderMtime := ""
    try folderMtime := FileGetTime(desktop, "M")
    catch {
        folderMtime := ""
    }
    key := desktop . "|" . folderMtime
    if (key = g_DesktopMtimeFolderKey && g_DesktopMtimeIndex.Count)
        return g_DesktopMtimeIndex
    idx := Map()
    loop files desktop . "\*", "FD" {
        if (StrLower(A_LoopFileName) = "desktop.ini")
            continue
        try {
            mt := FileGetTime(A_LoopFileFullPath, "M")
            sz := 0
            if (InStr(FileExist(A_LoopFileFullPath), "D") != 1)
                sz := FileGetSize(A_LoopFileFullPath)
            idx[A_LoopFileFullPath] := mt . "|" . sz
        } catch {
        }
    }
    g_DesktopMtimeIndex := idx
    g_DesktopMtimeFolderKey := key
    return idx
}

DesktopMtime_Newest(desktop := "") {
    idx := DesktopMtime_Refresh(desktop)
    newestPath := ""
    newestStamp := ""
    for path, stamp in idx {
        parts := StrSplit(stamp, "|")
        mt := parts.Length ? parts[1] : ""
        if (newestStamp = "" || mt > newestStamp) {
            newestStamp := mt
            newestPath := path
        }
    }
    return newestPath
}
