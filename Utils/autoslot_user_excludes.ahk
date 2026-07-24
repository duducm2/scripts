; =============================================================================
; Utils module: autoslot_user_excludes.ahk
; Persisted user AutoSlot ignore list (exe|title needles). Used by #!+L R/I and
; AutoSlot_IsExcludedExeOrTitle (no place/fill/occupancy — same effect as ClipAngel).
; =============================================================================

global g_AutoSlotUserExcludes := []
global g_AutoSlotUserExcludesReady := false
global g_AutoSlotUserExcludeManageResult := ""

AutoSlot_UserExcludes_IniPath() {
    return A_ScriptDir "\assets\data\autoslot_user_excludes.ini"
}

AutoSlot_UserExcludes_ArrJoin(arr, sep := "`n") {
    out := ""
    for i, v in arr {
        if (i > 1)
            out .= sep
        out .= v
    }
    return out
}

AutoSlot_UserExcludes_Register(&list, &seen, needle) {
    n := Trim(needle)
    if (n = "")
        return
    key := StrLower(n)
    if (seen.Has(key))
        return
    seen[key] := true
    list.Push(n)
}

AutoSlot_UserExcludes_ParseDiskEntries(raw) {
    entries := []
    seen := Map()
    for line in StrSplit(raw, "`n", "`r") {
        line := Trim(line)
        if (line = "" || SubStr(line, 1, 1) = ";" || line = "[Excludes]")
            continue
        if (SubStr(line, 1, 1) = "[")
            continue
        AutoSlot_UserExcludes_Register(&entries, &seen, line)
    }
    return entries
}

AutoSlot_UserExcludes_WriteList(list) {
    path := AutoSlot_UserExcludes_IniPath()
    try DirCreate(A_ScriptDir "\assets\data")
    lines := ["[Excludes]", "; One entry per line: title substring, or exe|title (preferred)"]
    for n in list
        lines.Push(n)
    try FileDelete(path)
    FileAppend(AutoSlot_UserExcludes_ArrJoin(lines, "`n") "`n", path, "UTF-8")
}

AutoSlot_UserExcludes_Init() {
    global g_AutoSlotUserExcludes, g_AutoSlotUserExcludesReady
    list := []
    seen := Map()
    path := AutoSlot_UserExcludes_IniPath()
    if (!FileExist(path)) {
        try AutoSlot_UserExcludes_WriteList(list)
        catch {
        }
        g_AutoSlotUserExcludes := list
        g_AutoSlotUserExcludesReady := true
        return
    }
    raw := ""
    try raw := FileRead(path, "UTF-8")
    catch {
        raw := ""
    }
    for entry in AutoSlot_UserExcludes_ParseDiskEntries(raw)
        AutoSlot_UserExcludes_Register(&list, &seen, entry)
    g_AutoSlotUserExcludes := list
    g_AutoSlotUserExcludesReady := true
}

AutoSlot_UserExcludes_Ensure() {
    global g_AutoSlotUserExcludesReady
    if (!g_AutoSlotUserExcludesReady)
        AutoSlot_UserExcludes_Init()
}

AutoSlot_UserExcludes_FormatNeedle(title, exe := "") {
    title := Trim(title)
    exe := Trim(exe)
    if (title = "" && exe = "")
        return ""
    if (exe != "" && title != "")
        return exe . "|" . title
    return title != "" ? title : exe
}

AutoSlot_UserExcludes_NeedleToTitleExe(needle) {
    needle := Trim(needle)
    if (InStr(needle, "|")) {
        parts := StrSplit(needle, "|", , 2)
        return { title: Trim(parts[2]), exe: Trim(parts[1]) }
    }
    return { title: needle, exe: "" }
}

; Match hwnd against persisted user ignore needles (exe|title or substring).
AutoSlot_UserExcludeMatch(hwnd) {
    global g_AutoSlotUserExcludes
    if (!hwnd)
        return false
    AutoSlot_UserExcludes_Ensure()
    title := ""
    exe := ""
    try title := WinGetTitle(hwnd)
    catch {
        title := ""
    }
    try exe := WinGetProcessName("ahk_id " hwnd)
    catch {
        exe := ""
    }
    if (title = "" && exe = "")
        return false
    t := StrLower(title)
    e := StrLower(exe)
    for needle in g_AutoSlotUserExcludes {
        n := Trim(needle)
        if (n = "")
            continue
        if (InStr(n, "|")) {
            parts := StrSplit(n, "|", , 2)
            exeNeedle := StrLower(Trim(parts[1]))
            titleNeedle := parts.Length > 1 ? StrLower(Trim(parts[2])) : ""
            if (exeNeedle != "" && e != "" && InStr(e, exeNeedle)) {
                if (titleNeedle = "" || (t != "" && InStr(t, titleNeedle)))
                    return true
            }
            continue
        }
        nLower := StrLower(n)
        if (t != "" && InStr(t, nLower))
            return true
        if (e != "" && InStr(e, nLower))
            return true
    }
    return false
}

AutoSlot_UserExcludes_IsNeedlePresent(needle) {
    global g_AutoSlotUserExcludes
    AutoSlot_UserExcludes_Ensure()
    key := StrLower(Trim(needle))
    if (key = "")
        return false
    for n in g_AutoSlotUserExcludes {
        if (StrLower(Trim(n)) = key)
            return true
    }
    return false
}

; Add hwnd to user ignore list. Returns true on new save.
AutoSlot_AddUserExcludeFromHwnd(hwnd) {
    global g_AutoSlotUserExcludes
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    title := ""
    exe := ""
    try title := WinGetTitle(hwnd)
    catch {
        title := ""
    }
    try exe := WinGetProcessName("ahk_id " hwnd)
    catch {
        exe := ""
    }
    needle := AutoSlot_UserExcludes_FormatNeedle(title, exe)
    if (needle = "")
        return false
    AutoSlot_UserExcludes_Ensure()
    if (AutoSlot_UserExcludes_IsNeedlePresent(needle)) {
        ShowCenteredOverlay_Utils("ℹ️ Already in AutoSlot ignore list", 2000, BANNER_ACCENT_INFO)
        return false
    }
    list := []
    seen := Map()
    for n in g_AutoSlotUserExcludes
        AutoSlot_UserExcludes_Register(&list, &seen, n)
    AutoSlot_UserExcludes_Register(&list, &seen, needle)
    try {
        AutoSlot_UserExcludes_WriteList(list)
    } catch as err {
        ShowCenteredOverlay_Utils("❌ Could not save ignore list: " . err.Message, 4000, BANNER_ACCENT_ERROR)
        return false
    }
    AutoSlot_UserExcludes_Init()
    label := needle
    if (StrLen(label) > 50)
        label := SubStr(label, 1, 47) . "..."
    ShowCenteredOverlay_Utils("✅ Ignored for rearrange: " . label, 2200, BANNER_ACCENT_SUCCESS)
    return true
}

; Remove by 1-based index. Returns true if removed.
AutoSlot_RemoveUserExcludeAt(index) {
    global g_AutoSlotUserExcludes
    AutoSlot_UserExcludes_Ensure()
    index := Integer(index)
    if (index < 1 || index > g_AutoSlotUserExcludes.Length)
        return false
    list := []
    seen := Map()
    for i, n in g_AutoSlotUserExcludes {
        if (i = index)
            continue
        AutoSlot_UserExcludes_Register(&list, &seen, n)
    }
    try {
        AutoSlot_UserExcludes_WriteList(list)
    } catch as err {
        ShowCenteredOverlay_Utils("❌ Could not save ignore list: " . err.Message, 4000, BANNER_ACCENT_ERROR)
        return false
    }
    AutoSlot_UserExcludes_Init()
    return true
}

AutoSlot_RemoveUserExcludeByNeedle(needle) {
    global g_AutoSlotUserExcludes
    needle := Trim(needle)
    if (needle = "")
        return false
    AutoSlot_UserExcludes_Ensure()
    key := StrLower(needle)
    idx := 0
    for i, n in g_AutoSlotUserExcludes {
        if (StrLower(Trim(n)) = key) {
            idx := i
            break
        }
    }
    if (!idx)
        return false
    return AutoSlot_RemoveUserExcludeAt(idx)
}

AutoSlot_UserExcludeManage_DigitSequence() {
    return ["1", "2", "3", "4", "5", "6", "7", "8", "9", "0"]
}

AutoSlot_UserExcludeManage_BuildState(list) {
    lines := ["AutoSlot ignore list — digit removes entry", ""]
    seq := AutoSlot_UserExcludeManage_DigitSequence()
    limit := Min(list.Length, seq.Length)
    loop limit {
        needle := list[A_Index]
        if (StrLen(needle) > 56)
            needle := SubStr(needle, 1, 53) . "..."
        lines.Push("[" . seq[A_Index] . "] " . needle)
    }
    if (list.Length > seq.Length)
        lines.Push("(" . (list.Length - seq.Length) . " more — remove some to see)")
    lines.Push("")
    lines.Push("[Esc] Done")
    return AutoSlot_UserExcludes_ArrJoin(lines, "`n")
}

AutoSlot_UserExcludeManage_OnRemove(index, *) {
    global g_AutoSlotUserExcludeManageResult
    g_AutoSlotUserExcludeManageResult := index
    try StandardLoadingBar_CloseKeysOverlay()
    catch {
    }
    try StandardLoadingBar_Hide(0)
    catch {
    }
}

AutoSlot_UserExcludeManage_OnDone(*) {
    global g_AutoSlotUserExcludeManageResult
    g_AutoSlotUserExcludeManageResult := "done"
    try StandardLoadingBar_CloseKeysOverlay()
    catch {
    }
    try StandardLoadingBar_Hide(0)
    catch {
    }
}

; Blocking manage UI: digit removes, Esc done; refreshes after each delete.
AutoSlot_ShowUserExcludeManageUI() {
    global g_AutoSlotUserExcludes, g_AutoSlotUserExcludeManageResult
    loop {
        AutoSlot_UserExcludes_Ensure()
        if (g_AutoSlotUserExcludes.Length = 0) {
            ShowCenteredOverlay_Utils("ℹ️ AutoSlot ignore list empty", 2200, BANNER_ACCENT_INFO)
            return
        }
        list := []
        for n in g_AutoSlotUserExcludes
            list.Push(n)
        state := AutoSlot_UserExcludeManage_BuildState(list)
        seq := AutoSlot_UserExcludeManage_DigitSequence()
        keyCallbacks := Map()
        limit := Min(list.Length, seq.Length)
        loop limit {
            idx := A_Index
            keyCallbacks[seq[idx]] := AutoSlot_UserExcludeManage_OnRemove.Bind(idx)
        }
        keyCallbacks["Escape"] := AutoSlot_UserExcludeManage_OnDone
        g_AutoSlotUserExcludeManageResult := ""
        StandardLoadingBar_ShowWithKeys(
            state,
            keyCallbacks,
            0,
            0,
            "",
            BANNER_ACCENT_INFO,
            620,
            16,
            "",
            false,
            "[1]… remove  [Esc] Done",
            true,
            false,
            true)
        start := A_TickCount
        while (g_AutoSlotUserExcludeManageResult = "") {
            if ((A_TickCount - start) >= 60000) {
                g_AutoSlotUserExcludeManageResult := "done"
                break
            }
            Sleep 50
        }
        result := g_AutoSlotUserExcludeManageResult
        g_AutoSlotUserExcludeManageResult := ""
        try StandardLoadingBar_CloseKeysOverlay()
        catch {
        }
        try StandardLoadingBar_Hide(0)
        catch {
        }
        if (result = "done" || result = "")
            return
        idx := Integer(result)
        if (idx >= 1 && idx <= list.Length) {
            removed := list[idx]
            if (AutoSlot_RemoveUserExcludeAt(idx)) {
                label := removed
                if (StrLen(label) > 50)
                    label := SubStr(label, 1, 47) . "..."
                ShowCenteredOverlay_Utils("✅ Removed: " . label, 1800, BANNER_ACCENT_SUCCESS)
                Sleep 400
            }
        }
    }
}
