; =============================================================================
; Utils module: desktop_cut_newest.ahk
; Cut, open (+ activate with quality gate), copy-path newest Desktop item, or paste clipboard file to Desktop.
; Trigger: Win+Alt+Shift+O (same tiering as #!+8 pronunciation + 3×):
;   1× = cut newest Desktop item, restore previous window
;   2× within 400 ms (AI_QD_DOUBLE_TAP_MS / ZMK tap-dance) = open with default app
;   3× within 400 ms windows = paste clipboard CF_HDROP copy to Desktop, or (Cursor/Code)
;       copy the active editor tab file to Desktop after sidebar/Explorer gates
;   hold 700 ms+ (PRONUNCIATION_HOLD_MS / Fast Copy / cheat sheet) = copy path as text
; =============================================================================

DESKTOP_CUT_NEWEST_HOLD_MS := 700

; Returns full path of newest item under desktopPath, or "" if none.
; Newest = later of Creation vs Modified; skips desktop.ini; includes files and folders.
DesktopCutNewest_ResolveNewestPath(desktopPath) {
    if (!desktopPath || !DirExist(desktopPath))
        return ""
    newestPath := ""
    newestStamp := ""
    loop files desktopPath "\*", "FD" {
        if (StrLower(A_LoopFileName) = "desktop.ini")
            continue
        try {
            tC := FileGetTime(A_LoopFileFullPath, "C")
            tM := FileGetTime(A_LoopFileFullPath, "M")
        } catch {
            continue
        }
        stamp := (tC >= tM) ? tC : tM
        if (newestStamp = "" || stamp > newestStamp) {
            newestStamp := stamp
            newestPath := A_LoopFileFullPath
        }
    }
    return newestPath
}

DesktopCutNewest_ResolveDesktopPath() {
    desktopPath := ""
    try desktopPath := GetDesktopToRecyclePath()
    catch
        desktopPath := A_Desktop
    if (!desktopPath || !DirExist(desktopPath))
        desktopPath := A_Desktop
    return DirExist(desktopPath) ? desktopPath : ""
}

DesktopCutNewest_Trigger() {
    desktopPath := DesktopCutNewest_ResolveDesktopPath()
    if (desktopPath = "") {
        ShowCenteredOverlay_Utils("❌ Desktop folder not found", 2500, BANNER_ACCENT_ERROR)
        return
    }

    newest := DesktopCutNewest_ResolveNewestPath(desktopPath)
    if (newest = "") {
        ShowCenteredOverlay_Utils("⚠ Desktop is empty", 2000, BANNER_ACCENT_INTERMEDIATE)
        return
    }
    DesktopCutNewest_CutPath(newest)
}

; Cut a specific Desktop file/folder path to the clipboard (CF_HDROP move).
DesktopCutNewest_CutPath(path) {
    origHwnd := WinExist("A")
    if (!path || !FileExist(path)) {
        ShowCenteredOverlay_Utils("❌ Desktop item not found", 2500, BANNER_ACCENT_ERROR)
        return false
    }

    if !Clipboard_CutFiles([path]) {
        ShowCenteredOverlay_Utils("❌ Failed to cut Desktop item", 2500, BANNER_ACCENT_ERROR)
        return false
    }
    if !Clipboard_ContainsFilePath(path) {
        ShowCenteredOverlay_Utils("❌ Cut verify failed", 2500, BANNER_ACCENT_ERROR)
        return false
    }

    SplitPath(path, &name)
    ShowCenteredOverlay_Utils("✂️ Cut: " name, 1800, BANNER_ACCENT_SUCCESS)

    if (origHwnd && WinExist("ahk_id " origHwnd)) {
        try WinActivate("ahk_id " origHwnd)
        WinWaitActive("ahk_id " origHwnd, , 1)
    }
    return true
}

DesktopCutNewest_OpenNewest() {
    desktopPath := DesktopCutNewest_ResolveDesktopPath()
    if (desktopPath = "") {
        ShowCenteredOverlay_Utils("❌ Desktop folder not found", 2500, BANNER_ACCENT_ERROR)
        return
    }

    newest := DesktopCutNewest_ResolveNewestPath(desktopPath)
    if (newest = "") {
        ShowCenteredOverlay_Utils("⚠ Desktop is empty", 2000, BANNER_ACCENT_INTERMEDIATE)
        return
    }

    SplitPath(newest, &name)
    browser := DesktopCutNewest_ResolveBrowserLaunch(newest)
    beforeMap := DesktopCutNewest_SnapshotWindowMap(browser ? browser.exeName : "")
    fgBefore := WinExist("A")

    ; Let the launched app take foreground (Windows focus-stealing guard).
    try DllCall("AllowSetForegroundWindow", "UInt", 0xFFFFFFFF)  ; ASFW_ANY
    catch {
    }

    try {
        if (browser)
            DesktopCutNewest_OpenInBrowserNewWindow(newest, browser)
        else
            Run('"' . newest . '"')
    } catch {
        ShowCenteredOverlay_Utils("❌ Failed to open Desktop item", 2500, BANNER_ACCENT_ERROR)
        return
    }

    newHwnd := DesktopCutNewest_WaitForOpenedWindow(beforeMap, fgBefore, browser ? 8000 : 5000, browser,
        name)
    if (!newHwnd)
        newHwnd := DesktopCutNewest_FindWindowByTitleHint(name)

    if (!newHwnd) {
        ShowCenteredOverlay_Utils("❌ Opened but window not found", 2500, BANNER_ACCENT_ERROR)
        return
    }

    ; Quality gate: must be foreground before success. Retry + re-resolve by title if needed.
    if !DesktopCutNewest_EnsureActivated(newHwnd, 3000) {
        alt := DesktopCutNewest_FindWindowByTitleHint(name)
        if (alt)
            newHwnd := alt
        if !DesktopCutNewest_EnsureActivated(newHwnd, 2000) {
            ShowCenteredOverlay_Utils("❌ Could not activate opened window", 2500, BANNER_ACCENT_ERROR)
            return
        }
    }

    ShowCenteredOverlay_Utils("📂 Open: " name, 1800, BANNER_ACCENT_SUCCESS)

    ; Overlay / AutoSlot can steal focus — re-activate after the banner and re-check the gate.
    DesktopCutNewest_ScheduleActivate(newHwnd, 1900, name)
}

DesktopCutNewest_SnapshotWindowMap(exeName := "") {
    snap := Map()
    try {
        listSpec := exeName ? ("ahk_exe " exeName) : ""
        for hwnd in WinGetList(listSpec)
            snap[hwnd] := true
    } catch {
    }
    return snap
}

; First new visible top-level window after open, or FG if it changed (reuse case).
; When browser is set, only new windows for that exe count (avoids tab-in-existing-window).
DesktopCutNewest_WaitForOpenedWindow(beforeMap, fgBefore, timeoutMs := 5000, browser := "", titleHint := "") {
    if (!IsObject(beforeMap))
        beforeMap := Map()
    listSpec := (IsObject(browser) && browser.exeName) ? ("ahk_exe " browser.exeName) : ""
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        try {
            best := 0
            for hwnd in WinGetList(listSpec) {
                if beforeMap.Has(hwnd)
                    continue
                if !DesktopCutNewest_IsCandidateOpenWindow(hwnd)
                    continue
                if (titleHint != "" && DesktopCutNewest_TitleMatchesHint(hwnd, titleHint))
                    return hwnd
                if (!best)
                    best := hwnd
            }
            if (best)
                return best
            fg := WinExist("A")
            if (fg && fg != fgBefore) {
                if (listSpec = "" || DesktopCutNewest_HwndMatchesExe(fg, browser ? browser.exeName : ""))
                    return fg
            }
        } catch {
        }
        Sleep 50
    }
    fg := WinExist("A")
    if (fg && fg != fgBefore) {
        if (listSpec = "" || DesktopCutNewest_HwndMatchesExe(fg, browser ? browser.exeName : ""))
            return fg
    }
    return 0
}

DesktopCutNewest_IsCandidateOpenWindow(hwnd) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    try {
        if !DllCall("IsWindowVisible", "Ptr", hwnd)
            return false
        if (DllCall("GetWindow", "Ptr", hwnd, "UInt", 4))  ; GW_OWNER
            return false
    } catch {
        return false
    }
    return true
}

DesktopCutNewest_HwndMatchesExe(hwnd, exeName) {
    if (exeName = "")
        return true
    try return (StrLower(WinGetProcessName("ahk_id " hwnd)) = StrLower(exeName))
    catch
        return false
}

DesktopCutNewest_TitleMatchesHint(hwnd, titleHint) {
    if (titleHint = "")
        return false
    try {
        title := WinGetTitle("ahk_id " hwnd)
        if (title = "")
            return false
        hint := StrLower(titleHint)
        titleLower := StrLower(title)
        if (InStr(titleLower, hint))
            return true
        SplitPath(titleHint, &hintName, , &hintExt)
        if (hintExt != "" && InStr(titleLower, StrLower(hintName)))
            return true
    } catch {
    }
    return false
}

DesktopCutNewest_FindWindowByTitleHint(titleHint) {
    if (titleHint = "")
        return 0
    try {
        for hwnd in WinGetList() {
            if !DesktopCutNewest_IsCandidateOpenWindow(hwnd)
                continue
            if DesktopCutNewest_TitleMatchesHint(hwnd, titleHint)
                return hwnd
        }
    } catch {
    }
    return 0
}

; Quality gate: window exists, is a visible top-level candidate, not minimized, and is foreground.
DesktopCutNewest_IsForegroundOk(hwnd) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    if !DesktopCutNewest_IsCandidateOpenWindow(hwnd)
        return false
    try {
        if (WinGetMinMax("ahk_id " hwnd) = -1)
            return false
        if DllCall("IsIconic", "Ptr", hwnd)
            return false
    } catch {
        return false
    }
    return !!WinActive("ahk_id " hwnd)
}

; Activate until quality gate passes (or timeout). Returns true only when foreground-ok.
DesktopCutNewest_EnsureActivated(hwnd, timeoutMs := 2500) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    if DesktopCutNewest_IsForegroundOk(hwnd)
        return true
    deadline := A_TickCount + Max(200, timeoutMs)
    while (A_TickCount < deadline) {
        DesktopCutNewest_ActivateHwnd(hwnd)
        if DesktopCutNewest_IsForegroundOk(hwnd)
            return true
        remainingSec := (deadline - A_TickCount) / 1000.0
        if (remainingSec > 0.05) {
            try WinWaitActive("ahk_id " hwnd, , Min(remainingSec, 0.5))
            catch {
            }
        }
        if DesktopCutNewest_IsForegroundOk(hwnd)
            return true
        Sleep 60
    }
    return DesktopCutNewest_IsForegroundOk(hwnd)
}

DesktopCutNewest_ScheduleActivate(hwnd, delayMs := 1500, titleHint := "") {
    global g_DesktopCutNewest_ScheduleActivateHwnd, g_DesktopCutNewest_ScheduleActivateTimer
    global g_DesktopCutNewest_ScheduleActivateTitleHint
    if (!hwnd || delayMs < 1)
        return
    g_DesktopCutNewest_ScheduleActivateHwnd := hwnd
    g_DesktopCutNewest_ScheduleActivateTitleHint := titleHint
    if (!g_DesktopCutNewest_ScheduleActivateTimer)
        g_DesktopCutNewest_ScheduleActivateTimer := ObjBindMethod(DesktopCutNewest_ScheduleActivateObj, "OnTimer")
    SetTimer(g_DesktopCutNewest_ScheduleActivateTimer, -delayMs)
}

class DesktopCutNewest_ScheduleActivateObj {
    static OnTimer() {
        global g_DesktopCutNewest_ScheduleActivateHwnd, g_DesktopCutNewest_ScheduleActivateTitleHint
        hwnd := g_DesktopCutNewest_ScheduleActivateHwnd
        titleHint := g_DesktopCutNewest_ScheduleActivateTitleHint
        g_DesktopCutNewest_ScheduleActivateHwnd := 0
        g_DesktopCutNewest_ScheduleActivateTitleHint := ""

        if (!hwnd || !WinExist("ahk_id " hwnd)) {
            if (titleHint != "")
                hwnd := DesktopCutNewest_FindWindowByTitleHint(titleHint)
        }
        if (!hwnd) {
            ShowCenteredOverlay_Utils("❌ Could not activate opened window", 2500, BANNER_ACCENT_ERROR)
            return
        }
        if DesktopCutNewest_IsForegroundOk(hwnd)
            return
        if !DesktopCutNewest_EnsureActivated(hwnd, 2000) {
            if (titleHint != "") {
                alt := DesktopCutNewest_FindWindowByTitleHint(titleHint)
                if (alt && alt != hwnd)
                    hwnd := alt
            }
            if !DesktopCutNewest_EnsureActivated(hwnd, 1500) {
                ShowCenteredOverlay_Utils("❌ Could not activate opened window", 2500, BANNER_ACCENT_ERROR)
                return
            }
        }
    }
}

DesktopCutNewest_ActivateHwnd(hwnd) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    try {
        try {
            pid := WinGetPID("ahk_id " hwnd)
            if (pid)
                DllCall("AllowSetForegroundWindow", "UInt", pid)
        } catch {
        }

        ; AttachThreadInput helps when another process holds foreground lock.
        attached := false
        targetTid := 0
        fgTid := 0
        try {
            targetTid := DllCall("GetWindowThreadProcessId", "Ptr", hwnd, "UInt*", 0, "UInt")
            fg := DllCall("GetForegroundWindow", "Ptr")
            if (fg)
                fgTid := DllCall("GetWindowThreadProcessId", "Ptr", fg, "UInt*", 0, "UInt")
            curTid := DllCall("GetCurrentThreadId", "UInt")
            if (fgTid && fgTid != curTid)
                attached := !!DllCall("AttachThreadInput", "UInt", curTid, "UInt", fgTid, "Int", 1)
            if (targetTid && targetTid != curTid && targetTid != fgTid)
                DllCall("AttachThreadInput", "UInt", curTid, "UInt", targetTid, "Int", 1)
        } catch {
        }

        try {
            if (WinGetMinMax("ahk_id " hwnd) = -1 || DllCall("IsIconic", "Ptr", hwnd))
                WinRestore("ahk_id " hwnd)
            try WinShow("ahk_id " hwnd)
            catch {
            }
            WinActivate("ahk_id " hwnd)
            if WinWaitActive("ahk_id " hwnd, , 0.8)
                return true
            DllCall("SwitchToThisWindow", "Ptr", hwnd, "Int", 1)
            DllCall("SetForegroundWindow", "Ptr", hwnd)
            DllCall("BringWindowToTop", "Ptr", hwnd)
            WinActivate("ahk_id " hwnd)
            if WinWaitActive("ahk_id " hwnd, , 0.5)
                return true
            ; Last-resort nudge used elsewhere when focus lock blocks WinActivate.
            WinSetAlwaysOnTop("On", "ahk_id " hwnd)
            Sleep 40
            WinSetAlwaysOnTop("Off", "ahk_id " hwnd)
            WinActivate("ahk_id " hwnd)
            return !!WinActive("ahk_id " hwnd)
        } finally {
            try {
                curTid := DllCall("GetCurrentThreadId", "UInt")
                if (targetTid && targetTid != curTid)
                    DllCall("AttachThreadInput", "UInt", curTid, "UInt", targetTid, "Int", 0)
                if (attached && fgTid)
                    DllCall("AttachThreadInput", "UInt", curTid, "UInt", fgTid, "Int", 0)
            } catch {
            }
        }
    } catch {
        return false
    }
}

; When Windows opens this path with a browser, return {exe, exeName, flag} for a new window.
DesktopCutNewest_ResolveBrowserLaunch(path) {
    if (!path || !FileExist(path))
        return ""
    SplitPath(path, , , &ext)
    if (ext = "")
        return ""
    extDot := "." . StrLower(ext)
    browser := DesktopCutNewest_ParseBrowserFromAssocCommand(DesktopCutNewest_GetAssocOpenCommand(extDot))
    if (browser)
        return browser
    ; Web formats only — daily driver is chrome.exe elsewhere in this repo.
    static webExts := Map("html", 1, "htm", 1, "svg", 1, "mhtml", 1, "xhtml", 1)
    if (webExts.Has(StrLower(ext)))
        return { exe: "chrome.exe", exeName: "chrome.exe", flag: "--new-window" }
    return ""
}

DesktopCutNewest_ParseBrowserFromAssocCommand(cmd) {
    if (cmd = "")
        return ""
    exe := ""
    if RegExMatch(cmd, '"(?P<exe>[^"]+\.exe)"', &m)
        exe := m.exe
    else if RegExMatch(cmd, '(?i)([A-Z]:\\[^\s"]+\.exe)', &m)
        exe := m[1]
    if (exe = "")
        return ""
    exeName := StrLower(RegExReplace(exe, ".*\\", ""))
    static browserFlags := Map(
        "chrome.exe", "--new-window",
        "msedge.exe", "--new-window",
        "brave.exe", "--new-window",
        "chromium.exe", "--new-window",
        "vivaldi.exe", "--new-window",
        "opera.exe", "--new-window",
        "firefox.exe", "-new-window"
    )
    if !browserFlags.Has(exeName)
        return ""
    return { exe: exe, exeName: exeName, flag: browserFlags[exeName] }
}

; ASSOCSTR_COMMAND for ".ext" via AssocQueryStringW; "" on failure.
DesktopCutNewest_GetAssocOpenCommand(extWithDot) {
    if (!extWithDot)
        return ""
    ; ASSOCF_NONE = 0, ASSOCSTR_COMMAND = 1
    pcch := 0
    hr := DllCall("shlwapi\AssocQueryStringW", "UInt", 0, "UInt", 1, "WStr", extWithDot, "Ptr", 0, "Ptr", 0,
        "UInt*", &pcch, "UInt")
    if (pcch < 2)
        return ""
    buf := Buffer(pcch * 2, 0)
    hr := DllCall("shlwapi\AssocQueryStringW", "UInt", 0, "UInt", 1, "WStr", extWithDot, "Ptr", 0, "Ptr", buf,
        "UInt*", &pcch, "UInt")
    if (hr != 0)
        return ""
    return StrGet(buf, "UTF-16")
}

DesktopCutNewest_PathToFileUrl(path) {
    p := StrReplace(path, "\", "/")
    if (RegExMatch(p, "i)^[a-z]:"))
        p := "/" . p
    p := StrReplace(p, " ", "%20")
    return "file://" . p
}

DesktopCutNewest_OpenInBrowserNewWindow(path, browser) {
    if (!IsObject(browser) || browser.exe = "")
        throw Error("No browser launch info")
    fileUrl := DesktopCutNewest_PathToFileUrl(path)
    Run('"' . browser.exe . '" ' . browser.flag . ' "' . fileUrl . '"')
}

DesktopCutNewest_CopyPath() {
    desktopPath := DesktopCutNewest_ResolveDesktopPath()
    if (desktopPath = "") {
        ShowCenteredOverlay_Utils("❌ Desktop folder not found", 2500, BANNER_ACCENT_ERROR)
        return
    }

    newest := DesktopCutNewest_ResolveNewestPath(desktopPath)
    if (newest = "") {
        ShowCenteredOverlay_Utils("⚠ Desktop is empty", 2000, BANNER_ACCENT_INTERMEDIATE)
        return
    }

    try {
        A_Clipboard := newest
    } catch {
        ShowCenteredOverlay_Utils("❌ Failed to copy path", 2500, BANNER_ACCENT_ERROR)
        return
    }
    if !ClipWait(1) {
        ShowCenteredOverlay_Utils("❌ Clipboard did not update", 2500, BANNER_ACCENT_ERROR)
        return
    }

    SplitPath(newest, &name)
    ShowCenteredOverlay_Utils("📋 Path: " name, 1800, BANNER_ACCENT_SUCCESS)
}

; Unique Desktop dest for srcPath (name.ext → name (1).ext → …).
DesktopCutNewest_UniqueDestPath(desktopPath, srcPath) {
    SplitPath(srcPath, &name, , &ext, &nameNoExt)
    if (name = "")
        return ""
    dest := desktopPath "\" name
    if !FileExist(dest)
        return dest
    isDir := !!DirExist(srcPath)
    i := 1
    loop {
        if (!isDir && ext != "")
            candidate := desktopPath "\" nameNoExt " (" i ")." ext
        else
            candidate := desktopPath "\" name " (" i ")"
        if !FileExist(candidate)
            return candidate
        i += 1
        if (i > 9999)
            return ""
    }
}

DesktopCutNewest_IsEditorActive() {
    return !!(WinActive("ahk_exe Cursor.exe") || WinActive("ahk_exe Code.exe"))
}

DesktopCutNewest_NormalizeRevealBasename(raw) {
    if (raw = "")
        return ""
    s := Trim(raw)
    ; Dirty / preview markers in editor titles (Cursor/VS Code).
    s := RegExReplace(s, "^[\x{25CF}\x{25A0}*•]+\s*", "")
    s := Trim(s)
    if InStr(s, ",")
        s := Trim(SubStr(s, 1, InStr(s, ",") - 1))
    if (InStr(s, "\") || InStr(s, "/")) {
        SplitPath(s, &name)
        if (name != "")
            s := name
    }
    return s
}

DesktopCutNewest_GetBasenameFromEditorTitle(editorHwnd) {
    if !(editorHwnd is Integer) || editorHwnd <= 0
        return ""
    try {
        title := WinGetTitle("ahk_id " editorHwnd)
        if (title = "")
            return ""
        parts := StrSplit(title, " - ", , 2)
        if (parts.Length >= 1 && parts[1] != "") {
            candidate := DesktopCutNewest_NormalizeRevealBasename(Trim(parts[1]))
            if (candidate != "" && StrLen(candidate) <= 180 && !InStr(candidate, "`n"))
                return candidate
        }
    } catch {
    }
    return ""
}

DesktopCutNewest_FindWorkbenchToggleButton(root, nameSubstring) {
    if !root || !nameSubstring
        return 0
    toggleBtn := 0
    try toggleBtn := root.FindFirst({ Name: nameSubstring, Type: UIA.Type.Button })
    catch
        toggleBtn := 0
    if toggleBtn
        return toggleBtn
    try {
        allButtons := root.FindAll({ Type: UIA.Type.Button })
        if allButtons {
            for btn in allButtons {
                try {
                    if InStr(btn.Name, nameSubstring) {
                        toggleBtn := btn
                        break
                    }
                } catch {
                }
            }
        }
    } catch {
    }
    return toggleBtn
}

DesktopCutNewest_IsWorkbenchToggleOn(root, nameSubstring) {
    toggleBtn := DesktopCutNewest_FindWorkbenchToggleButton(root, nameSubstring)
    if !toggleBtn
        return false
    try {
        if InStr(toggleBtn.ClassName, "checked")
            return true
    } catch {
    }
    try {
        if toggleBtn.GetPropertyValue(UIA.Property.IsTogglePatternAvailable) {
            toggleState := toggleBtn.TogglePattern.ToggleState
            return (toggleState = 1)
        }
    } catch {
    }
    return false
}

DesktopCutNewest_IsPrimarySidebarVisible(editorHwnd := 0) {
    try {
        if !editorHwnd
            editorHwnd := WinExist("A")
        if !editorHwnd
            return false
        root := UIA.ElementFromHandle(editorHwnd)
        if !root
            return false
        for el in root.FindAll({ Type: UIA.Type.Pane }) {
            try {
                cls := el.ClassName
                if InStr(cls, "monaco-workbench") && InStr(cls, "sidebarvisible")
                    return true
            } catch {
            }
        }
        if DesktopCutNewest_IsWorkbenchToggleOn(root, "Toggle Primary Side Bar")
            return true
    } catch {
    }
    return false
}

DesktopCutNewest_FindFilesExplorerTree(root) {
    if !root
        return 0
    treeType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Tree)
    feEn := UIA.CreatePropertyCondition(UIA.Property.Name, "Files Explorer")
    fePt := UIA.CreatePropertyCondition(UIA.Property.Name, "Explorador de Arquivos")
    feName := UIA.CreateOrCondition(feEn, fePt)
    feCond := UIA.CreateAndCondition(treeType, feName)
    fileTree := 0
    try fileTree := root.FindElement(feCond, UIA.TreeScope.Descendants)
    if fileTree
        return fileTree
    for autoId in ["FileExplorer3", "FileExplorer2", "FileExplorer"] {
        try {
            cond := UIA.CreatePropertyCondition(UIA.Property.AutomationId, autoId)
            fileTree := root.FindElement(cond, UIA.TreeScope.Descendants)
            if fileTree
                return fileTree
        }
    }
    return 0
}

DesktopCutNewest_UiaElementHasAncestor(el, ancestor) {
    if !el || !ancestor
        return false
    current := el
    loop 40 {
        if !current
            break
        try {
            if UIA.CompareElements(ancestor, current)
                return true
        } catch {
        }
        try current := UIA.TreeWalkerTrue.GetParentElement(current)
        catch
            break
    }
    return false
}

DesktopCutNewest_FocusIsInFilesExplorer(editorHwnd := 0) {
    try {
        if !DesktopCutNewest_IsEditorActive()
            return false
        if !editorHwnd
            editorHwnd := WinExist("A")
        if !editorHwnd
            return false
        root := UIA.ElementFromHandle(editorHwnd)
        if !root
            return false
        fileTree := DesktopCutNewest_FindFilesExplorerTree(root)
        if !fileTree
            return false
        fe := UIA.GetFocusedElement()
        if !fe
            return false
        return DesktopCutNewest_UiaElementHasAncestor(fe, fileTree)
    } catch {
        return false
    }
}

DesktopCutNewest_TryFocusFilesExplorerTree(editorHwnd) {
    try {
        if !editorHwnd
            return false
        root := UIA.ElementFromHandle(editorHwnd)
        if !root
            return false
        fileTree := DesktopCutNewest_FindFilesExplorerTree(root)
        if !fileTree
            return false
        fileTree.SetFocus()
        return true
    } catch {
        return false
    }
}

DesktopCutNewest_WaitForSidebarExplorerFocus(editorHwnd, timeoutMs := 800) {
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        if !DesktopCutNewest_IsPrimarySidebarVisible(editorHwnd) {
            Sleep 50
            continue
        }
        if DesktopCutNewest_TryFocusFilesExplorerTree(editorHwnd)
            return true
        Sleep 50
    }
    return DesktopCutNewest_FocusIsInFilesExplorer(editorHwnd)
}

DesktopCutNewest_EnsureFilesExplorerSidebarFocused(editorHwnd) {
    if DesktopCutNewest_FocusIsInFilesExplorer(editorHwnd)
        return true
    if !editorHwnd
        return false
    try WinActivate("ahk_id " editorHwnd)
    catch {
    }
    Send "^+e"
    if DesktopCutNewest_WaitForSidebarExplorerFocus(editorHwnd, 800)
        return true
    Send "^!+e"
    return DesktopCutNewest_WaitForSidebarExplorerFocus(editorHwnd, 400)
}

DesktopCutNewest_HidePrimarySidebar(editorHwnd) {
    if !editorHwnd || !DesktopCutNewest_IsPrimarySidebarVisible(editorHwnd)
        return true
    try WinActivate("ahk_id " editorHwnd)
    catch {
    }
    Send "^b"
    deadline := A_TickCount + 600
    while (A_TickCount < deadline) {
        Sleep 50
        if !DesktopCutNewest_IsPrimarySidebarVisible(editorHwnd)
            return true
    }
    return !DesktopCutNewest_IsPrimarySidebarVisible(editorHwnd)
}

DesktopCutNewest_ReturnFocusToMainEditor(editorHwnd) {
    if !editorHwnd
        return false
    try WinActivate("ahk_id " editorHwnd)
    catch {
    }
    Send "^+e"
    Sleep 80
    Send "^+e"
    return true
}

; Probe active-editor path via Files Explorer focus + Ctrl+2 (user copyFilePath bindings).
; Caller restores clipboard after reading the path (success or failure).
DesktopCutNewest_ResolveActiveEditorFilePath(editorHwnd) {
    if !editorHwnd
        return ""
    expectedBasename := DesktopCutNewest_GetBasenameFromEditorTitle(editorHwnd)
    A_Clipboard := ""
    SendInput "^2"
    clipOk := ClipWait(0.6)
    if !clipOk {
        Sleep 80
        A_Clipboard := ""
        SendInput "^2"
        clipOk := ClipWait(0.6)
    }
    if !clipOk
        return ""
    pathText := Trim(Trim(A_Clipboard), Chr(34))
    pathText := StrReplace(pathText, "/", "\")
    if (pathText = "" || !Clipboard_PathIsExistingFile(pathText))
        return ""
    if (expectedBasename != "") {
        SplitPath(pathText, &pathName)
        if (StrLower(pathName) != StrLower(expectedBasename))
            return ""
    }
    return pathText
}

; Cursor/Code ×3: copy active editor tab file to Desktop (leave original).
DesktopCutNewest_CopyActiveEditorFileToDesktop(desktopPath) {
    editorHwnd := WinExist("A")
    if !editorHwnd || !DesktopCutNewest_IsEditorActive() {
        ShowCenteredOverlay_Utils("❌ No file in clipboard", 2500, BANNER_ACCENT_ERROR)
        return
    }

    savedClip := 0
    try savedClip := ClipboardAll()
    catch {
        savedClip := 0
    }

    sidebarWasVisible := DesktopCutNewest_IsPrimarySidebarVisible(editorHwnd)
    pathText := ""
    try {
        if !DesktopCutNewest_EnsureFilesExplorerSidebarFocused(editorHwnd) {
            ShowCenteredOverlay_Utils("❌ Files Explorer not ready", 2500, BANNER_ACCENT_ERROR)
            return
        }
        pathText := DesktopCutNewest_ResolveActiveEditorFilePath(editorHwnd)
        if (pathText = "") {
            ShowCenteredOverlay_Utils("❌ Could not resolve active file", 2500, BANNER_ACCENT_ERROR)
            return
        }
    } finally {
        if IsObject(savedClip) {
            try A_Clipboard := savedClip
            catch {
            }
        }
        DesktopCutNewest_ReturnFocusToMainEditor(editorHwnd)
        if !sidebarWasVisible
            DesktopCutNewest_HidePrimarySidebar(editorHwnd)
    }

    if (pathText = "")
        return

    dest := DesktopCutNewest_UniqueDestPath(desktopPath, pathText)
    if (dest = "") {
        ShowCenteredOverlay_Utils("❌ Could not build Desktop path", 2500, BANNER_ACCENT_ERROR)
        return
    }
    try FileCopy(pathText, dest)
    catch {
        ShowCenteredOverlay_Utils("❌ Failed to paste to Desktop", 2500, BANNER_ACCENT_ERROR)
        return
    }
    if !FileExist(dest) {
        ShowCenteredOverlay_Utils("❌ Paste verify failed", 2500, BANNER_ACCENT_ERROR)
        return
    }
    SplitPath(dest, &destName)
    ShowCenteredOverlay_Utils("📎 Pasted: " destName, 1800, BANNER_ACCENT_SUCCESS)
}

; 3×: copy clipboard CF_HDROP file(s)/folder(s) onto Desktop (leave originals),
; or when Cursor/Code is focused with no file drop, copy the active editor file.
DesktopCutNewest_PasteClipboardToDesktop() {
    desktopPath := DesktopCutNewest_ResolveDesktopPath()
    if (desktopPath = "") {
        ShowCenteredOverlay_Utils("❌ Desktop folder not found", 2500, BANNER_ACCENT_ERROR)
        return
    }

    if !Clipboard_HasFileDrop() {
        if DesktopCutNewest_IsEditorActive() {
            DesktopCutNewest_CopyActiveEditorFileToDesktop(desktopPath)
            return
        }
        ShowCenteredOverlay_Utils("❌ No file in clipboard", 2500, BANNER_ACCENT_ERROR)
        return
    }

    paths := Clipboard_GetFilePaths()
    if (!paths || paths.Length < 1) {
        ShowCenteredOverlay_Utils("❌ No file in clipboard", 2500, BANNER_ACCENT_ERROR)
        return
    }

    copiedNames := []
    for src in paths {
        if (src = "" || !FileExist(src)) {
            ShowCenteredOverlay_Utils("❌ Clipboard file not found", 2500, BANNER_ACCENT_ERROR)
            return
        }
        dest := DesktopCutNewest_UniqueDestPath(desktopPath, src)
        if (dest = "") {
            ShowCenteredOverlay_Utils("❌ Could not build Desktop path", 2500, BANNER_ACCENT_ERROR)
            return
        }
        try {
            if DirExist(src)
                DirCopy(src, dest)
            else
                FileCopy(src, dest)
        } catch {
            ShowCenteredOverlay_Utils("❌ Failed to paste to Desktop", 2500, BANNER_ACCENT_ERROR)
            return
        }
        if !FileExist(dest) {
            ShowCenteredOverlay_Utils("❌ Paste verify failed", 2500, BANNER_ACCENT_ERROR)
            return
        }
        SplitPath(dest, &destName)
        copiedNames.Push(destName)
    }

    if (copiedNames.Length = 1)
        ShowCenteredOverlay_Utils("📎 Pasted: " copiedNames[1], 1800, BANNER_ACCENT_SUCCESS)
    else
        ShowCenteredOverlay_Utils("📎 Pasted " copiedNames.Length " items to Desktop", 1800, BANNER_ACCENT_SUCCESS)
}

; --- Win+Alt+Shift+O tap / double / triple / hold ----------------------------
global g_DesktopCutNewest_TapCount := 0
global g_DesktopCutNewest_LastPressTick := 0
global g_DesktopCutNewest_TapTimer := 0
global g_DesktopCutNewest_ScheduleActivateHwnd := 0
global g_DesktopCutNewest_ScheduleActivateTitleHint := ""
global g_DesktopCutNewest_ScheduleActivateTimer := 0

class DesktopCutNewest_TapTimerObj {
    static OnTapTimeout() {
        global g_DesktopCutNewest_TapCount, g_DesktopCutNewest_TapTimer
        count := g_DesktopCutNewest_TapCount
        g_DesktopCutNewest_TapCount := 0
        g_DesktopCutNewest_TapTimer := 0
        if (count = 1)
            DesktopCutNewest_Trigger()
        else if (count = 2)
            DesktopCutNewest_OpenNewest()
    }
}

DesktopCutNewest_DisarmTapDance() {
    global g_DesktopCutNewest_TapCount, g_DesktopCutNewest_TapTimer
    global g_DesktopCutNewest_LastPressTick
    g_DesktopCutNewest_TapCount := 0
    g_DesktopCutNewest_LastPressTick := 0
    if (g_DesktopCutNewest_TapTimer) {
        SetTimer(g_DesktopCutNewest_TapTimer, 0)
        g_DesktopCutNewest_TapTimer := 0
    }
}

DesktopCutNewest_ArmTapTimer(thresholdMs) {
    global g_DesktopCutNewest_LastPressTick, g_DesktopCutNewest_TapTimer
    if (g_DesktopCutNewest_TapTimer) {
        SetTimer(g_DesktopCutNewest_TapTimer, 0)
        g_DesktopCutNewest_TapTimer := 0
    }
    g_DesktopCutNewest_LastPressTick := A_TickCount
    g_DesktopCutNewest_TapTimer := ObjBindMethod(DesktopCutNewest_TapTimerObj, "OnTapTimeout")
    SetTimer(g_DesktopCutNewest_TapTimer, -thresholdMs)
}

DesktopCutNewest_OnHotkey() {
    global g_DesktopCutNewest_TapCount, g_DesktopCutNewest_LastPressTick

    ; Hotkey fires on key-down. Drop queued auto-repeat ghosts that run after a hold
    ; released (those start with O already up and would otherwise arm single-tap cut).
    if !GetKeyState("o", "P")
        return

    thresholdMs := 400
    try thresholdMs := AI_QD_DOUBLE_TAP_MS
    catch {
        thresholdMs := 400
    }

    ; Count taps on key-down (before KeyWait) so a slow release cannot miss the window.
    pressTime := A_TickCount
    elapsed := (g_DesktopCutNewest_LastPressTick > 0) ? (pressTime - g_DesktopCutNewest_LastPressTick) : 9999
    inWindow := (g_DesktopCutNewest_TapCount > 0) && elapsed >= 0 && elapsed < thresholdMs

    KeyWait "o", "T" . (DESKTOP_CUT_NEWEST_HOLD_MS / 1000)
    isHold := (A_TickCount - pressTime) >= DESKTOP_CUT_NEWEST_HOLD_MS

    if (isHold) {
        DesktopCutNewest_DisarmTapDance()
        DesktopCutNewest_CopyPath()
        ; Stay in this thread until physical release so a repeat cannot start mid-hold
        ; and arm single-tap after we return.
        KeyWait "o"
        return
    }

    if (inWindow) {
        g_DesktopCutNewest_TapCount += 1
        if (g_DesktopCutNewest_TapCount >= 3) {
            DesktopCutNewest_DisarmTapDance()
            DesktopCutNewest_PasteClipboardToDesktop()
            return
        }
        ; 2×: wait for a possible 3× before opening.
        DesktopCutNewest_ArmTapTimer(thresholdMs)
        return
    }

    ; First tap: arm single-tap window (matches #!+8 / #!+9 timing).
    g_DesktopCutNewest_TapCount := 1
    DesktopCutNewest_ArmTapTimer(thresholdMs)
}
