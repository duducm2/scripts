; =============================================================================
; Utils module: task_launcher.ahk
; Tasks — thin AHK launcher: start localhost server + open Chrome web app
; =============================================================================

global g_TaskDashboardHwnd := 0

Task_ServerPort() {
    return 8766
}

Task_LaunchApp() {
    Task_EnsureData()
    existing := Task_FindExistingDashboardHwnd()
    if (existing) {
        if (!Task_IsServerRunning()) {
            if (!Task_EnsureServer(false))
                return
        }
        Task_ActivateDashboard(existing)
        return
    }
    try StandardLoadingBar_Show("⏳ Opening Tasks…", BANNER_ACCENT_INTERMEDIATE, { passive: false })
    catch {
    }
    if (!Task_EnsureServer(true)) {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        return
    }
    url := "http://127.0.0.1:" . Task_ServerPort() . "/?t=" . A_TickCount
    try StandardLoadingBar_Update("⏳ Opening Chrome…", BANNER_ACCENT_INTERMEDIATE)
    catch {
    }
    ok := Task_OpenInChrome(url)
    try StandardLoadingBar_Hide(300)
    catch {
    }
    if (!ok)
        Task_Notify("Chrome failed to open Tasks", 2500, BANNER_ACCENT_ERROR)
}

Task_CloseWebApp() {
    global g_TaskDashboardHwnd
    hwnd := Task_DashboardHwndCacheGet()
    if (hwnd) {
        try WinClose("ahk_id " hwnd)
        catch {
        }
    }
    ; Also close any Chrome window titled Tasks
    for h in WinGetList("ahk_exe chrome.exe") {
        try title := WinGetTitle("ahk_id " h)
        catch {
            continue
        }
        if (Task_IsChromeWindowTitle(title)) {
            try WinClose("ahk_id " h)
            catch {
            }
        }
    }
    Task_DashboardHwndCacheClear()
    Task_StopServer()
}

Task_StopServer(port := 0) {
    if (port = 0)
        port := Task_ServerPort()
    q := Chr(39)
    cmd := "for /f `"tokens=5`" %a in (" . q . "netstat -ano ^| findstr :" . port
        . " ^| findstr LISTENING" . q . ") do taskkill /F /PID %a >nul 2>&1"
    loop 3 {
        try RunWait(A_ComSpec . " /c " . cmd, , "Hide")
        catch {
        }
        Sleep 150
        if (!Task_IsServerRunning(port))
            break
    }
}

Task_EnsureServer(forceRestart := false) {
    port := Task_ServerPort()
    if (forceRestart)
        Task_StopServer(port)
    else if (Task_IsServerRunning(port))
        return true
    py := Task_PythonDir() . "\task_server.py"
    if (!FileExist(py)) {
        Task_Notify("task_server.py not found", 2200, BANNER_ACCENT_ERROR)
        return false
    }
    pyCmd := Task_FindPythonCmd()
    if (pyCmd = "") {
        Task_Notify("Python not found for Tasks server", 2500, BANNER_ACCENT_ERROR)
        return false
    }
    dataDir := Task_DataDir()
    scriptsRoot := A_ScriptDir
    notesRoot := ""
    ; sibling notes repo
    cand := A_ScriptDir . "\..\notes"
    try {
        if (DirExist(cand))
            notesRoot := cand
    } catch {
    }
    cmd := pyCmd . ' "' . py . '" --data-dir "' . dataDir . '" --scripts-root "' . scriptsRoot
        . '" --port ' . port
    if (notesRoot != "")
        cmd .= ' --notes-root "' . notesRoot . '"'
    try Run(A_ComSpec . ' /c ' . cmd, A_ScriptDir, "Hide")
    catch as e {
        Task_Notify("Tasks server failed: " . e.Message, 2800, BANNER_ACCENT_ERROR)
        return false
    }
    loop 25 {
        if (Task_IsServerRunning(port))
            return true
        Sleep 150
    }
    Task_Notify("Tasks server did not start", 2800, BANNER_ACCENT_ERROR)
    return false
}

Task_IsServerRunning(port := 0) {
    if (port = 0)
        port := Task_ServerPort()
    try {
        whr := ComObject("WinHttp.WinHttpRequest.5.1")
        whr.Open("GET", "http://127.0.0.1:" . port . "/health", false)
        whr.Send()
        return (whr.Status = 200)
    } catch {
        return false
    }
}

Task_IsChromeWindowTitle(title) {
    t := Trim(title)
    if (t = "Tasks" || t = "Tasks - Google Chrome")
        return true
    if (InStr(t, "Tasks") = 1)
        return true
    return false
}

Task_DashboardHwndValid(hwnd) {
    try hwnd := Integer(hwnd)
    catch {
        return 0
    }
    if (!(hwnd is Integer) || hwnd <= 0)
        return 0
    try {
        if !WinExist("ahk_id " hwnd)
            return 0
    } catch {
        return 0
    }
    return hwnd
}

Task_DashboardHwndCacheGet() {
    global g_TaskDashboardHwnd
    hwnd := Task_DashboardHwndValid(g_TaskDashboardHwnd)
    if (hwnd)
        return hwnd
    raw := Trim(Task_Setting("General", "DashboardChromeHwnd", ""))
    if (raw = "")
        return 0
    hwnd := Task_DashboardHwndValid(raw)
    g_TaskDashboardHwnd := hwnd
    return hwnd
}

Task_DashboardHwndCacheSet(hwnd) {
    global g_TaskDashboardHwnd
    hwnd := Task_DashboardHwndValid(hwnd)
    g_TaskDashboardHwnd := hwnd
    Task_SetSetting("General", "DashboardChromeHwnd", hwnd ? String(hwnd) : "")
}

Task_DashboardHwndCacheClear() {
    Task_DashboardHwndCacheSet(0)
}

Task_HwndLooksLikeDashboard(hwnd) {
    hwnd := Task_DashboardHwndValid(hwnd)
    if (!hwnd)
        return 0
    try title := WinGetTitle("ahk_id " hwnd)
    catch {
        return 0
    }
    if !Task_IsChromeWindowTitle(title)
        return 0
    return hwnd
}

Task_FindExistingDashboardHwnd() {
    hwnd := Task_HwndLooksLikeDashboard(Task_DashboardHwndCacheGet())
    if (hwnd)
        return hwnd
    for h in WinGetList("ahk_exe chrome.exe") {
        hwnd := Task_HwndLooksLikeDashboard(h)
        if (hwnd) {
            Task_DashboardHwndCacheSet(hwnd)
            return hwnd
        }
    }
    Task_DashboardHwndCacheClear()
    return 0
}

Task_ActivateDashboard(hwnd) {
    hwnd := Task_DashboardHwndValid(hwnd)
    if (!hwnd)
        return false
    try {
        if (WinGetMinMax("ahk_id " hwnd) = -1)
            WinRestore("ahk_id " hwnd)
    } catch {
    }
    try WinActivate("ahk_id " hwnd)
    catch {
        return false
    }
    try WinWaitActive("ahk_id " hwnd, , 1)
    catch {
    }
    Task_DashboardHwndCacheSet(hwnd)
    return true
}

Task_OpenInChrome(url) {
    baseline := Map()
    for h in WinGetList("ahk_exe chrome.exe")
        baseline[h] := true
    try Run('chrome.exe --new-window "' . url . '"')
    catch {
        try Run(url)
        catch {
            return false
        }
    }
    deadline := A_TickCount + 8000
    newHwnd := 0
    while (A_TickCount < deadline) {
        for h in WinGetList("ahk_exe chrome.exe") {
            if (baseline.Has(h))
                continue
            try title := WinGetTitle("ahk_id " h)
            catch {
                title := ""
            }
            if (Task_IsChromeWindowTitle(title) || title = "") {
                newHwnd := h
                if (Task_IsChromeWindowTitle(title))
                    break 2
            }
        }
        Sleep 100
    }
    if (!newHwnd) {
        for h in WinGetList("ahk_exe chrome.exe") {
            try title := WinGetTitle("ahk_id " h)
            catch {
                continue
            }
            if (Task_IsChromeWindowTitle(title)) {
                newHwnd := h
                break
            }
        }
    }
    if (newHwnd) {
        Task_DashboardHwndCacheSet(newHwnd)
        try WinActivate("ahk_id " newHwnd)
        catch {
        }
        return true
    }
    return true ; URL was launched even if HWND not found
}
