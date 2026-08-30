; =============================================================================
; Utils module: task_launcher.ahk
; Tasks — thin AHK launcher: start localhost server + open Chrome web app
; =============================================================================

global g_TaskDashboardHwnd := 0
global g_TaskServerPid := 0

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
    if (!Task_EnsureServer(false)) {
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

Task_PidPath() {
    return Task_DataDir() . "\task_server.pid"
}

Task_PidRead() {
    global g_TaskServerPid
    pid := 0
    try pid := Integer(g_TaskServerPid)
    catch {
        pid := 0
    }
    if (pid > 0) {
        try {
            if ProcessExist(pid)
                return pid
        } catch {
        }
    }
    path := Task_PidPath()
    if (!FileExist(path))
        return 0
    raw := ""
    try raw := Trim(FileRead(path))
    catch {
        return 0
    }
    try pid := Integer(raw)
    catch {
        return 0
    }
    if (pid <= 0)
        return 0
    try {
        if ProcessExist(pid)
            return pid
    } catch {
    }
    return 0
}

Task_PidWrite(pid) {
    global g_TaskServerPid
    try pid := Integer(pid)
    catch {
        pid := 0
    }
    g_TaskServerPid := pid
    path := Task_PidPath()
    if (pid > 0) {
        try FileDelete(path)
        catch {
        }
        try FileAppend(String(pid), path, "UTF-8")
        catch {
        }
        return
    }
    try FileDelete(path)
    catch {
    }
}

Task_KillPidTree(pid) {
    try pid := Integer(pid)
    catch {
        return
    }
    if (pid <= 0)
        return
    try RunWait(A_ComSpec . " /c taskkill /F /T /PID " . pid . " >nul 2>&1", , "Hide")
    catch {
    }
}

Task_StopServerNetstat(port) {
    q := Chr(39)
    cmd := "for /f `"tokens=5`" %a in (" . q . "netstat -ano ^| findstr :" . port
        . " ^| findstr LISTENING" . q . ") do taskkill /F /PID %a >nul 2>&1"
    try RunWait(A_ComSpec . " /c " . cmd, , "Hide")
    catch {
    }
}

Task_StopServer(port := 0) {
    if (port = 0)
        port := Task_ServerPort()
    pid := Task_PidRead()
    if (pid > 0)
        Task_KillPidTree(pid)
    deadline := A_TickCount + 1500
    while (A_TickCount < deadline) {
        if (!Task_IsServerRunning(port))
            break
        Sleep 50
    }
    if (Task_IsServerRunning(port))
        Task_StopServerNetstat(port)
    deadline := A_TickCount + 1500
    while (A_TickCount < deadline) {
        if (!Task_IsServerRunning(port))
            break
        Sleep 50
    }
    Task_PidWrite(0)
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
    cmd := pyCmd . ' "' . py . '" --data-dir "' . dataDir . '" --scripts-root "' . scriptsRoot
        . '" --port ' . port
    pid := 0
    try pid := Run(cmd, A_ScriptDir, "Hide")
    catch as e {
        Task_Notify("Tasks server failed: " . e.Message, 2800, BANNER_ACCENT_ERROR)
        return false
    }
    if (pid > 0)
        Task_PidWrite(pid)
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
    try Run('chrome.exe --new-window "' . url . '"')
    catch {
        try Run(url)
        catch {
            return false
        }
    }
    newHwnd := 0
    prev := A_TitleMatchMode
    try {
        SetTitleMatchMode(1)
        if WinWait("Tasks ahk_exe chrome.exe", , 8) {
            try newHwnd := WinExist("Tasks ahk_exe chrome.exe")
            catch {
                newHwnd := 0
            }
        }
    } finally {
        SetTitleMatchMode(prev)
    }
    newHwnd := Task_HwndLooksLikeDashboard(newHwnd)
    if (!newHwnd) {
        for h in WinGetList("ahk_exe chrome.exe") {
            newHwnd := Task_HwndLooksLikeDashboard(h)
            if (newHwnd)
                break
        }
    }
    if (newHwnd) {
        Task_DashboardHwndCacheSet(newHwnd)
        try WinActivate("ahk_id " newHwnd)
        catch {
        }
        return true
    }
    return true
}
