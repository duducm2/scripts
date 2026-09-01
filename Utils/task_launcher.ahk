; =============================================================================
; Utils module: task_launcher.ahk
; Tasks — thin AHK launcher: start localhost server + open Chrome web app
; =============================================================================

global g_TaskDashboardHwnd := 0
global g_TaskServerPid := 0
global g_TaskPortListenCache := Map()

Task_ServerPort() {
    return 8766
}

; Work PC → work column; personal PC → personal column (IS_WORK_ENVIRONMENT).
Task_EnvDefaultFocus() {
    global IS_WORK_ENVIRONMENT
    try {
        if (IsSet(IS_WORK_ENVIRONMENT) && IS_WORK_ENVIRONMENT)
            return "work"
    } catch {
    }
    return "personal"
}

Task_WriteEnvDefaultFocus() {
    focus := Task_EnvDefaultFocus()
    path := Task_DataDir() . "\environment.txt"
    try FileDelete(path)
    catch {
    }
    try FileAppend(focus, path, "UTF-8")
    catch {
    }
    return focus
}

Task_LaunchApp() {
    Task_EnsureData()
    Task_WriteEnvDefaultFocus()
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
    . "&focus=" . Task_EnvDefaultFocus()
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

; PID currently LISTENING on port (0 if none). Prefer this over Run()'s py launcher PID.
Task_PortListeningPidProbe(port) {
    try port := Integer(port)
    catch {
        return 0
    }
    if (port <= 0)
        return 0
    output := ""
    try RunWait(A_ComSpec . ' /c netstat -ano | findstr /C:"' . port
        . '" | findstr LISTENING', &output, "Hide")
    catch {
        return 0
    }
    for line in StrSplit(output, "`n", "`r") {
        line := Trim(line)
        if (line = "")
            continue
        parts := RegExReplace(line, "\s+", " ")
        cols := StrSplit(parts, " ")
        if (cols.Length < 5)
            continue
        if (StrUpper(cols[4]) != "LISTENING")
            continue
        try return Integer(cols[5])
        catch {
        }
    }
    return 0
}

Task_PortListeningPid(port) {
    global g_TaskPortListenCache
    try port := Integer(port)
    catch {
        return 0
    }
    if (port <= 0)
        return 0
    key := String(port)
    now := A_TickCount
    if (g_TaskPortListenCache.Has(key)) {
        ent := g_TaskPortListenCache[key]
        if (now - ent["tick"] < 500)
            return ent["pid"]
    }
    pid := Task_PortListeningPidProbe(port)
    g_TaskPortListenCache[key] := Map("pid", pid, "tick", now)
    return pid
}

Task_PortIsListening(port) {
    return Task_PortListeningPid(port) > 0
}

Task_StopServerNetstat(port) {
    q := Chr(39)
    cmd := "for /f `"tokens=5`" %a in (" . q . "netstat -ano ^| findstr :" . port
        . " ^| findstr LISTENING" . q . ") do taskkill /F /PID %a >nul 2>&1"
    try RunWait(A_ComSpec . " /c " . cmd, , "Hide")
    catch {
    }
}

Task_WaitPortFree(port, timeoutMs := 2500) {
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        if (!Task_PortIsListening(port) && !Task_IsServerRunning(port))
            return true
        Sleep 50
    }
    return !Task_PortIsListening(port)
}

Task_StopServer(port := 0) {
    if (port = 0)
        port := Task_ServerPort()
    pid := Task_PidRead()
    if (pid > 0)
        Task_KillPidTree(pid)
    listenPid := Task_PortListeningPid(port)
    if (listenPid > 0 && listenPid != pid)
        Task_KillPidTree(listenPid)
    ; Always free LISTENING holders (hung sockets may fail /health).
    Task_StopServerNetstat(port)
    Task_WaitPortFree(port, 2500)
    Task_PidWrite(0)
}

Task_StartServerProcess(port) {
    py := Task_PythonDir() . "\task_server.py"
    if (!FileExist(py)) {
        Task_Notify("task_server.py not found", 2200, BANNER_ACCENT_ERROR)
        return 0
    }
    pyCmd := Task_FindPythonCmd()
    if (pyCmd = "") {
        Task_Notify("Python not found for Tasks server", 2500, BANNER_ACCENT_ERROR)
        return 0
    }
    dataDir := Task_DataDir()
    scriptsRoot := A_ScriptDir
    cmd := pyCmd . ' -u "' . py . '" --data-dir "' . dataDir . '" --scripts-root "' . scriptsRoot
        . '" --port ' . port
    pid := 0
    try Run(cmd, A_ScriptDir, "Hide", &pid)
    catch as e {
        Task_Notify("Tasks server failed: " . e.Message, 2800, BANNER_ACCENT_ERROR)
        return 0
    }
    try return Integer(pid)
    catch {
        return 0
    }
}

Task_EnsureServer(forceRestart := false) {
    port := Task_ServerPort()
    if (!forceRestart && Task_IsServerRunning(port)) {
        listenPid := Task_PortListeningPid(port)
        if (listenPid > 0)
            Task_PidWrite(listenPid)
        return true
    }

    loop 2 {
        Task_StopServer(port)
        if (Task_PortIsListening(port) || Task_IsServerRunning(port))
            Sleep 150
        if (!FileExist(Task_PythonDir() . "\task_server.py")) {
            Task_Notify("task_server.py not found", 2200, BANNER_ACCENT_ERROR)
            return false
        }
        if (Task_FindPythonCmd() = "") {
            Task_Notify("Python not found for Tasks server", 2500, BANNER_ACCENT_ERROR)
            return false
        }
        runPid := Task_StartServerProcess(port)
        if (runPid > 0)
            Task_PidWrite(runPid)

        deadline := A_TickCount + 6000
        while (A_TickCount < deadline) {
            if (Task_IsServerRunning(port)) {
                listenPid := Task_PortListeningPid(port)
                if (listenPid > 0)
                    Task_PidWrite(listenPid)
                else if (runPid > 0)
                    Task_PidWrite(runPid)
                return true
            }
            Sleep 150
        }
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
        whr.SetTimeouts(400, 400, 1200, 1200)
        whr.Send()
        return (whr.Status = 200)
    } catch {
        return false
    }
}

; Kill orphaned :8766 when Utils.ahk exits / reloads.
Task_OnExitStopServer(*) {
    try Task_StopServer()
    catch {
    }
}
OnExit(Task_OnExitStopServer, -1)

Task_IsChromeWindowTitle(title) {
    t := Trim(title)
    if (t = "")
        return false
    ; Chrome notification badge: "(1) Tasks - Google Chrome"
    t := RegExReplace(t, "^\(\d+\)\s+", "")
    if (t = "Tasks" || t = "Tasks - Google Chrome")
        return true
    if (InStr(t, "Tasks") = 1)
        return true
    ; Title still on the localhost URL (tab not fully titled yet, or URL bar mode).
    port := String(Task_ServerPort())
    if (InStr(t, "127.0.0.1:" . port) || InStr(t, "localhost:" . port))
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
