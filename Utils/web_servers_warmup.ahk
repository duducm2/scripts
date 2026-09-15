; =============================================================================
; Utils/web_servers_warmup.ahk
; Persistent keep-alive for local web services (prioritize zero cold-start latency):
;   Tasks (:8766), Memory Palace (:8767), Finance dashboard (:8765).
; Started by Act.ahk and/or Utils.ahk. Heartbeats restart any dead process.
; Self-contained — does not #include task/palace/finance launchers (no UIA).
; Does not stop Python servers on exit (Utils reload must not create cold starts).
; =============================================================================
#Requires AutoHotkey v2.0+
#SingleInstance Force

#Include %A_ScriptDir%\..\env.ahk

global WEB_WARMUP_TASK_PORT := 8766
global WEB_WARMUP_PALACE_PORT := 8767
global WEB_WARMUP_FINANCE_PORT := 8765
; Heartbeat interval — keep runtimes warm; RAM cost is acceptable.
global WEB_WARMUP_HEARTBEAT_MS := 30000
global WEB_WARMUP_HEALTH_PATHS := Map(
    8766, "/health",
    8767, "/health",
    8765, "/health"
)

WebWarmup_ScriptsRoot() {
    return A_ScriptDir . "\.."
}

WebWarmup_PidPath() {
    return WebWarmup_ScriptsRoot() . "\assets\data\web_servers_warmup.pid"
}

WebWarmup_WriteOwnPid() {
    path := WebWarmup_PidPath()
    WebWarmup_EnsureDir(WebWarmup_ScriptsRoot() . "\assets\data")
    try FileDelete(path)
    catch {
    }
    try FileAppend(String(ProcessExist()), path, "UTF-8")
    catch {
    }
}

WebWarmup_ClearOwnPid() {
    try FileDelete(WebWarmup_PidPath())
    catch {
    }
}

WebWarmup_FindPythonCmd() {
    static cached := ""
    if (cached != "")
        return cached
    candidates := ["py -3", "py", "python3", "python"]
    for c in candidates {
        try {
            ec := RunWait(A_ComSpec . ' /c ' . c . ' -c "print(1)" >nul 2>&1', , "Hide")
            if (ec = 0) {
                cached := c
                return c
            }
        } catch {
        }
    }
    localApps := EnvGet("LOCALAPPDATA")
    pathGlobs := [
        localApps . "\Programs\Python\Python3*\python.exe",
        EnvGet("ProgramFiles") . "\Python3*\python.exe",
        "C:\Python3*\python.exe"
    ]
    for g in pathGlobs {
        loop files g, "F" {
            try {
                ec := RunWait('"' . A_LoopFileFullPath . '" -c "print(1)"', , "Hide")
                if (ec = 0) {
                    cached := '"' . A_LoopFileFullPath . '"'
                    return cached
                }
            } catch {
            }
        }
    }
    return ""
}

WebWarmup_HealthOk(port) {
    global WEB_WARMUP_HEALTH_PATHS
    path := "/health"
    try {
        if (WEB_WARMUP_HEALTH_PATHS.Has(port))
            path := WEB_WARMUP_HEALTH_PATHS[port]
    } catch {
    }
    try {
        whr := ComObject("WinHttp.WinHttpRequest.5.1")
        whr.Open("GET", "http://127.0.0.1:" . port . path, false)
        whr.SetTimeouts(400, 400, 1200, 1200)
        whr.Send()
        if (whr.Status = 200)
            return true
    } catch {
    }
    ; Finance historically exposed /api/health only.
    if (port = 8765) {
        try {
            whr2 := ComObject("WinHttp.WinHttpRequest.5.1")
            whr2.Open("GET", "http://127.0.0.1:" . port . "/api/health", false)
            whr2.SetTimeouts(400, 400, 1200, 1200)
            whr2.Send()
            return (whr2.Status = 200)
        } catch {
        }
    }
    return false
}

WebWarmup_WaitHealthy(port, deadlineMs := 6000) {
    deadline := A_TickCount + deadlineMs
    while (A_TickCount < deadline) {
        if (WebWarmup_HealthOk(port))
            return true
        Sleep 150
    }
    return WebWarmup_HealthOk(port)
}

WebWarmup_EnsureDir(path) {
    if (path != "" && !DirExist(path))
        DirCreate(path)
}

WebWarmup_TaskEnvFocus() {
    global IS_WORK_ENVIRONMENT
    try {
        if (IsSet(IS_WORK_ENVIRONMENT) && IS_WORK_ENVIRONMENT)
            return "work"
    } catch {
    }
    return "personal"
}

WebWarmup_WriteTaskEnvFocus(dataDir) {
    focus := WebWarmup_TaskEnvFocus()
    path := dataDir . "\environment.txt"
    try FileDelete(path)
    catch {
    }
    try FileAppend(focus, path, "UTF-8")
    catch {
    }
}

WebWarmup_StartTaskServer() {
    global WEB_WARMUP_TASK_PORT
    port := WEB_WARMUP_TASK_PORT
    if (WebWarmup_HealthOk(port))
        return true
    root := WebWarmup_ScriptsRoot()
    py := root . "\tasks\python\task_server.py"
    if (!FileExist(py))
        return false
    pyCmd := WebWarmup_FindPythonCmd()
    if (pyCmd = "")
        return false
    dataDir := root . "\tasks\data"
    WebWarmup_EnsureDir(dataDir)
    WebWarmup_EnsureDir(dataDir . "\attachments")
    WebWarmup_WriteTaskEnvFocus(dataDir)
    cmd := pyCmd . ' -u "' . py . '" --data-dir "' . dataDir . '" --scripts-root "' . root
        . '" --port ' . port
    try Run(cmd, root, "Hide")
    catch {
        return false
    }
    return WebWarmup_WaitHealthy(port)
}

WebWarmup_StartPalaceServer() {
    global WEB_WARMUP_PALACE_PORT
    port := WEB_WARMUP_PALACE_PORT
    if (WebWarmup_HealthOk(port))
        return true
    root := WebWarmup_ScriptsRoot()
    py := root . "\mnemonics\python\palace_server.py"
    if (!FileExist(py))
        return false
    pyCmd := WebWarmup_FindPythonCmd()
    if (pyCmd = "")
        return false
    dataDir := root . "\mnemonics\data"
    outDir := root . "\mnemonics\output"
    studiesRoot := root . "\mnemonics\studies"
    WebWarmup_EnsureDir(dataDir)
    WebWarmup_EnsureDir(outDir)
    cmd := pyCmd . ' -u "' . py . '" --data-dir "' . dataDir . '" --output-dir "' . outDir
        . '" --studies-root "' . studiesRoot . '" --scripts-root "' . root
        . '" --port ' . port
    try Run(cmd, root, "Hide")
    catch {
        return false
    }
    return WebWarmup_WaitHealthy(port)
}

WebWarmup_StartFinanceServer() {
    global WEB_WARMUP_FINANCE_PORT
    port := WEB_WARMUP_FINANCE_PORT
    if (WebWarmup_HealthOk(port))
        return true
    root := WebWarmup_ScriptsRoot()
    py := root . "\finances\python\dashboard_server.py"
    if (!FileExist(py))
        return false
    pyCmd := WebWarmup_FindPythonCmd()
    if (pyCmd = "")
        return false
    dataDir := root . "\finances\data"
    outDir := root . "\finances\output"
    WebWarmup_EnsureDir(dataDir)
    WebWarmup_EnsureDir(outDir)
    ; Ensure a minimal dashboard exists so first open is not blocked on cold generate.
    if (!FileExist(outDir . "\dashboard.html")) {
        chartPy := root . "\finances\python\chart_generator.py"
        if (FileExist(chartPy)) {
            try RunWait(A_ComSpec . ' /c ' . pyCmd . ' "' . chartPy . '" --data-dir "'
                . dataDir . '" --output-dir "' . outDir . '"', root, "Hide")
            catch {
            }
        }
    }
    cmd := pyCmd . ' -u "' . py . '" --data-dir "' . dataDir . '" --output-dir "' . outDir
        . '" --port ' . port
    try Run(cmd, root, "Hide")
    catch {
        return false
    }
    return WebWarmup_WaitHealthy(port)
}

WebWarmup_EnsureAll() {
    WebWarmup_StartTaskServer()
    WebWarmup_StartPalaceServer()
    WebWarmup_StartFinanceServer()
}

WebWarmup_Heartbeat(*) {
    WebWarmup_EnsureAll()
}

WebWarmup_OnExit(*) {
    WebWarmup_ClearOwnPid()
}

WebWarmup_WriteOwnPid()
OnExit(WebWarmup_OnExit)
WebWarmup_EnsureAll()
SetTimer(WebWarmup_Heartbeat, WEB_WARMUP_HEARTBEAT_MS)
; Stay resident — persistent keep-alive (do not ExitApp).
