; =============================================================================
; Utils/web_servers_warmup.ahk
; Act.ahk startup: pre-start Tasks (:8766) and Memory Palace (:8767) Python servers.
; Self-contained — does not #include task_helpers / palace_launcher (no UIA, no stubs).
; Best-effort, silent; exits without stopping servers (no OnExit handlers).
; =============================================================================
#Requires AutoHotkey v2.0+
#SingleInstance Off

#Include %A_ScriptDir%\..\env.ahk

global WEB_WARMUP_TASK_PORT := 8766
global WEB_WARMUP_PALACE_PORT := 8767

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
        return
    py := A_ScriptDir . "\..\tasks\python\task_server.py"
    if (!FileExist(py))
        return
    pyCmd := WebWarmup_FindPythonCmd()
    if (pyCmd = "")
        return
    dataDir := A_ScriptDir . "\..\tasks\data"
    WebWarmup_EnsureDir(dataDir)
    WebWarmup_EnsureDir(dataDir . "\attachments")
    WebWarmup_WriteTaskEnvFocus(dataDir)
    scriptsRoot := A_ScriptDir . "\.."
    cmd := pyCmd . ' -u "' . py . '" --data-dir "' . dataDir . '" --scripts-root "' . scriptsRoot
        . '" --port ' . port
    try Run(cmd, scriptsRoot, "Hide")
    catch {
        return
    }
    WebWarmup_WaitHealthy(port)
}

WebWarmup_StartPalaceServer() {
    global WEB_WARMUP_PALACE_PORT
    port := WEB_WARMUP_PALACE_PORT
    if (WebWarmup_HealthOk(port))
        return
    py := A_ScriptDir . "\..\mnemonics\python\palace_server.py"
    if (!FileExist(py))
        return
    pyCmd := WebWarmup_FindPythonCmd()
    if (pyCmd = "")
        return
    dataDir := A_ScriptDir . "\..\mnemonics\data"
    outDir := A_ScriptDir . "\..\mnemonics\output"
    scriptsRoot := A_ScriptDir . "\.."
    studiesRoot := scriptsRoot . "\mnemonics\studies"
    WebWarmup_EnsureDir(dataDir)
    WebWarmup_EnsureDir(outDir)
    cmd := pyCmd . ' -u "' . py . '" --data-dir "' . dataDir . '" --output-dir "' . outDir
        . '" --studies-root "' . studiesRoot . '" --scripts-root "' . scriptsRoot
        . '" --port ' . port
    try Run(cmd, scriptsRoot, "Hide")
    catch {
        return
    }
    WebWarmup_WaitHealthy(port)
}

WebWarmup_StartTaskServer()
WebWarmup_StartPalaceServer()
ExitApp