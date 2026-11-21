; =============================================================================
; Lib: AiCompanionCDP.ahk
; Always-on CDP AI Companion manager
; and WinHttp client for infra/cdp/ai_companion_cdp_server.py (127.0.0.1:8768).
; Included from Utils.ahk and Gemini.ahk.
; =============================================================================

global AI_COMPANION_CDP_PORT := 8768
global AI_COMPANION_CDP_CHROME_PORT := 9222

_AiCDP_RepoRoot() {
    static root := ""
    if (root != "")
        return root
    try {
        SplitPath(A_LineFile, , &libDir)
        SplitPath(libDir, , &parentDir)
        if (parentDir != "" && FileExist(parentDir . "\infra\cdp\ai_companion_cdp_server.py")) {
            root := parentDir
            return root
        }
    } catch {
    }
    root := A_ScriptDir
    return root
}

AiCompanionMode_Get() {
    return "cdp"
}

AiCompanionCDP_NormalizeCompanion(companionId := "") {
    if (companionId = "") {
        try companionId := ResolveGlobalAICompanion()
        catch {
            companionId := "gemini"
        }
    }
    c := StrLower(Trim(companionId))
    if (c = "enterprise" || c = "geminienterprise")
        return "enterprise"
    if (c = "gemini")
        return "gemini"
    return ""
}

AiCompanionMode_IsCDP(companionId := "") {
    return AiCompanionCDP_NormalizeCompanion(companionId) != ""
}

; --- Minimal JSON helpers (self-contained for Utils & Gemini hosts) ----------

_AiCDP_JsonEscape(s) {
    s := StrReplace(s, "\", "\\")
    s := StrReplace(s, '"', '\"')
    s := StrReplace(s, "`n", "\n")
    s := StrReplace(s, "`r", "\r")
    s := StrReplace(s, "`t", "\t")
    return s
}

_AiCDP_JsonEncode(val) {
    if (Type(val) = "Integer" || Type(val) = "Float")
        return String(val)
    if (Type(val) = "String")
        return '"' . _AiCDP_JsonEscape(val) . '"'
    if (Type(val) = "Map") {
        out := ""
        for k, v in val {
            if (out != "")
                out .= ","
            out .= '"' . _AiCDP_JsonEscape(String(k)) . '":' . _AiCDP_JsonEncode(v)
        }
        return "{" . out . "}"
    }
    if (Type(val) = "Array") {
        out := ""
        for item in val {
            if (out != "")
                out .= ","
            out .= _AiCDP_JsonEncode(item)
        }
        return "[" . out . "]"
    }
    if (val = true)
        return "true"
    if (val = false)
        return "false"
    return "null"
}

_AiCDP_JsonUnescape(s) {
    out := ""
    i := 1
    len := StrLen(s)
    while (i <= len) {
        ch := SubStr(s, i, 1)
        if (ch = "\" && i < len) {
            nxt := SubStr(s, i + 1, 1)
            if (nxt = '"') {
                out .= '"'
                i += 2
                continue
            } else if (nxt = "\") {
                out .= "\"
                i += 2
                continue
            } else if (nxt = "n") {
                out .= "`n"
                i += 2
                continue
            } else if (nxt = "r") {
                out .= "`r"
                i += 2
                continue
            } else if (nxt = "t") {
                out .= "`t"
                i += 2
                continue
            } else if (nxt = "/") {
                out .= "/"
                i += 2
                continue
            }
        }
        out .= ch
        i++
    }
    return out
}

_AiCDP_JsonGetBool(raw, key, defaultVal := false) {
    if RegExMatch(raw, '"' . key . '"\s*:\s*(true|false)', &m)
        return (m[1] = "true")
    return defaultVal
}

_AiCDP_JsonGetInt(raw, key, defaultVal := 0) {
    if RegExMatch(raw, '"' . key . '"\s*:\s*(-?\d+)', &m)
        return Integer(m[1])
    return defaultVal
}

_AiCDP_JsonGetStr(raw, key, defaultVal := "") {
    if !RegExMatch(raw, '"' . key . '"\s*:\s*"', &m)
        return defaultVal
    pos := m.Pos + m.Len
    len := StrLen(raw)
    out := ""
    while (pos <= len) {
        nextQuote := InStr(raw, '"', , pos)
        nextSlash := InStr(raw, '\', , pos)
        if (!nextQuote)
            return defaultVal
        if (!nextSlash || nextQuote < nextSlash) {
            out .= SubStr(raw, pos, nextQuote - pos)
            return out
        }
        if (nextSlash > pos)
            out .= SubStr(raw, pos, nextSlash - pos)
        if (nextSlash >= len)
            return defaultVal
        esc := SubStr(raw, nextSlash + 1, 1)
        if (esc = '"')
            out .= '"'
        else if (esc = "\")
            out .= "\"
        else if (esc = "n")
            out .= "`n"
        else if (esc = "r")
            out .= "`r"
        else if (esc = "t")
            out .= "`t"
        else if (esc = "/")
            out .= "/"
        else if (esc = "b")
            out .= "`b"
        else if (esc = "f")
            out .= "`f"
        else if (esc = "u" && nextSlash + 5 <= len) {
            hex := SubStr(raw, nextSlash + 2, 4)
            try out .= Chr(Integer("0x" . hex))
            catch
                out .= "\u" . hex
            pos := nextSlash + 6
            continue
        } else {
            out .= esc
        }
        pos := nextSlash + 2
    }
    return defaultVal
}

; --- Python / Server bootstrap -----------------------------------------------

AiCompanionCDP_FindPythonExe() {
    static cached := ""
    if (cached != "")
        return cached
    localApp := EnvGet("LOCALAPPDATA")
    candidates := [
        localApp . "\Python\pythoncore-3.14-64\python.exe",
        "C:\Users\fie7ca\AppData\Local\Python\pythoncore-3.14-64\python.exe"
    ]
    for glob in [
        localApp . "\Programs\Python\Python3*\python.exe",
        EnvGet("ProgramFiles") . "\Python3*\python.exe",
        "C:\Python3*\python.exe"
    ] {
        loop files glob, "F"
            candidates.Push(A_LoopFileFullPath)
    }
    if (FileExist(A_WinDir . "\py.exe"))
        candidates.Push(A_WinDir . "\py.exe")
    for c in candidates {
        if !FileExist(c)
            continue
        try {
            if (RunWait('"' . c . '" -c "import playwright"', , "Hide") = 0) {
                cached := '"' . c . '"'
                return cached
            }
        } catch {
        }
    }
    for c in candidates {
        if FileExist(c) {
            cached := '"' . c . '"'
            return cached
        }
    }
    for cmd in ["python", "py"] {
        try {
            if (RunWait(cmd . ' -c "import playwright"', , "Hide") = 0) {
                cached := cmd
                return cached
            }
        } catch {
        }
    }
    return "py"
}

AiCompanionCDP_HealthOk(&cdpUp := false) {
    global AI_COMPANION_CDP_PORT
    cdpUp := false
    try {
        whr := ComObject("WinHttp.WinHttpRequest.5.1")
        try whr.SetProxy(1)
        whr.Open("GET", "http://127.0.0.1:" . AI_COMPANION_CDP_PORT . "/health", false)
        whr.SetTimeouts(800, 800, 2000, 2000)
        whr.Send()
        if (whr.Status != 200)
            return false
        if (!InStr(whr.ResponseText, '"stale"')) {
            ; Legacy CDP server without /health stale detection: terminate via WMI so EnsureServer starts the updated script.
            try {
                wmi := ComObjGet("winmgmts:")
                procs := wmi.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE CommandLine LIKE '%ai_companion_cdp_server.py%'")
                for proc in procs {
                    try ProcessClose(Integer(proc.ProcessId))
                    catch {
                    }
                }
            } catch {
            }
            Sleep 150
            return false
        }
        if (_AiCDP_JsonGetBool(whr.ResponseText, "stale", false)) {
            try {
                shut := ComObject("WinHttp.WinHttpRequest.5.1")
                try shut.SetProxy(1)
                shut.Open("POST", "http://127.0.0.1:" . AI_COMPANION_CDP_PORT . "/shutdown", false)
                shut.SetTimeouts(300, 300, 500, 500)
                shut.Send("{}")
            } catch {
            }
            Sleep 150
            return false
        }
        cdpUp := _AiCDP_JsonGetBool(whr.ResponseText, "cdp_up", true)
        return true
    } catch {
        return false
    }
}

AiCompanionCDP_KillStaleServer() {
    try {
        wmi := ComObjGet("winmgmts:")
        procs := wmi.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE CommandLine LIKE '%ai_companion_cdp_server.py%'")
        for proc in procs {
            try ProcessClose(Integer(proc.ProcessId))
            catch {
            }
        }
    } catch {
    }
}

global g_AiCompanionCDP_LastRefreshTick := 0

AiCompanionCDP_EnsureServer(&cdpUp := false) {
    if AiCompanionCDP_HealthOk(&cdpUp)
        return true
    repoRoot := _AiCDP_RepoRoot()
    py := repoRoot . "\infra\cdp\ai_companion_cdp_server.py"
    if !FileExist(py)
        return false
    pyExe := AiCompanionCDP_FindPythonExe()
    loop 3 {
        AiCompanionCDP_KillStaleServer()
        Sleep 150
        psPid := 0
        try Run(pyExe . ' -u "' . py . '"', repoRoot, "Hide", &psPid)
        catch {
            Sleep 250
            continue
        }
        deadline := A_TickCount + 8000
        while (A_TickCount < deadline) {
            if AiCompanionCDP_HealthOk(&cdpUp)
                return true
            Sleep 150
        }
    }
    return false
}

QuickUpdate_RefreshCdpServer(&cdpUp := false) {
    global g_AiCompanionCDP_LastRefreshTick
    try {
        g_AiCompanionCDP_LastRefreshTick := A_TickCount
        AiCompanionCDP_KillStaleServer()
        Sleep 120
        return AiCompanionCDP_EnsureServer(&cdpUp)
    } catch {
        return false
    }
}

AiCompanionCDP_TrySelfHealServer(&cdpUp := false, minIntervalMs := 2500) {
    global g_AiCompanionCDP_LastRefreshTick
    now := A_TickCount
    if (g_AiCompanionCDP_LastRefreshTick > 0 && (now - g_AiCompanionCDP_LastRefreshTick) < minIntervalMs) {
        return AiCompanionCDP_EnsureServer(&cdpUp)
    }
    return QuickUpdate_RefreshCdpServer(&cdpUp)
}

global g_AiCompanionCDP_LastError := ""

AiCompanionCDP_GetLastError() {
    global g_AiCompanionCDP_LastError
    return g_AiCompanionCDP_LastError
}

AiCompanionCDP_FormatError(rawErr := "", companionId := "") {
    err := Trim(rawErr)
    if (err = "")
        return "Unknown CDP error — check companion Chrome window"
    low := StrLower(err)
    compLabel := (StrLower(Trim(companionId)) = "enterprise") ? "Gemini Enterprise" : "Gemini"
    if (InStr(low, "auth_required"))
        return "Sign in required in CDP Chrome — press Win+Alt+Shift+I to open " . compLabel . " (CDP) and log in"
    if (InStr(low, "cdp_not_running") || InStr(low, "cdp_launch_failed") || InStr(low, "cdp_unreachable") || InStr(low, "econnrefused"))
        return "CDP Chrome is not open — press Win+Alt+Shift+I to open " . compLabel
    if (InStr(low, "cdp server unavailable") || InStr(low, "0x80072efd") || InStr(low, "winhttp") || InStr(low, "conexão com o servidor") || InStr(low, "connection with the server"))
        return "CDP bridge unavailable — check Python / port 8768"
    if (InStr(low, "composer_not_found"))
        return compLabel . " prompt box not found — wait for page load or sign in"
    if (InStr(low, "no_response_found"))
        return "No AI reply found in this " . compLabel . " chat yet"
    if (InStr(low, "model_not_found:")) {
        mName := Trim(SubStr(err, InStr(low, "model_not_found:") + 16))
        return "Model '" . mName . "' not found in " . compLabel . " picker"
    }
    if (InStr(low, "empty_model"))
        return "No model configured for " . compLabel . " — press Shift+L to set one"
    if (InStr(low, "read_aloud_not_found"))
        return "Read aloud button not available on last reply"
    if (InStr(low, "new_chat_not_found"))
        return "Could not start a new chat in " . compLabel . " CDP tab"
    if (InStr(low, "only_one_tab"))
        return "Only 1 " . compLabel . " tab open — keeping it open"
    if (InStr(low, "file_attach_failed"))
        return "Could not attach context file(s) in " . compLabel . " CDP Chrome"
    if (InStr(low, "paste_failed"))
        return "Could not paste text into " . compLabel . " CDP composer"
    if (InStr(low, "unsupported companion"))
        return "Current AI companion does not support CDP mode"
    return err
}

AiCompanionCDP_NotifyError(actionLabel, rawErr := "", companionId := "", durationMs := 2800) {
    global g_AiCompanionCDP_LastError
    err := (rawErr != "") ? rawErr : g_AiCompanionCDP_LastError
    msg := "❌ " . actionLabel . ": " . AiCompanionCDP_FormatError(err, companionId)
    try ShowCenteredOverlay_Utils(msg, durationMs, BANNER_ACCENT_ERROR)
    catch {
        try ToolTip(msg)
        catch {
        }
    }
}

AiCompanionCDP_Post(endpoint, payload, timeoutMs := 25000, allowLaunch := false) {
    global AI_COMPANION_CDP_PORT, g_AiCompanionCDP_LastError
    
    ; Ensure logging helper exists
    if !IsSet(SafeDebugLog)
        SafeDebugLog := (*) => 0
    cdpUp := false
    if !AiCompanionCDP_EnsureServer(&cdpUp) {
        g_AiCompanionCDP_LastError := "CDP server unavailable"
        return { ok: false, raw: "", error: "CDP server unavailable" }
    }

    if (!cdpUp && !allowLaunch) {
        g_AiCompanionCDP_LastError := "CDP_NOT_RUNNING"
        return { ok: false, raw: "", error: "CDP_NOT_RUNNING" }
    }

    isPassivePoll := (endpoint = "/status" || endpoint = "/close_tab")
    effectiveTimeout := timeoutMs
    shownColdNotice := false
    if (!cdpUp && allowLaunch && !isPassivePoll) {
        if (effectiveTimeout < 60000)
            effectiveTimeout := 60000
        try {
            StandardLoadingBar_Show("⏳ CDP Chrome is closed — starting it now (first run may take longer)…",
                BANNER_ACCENT_INTERMEDIATE, { passive: true, noActivate: true, centerOnHwnd: 0 })
            shownColdNotice := true
        } catch {
        }
    }

    startTime := A_TickCount
    ; #region agent log
    SafeDebugLog Format(
        "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Lib\AiCompanionCDP.ahk`",`"message`":`"AiCompanionCDP_Post request started`",`"data`":{`"endpoint`":`"{4}`"},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"debug-timeout`"}`n",
        A_TickCount, Random(1000, 9999), A_TickCount, endpoint)
    ; #endregion

    if (allowLaunch && payload is Map && !payload.Has("launch_if_down"))
        payload["launch_if_down"] := true

    body := _AiCDP_JsonEncode(payload)
    maxAttempts := isPassivePoll ? 1 : 3
    lastRes := { ok: false, raw: "", error: "CDP request failed" }
    loop maxAttempts {
        try {
            whr := ComObject("WinHttp.WinHttpRequest.5.1")
            try whr.SetProxy(1)
            whr.Open("POST", "http://127.0.0.1:" . AI_COMPANION_CDP_PORT . endpoint, false)
            whr.SetRequestHeader("Content-Type", "application/json; charset=utf-8")
            whr.SetTimeouts(2000, 2000, effectiveTimeout, effectiveTimeout)
            whr.Send(body)
            raw := whr.ResponseText
            ok := _AiCDP_JsonGetBool(raw, "ok", false)
            err := _AiCDP_JsonGetStr(raw, "error", "")
            g_AiCompanionCDP_LastError := ok ? "" : err
            lastRes := { ok: ok, raw: raw, error: err }

            elapsedTime := A_TickCount - startTime
            ; #region agent log
            SafeDebugLog Format(
                "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Lib\AiCompanionCDP.ahk`",`"message`":`"AiCompanionCDP_Post request completed`",`"data`":{`"endpoint`":`"{4}`",`"ok`":{5},`"elapsedTime`":{6}},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"debug-timeout`"}`n",
                A_TickCount, Random(1000, 9999), A_TickCount, endpoint, ok, elapsedTime)
            ; #endregion

            if (ok)
                break
            lowErr := StrLower(err)
            isFatal := InStr(lowErr, "auth_required") || InStr(lowErr, "unsupported") || InStr(lowErr, "empty_contact") || InStr(lowErr, "empty_model") || InStr(lowErr, "only_one_tab") || InStr(lowErr, "contact_not_found:")
            if (A_Index < maxAttempts && !isFatal) {
                needsBridgeHeal := InStr(lowErr, "target") || InStr(lowErr, "browser") || InStr(lowErr, "context") || InStr(lowErr, "closed") || InStr(lowErr, "cdp_") || InStr(lowErr, "connect") || InStr(lowErr, "timeout") || (A_Index >= 2)
                if (needsBridgeHeal) {
                    if !AiCompanionCDP_TrySelfHealServer(&cdpUp) {
                        Sleep 300
                        if !AiCompanionCDP_EnsureServer(&cdpUp) {
                            g_AiCompanionCDP_LastError := "CDP server unavailable"
                            lastRes := { ok: false, raw: "", error: "CDP server unavailable" }
                            break
                        }
                    }
                }
                Sleep 600
                continue
            }
            break
        } catch as e {
            elapsedTime := A_TickCount - startTime
            ; #region agent log
            SafeDebugLog Format(
                "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Lib\AiCompanionCDP.ahk`",`"message`":`"AiCompanionCDP_Post request failed`",`"data`":{`"endpoint`":`"{4}`",`"error`":`"{5}`",`"elapsedTime`":{6}},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"debug-timeout`"}`n",
                A_TickCount, Random(1000, 9999), A_TickCount, endpoint, e.Message, elapsedTime)
            ; #endregion

            g_AiCompanionCDP_LastError := e.Message
            lastRes := { ok: false, raw: "", error: e.Message }
            if (A_Index < maxAttempts) {
                if !AiCompanionCDP_TrySelfHealServer(&cdpUp) {
                    Sleep 300
                    if !AiCompanionCDP_EnsureServer(&cdpUp) {
                        g_AiCompanionCDP_LastError := "CDP server unavailable"
                        lastRes := { ok: false, raw: "", error: "CDP server unavailable" }
                        break
                    }
                }
                Sleep 600
                continue
            }
            break
        }
    }
    if (shownColdNotice) {
        try StandardLoadingBar_Hide(0)
        catch {
        }
    }
    return lastRes
}

; --- Direct Chrome CDP HTTP helpers (port 9222 fallback) ---------------------

AiCompanionCDP_IsChromeCdpUp(timeoutMs := 600) {
    global AI_COMPANION_CDP_CHROME_PORT
    try {
        whr := ComObject("WinHttp.WinHttpRequest.5.1")
        try whr.SetProxy(1)
        whr.Open("GET", "http://127.0.0.1:" . AI_COMPANION_CDP_CHROME_PORT . "/json/version", false)
        whr.SetTimeouts(timeoutMs, timeoutMs, timeoutMs, timeoutMs)
        whr.Send()
        return (whr.Status = 200)
    } catch {
        return false
    }
}

_AiCDP_FindChromeExe() {
    candidates := [
        EnvGet("ProgramFiles") . "\Google\Chrome\Application\chrome.exe",
        EnvGet("ProgramFiles(x86)") . "\Google\Chrome\Application\chrome.exe",
        EnvGet("LOCALAPPDATA") . "\Google\Chrome\Application\chrome.exe"
    ]
    for c in candidates {
        if (c != "" && FileExist(c))
            return '"' . c . '"'
    }
    return "chrome.exe"
}

AiCompanionCDP_DirectOpenUrl(targetUrl, matchNeedle := "", activate := false, companionId := "") {
    global AI_COMPANION_CDP_CHROME_PORT
    baseUrl := "http://127.0.0.1:" . AI_COMPANION_CDP_CHROME_PORT
    if !AiCompanionCDP_IsChromeCdpUp(800)
        return { ok: false, hwnd: 0, error: "CDP_NOT_RUNNING" }

    ; 1. Check /json/list for an already-open tab matching matchNeedle and activate it.
    if (matchNeedle != "") {
        try {
            whr := ComObject("WinHttp.WinHttpRequest.5.1")
            try whr.SetProxy(1)
            whr.Open("GET", baseUrl . "/json/list", false)
            whr.SetTimeouts(800, 800, 2000, 2000)
            whr.Send()
            if (whr.Status = 200) {
                raw := whr.ResponseText
                needleLow := StrLower(matchNeedle)
                pos := 1
                while (RegExMatch(raw, "\{[^{}]*\}", &m, pos)) {
                    obj := m[0]
                    pos := m.Pos + m.Len
                    if (_AiCDP_JsonGetStr(obj, "type", "") != "page")
                        continue
                    tabUrl := StrLower(_AiCDP_JsonGetStr(obj, "url", ""))
                    if !InStr(tabUrl, needleLow)
                        continue
                    tabId := _AiCDP_JsonGetStr(obj, "id", "")
                    if (tabId != "") {
                        try {
                            act := ComObject("WinHttp.WinHttpRequest.5.1")
                            try act.SetProxy(1)
                            act.Open("GET", baseUrl . "/json/activate/" . tabId, false)
                            act.SetTimeouts(800, 800, 2000, 2000)
                            act.Send()
                        } catch {
                        }
                        hwnd := activate ? AiCompanionCDP_ActivateWindow(0, companionId) : AiCompanionCDP_FindWindowHwnd(0, companionId)
                        return { ok: true, hwnd: hwnd, error: "" }
                    }
                }
            }
        } catch {
        }
    }

    ; 2. No matching tab yet — open a new tab directly inside the running CDP Chrome window.
    for method in ["PUT", "GET"] {
        try {
            whr := ComObject("WinHttp.WinHttpRequest.5.1")
            try whr.SetProxy(1)
            whr.Open(method, baseUrl . "/json/new?" . targetUrl, false)
            whr.SetTimeouts(1000, 1000, 4000, 4000)
            whr.Send()
            if (whr.Status >= 200 && whr.Status < 300) {
                hwnd := activate ? AiCompanionCDP_ActivateWindow(0, companionId) : AiCompanionCDP_FindWindowHwnd(0, companionId)
                return { ok: true, hwnd: hwnd, error: "" }
            }
        } catch {
        }
    }
    return { ok: false, hwnd: 0, error: "CDP_DIRECT_OPEN_FAILED" }
}

AiCompanionCDP_DirectLaunchChrome(targetUrl, minimized := true, waitMs := 15000, companionId := "") {
    if AiCompanionCDP_IsChromeCdpUp(600)
        return AiCompanionCDP_DirectOpenUrl(targetUrl, "", !minimized, companionId)

    cdpProfile := EnvGet("LOCALAPPDATA") . "\GeminiCDP"
    chromeExe := _AiCDP_FindChromeExe()
    cmd := chromeExe . ' --user-data-dir="' . cdpProfile . '" --remote-debugging-port=9222 --disable-features=CalculateNativeWindowOcclusion --disable-backgrounding-occluded-windows'
    if (minimized)
        cmd .= " --start-minimized"
    cmd .= ' "' . targetUrl . '"'

    try Run(cmd, , minimized ? "Min" : "")
    catch {
        return { ok: false, hwnd: 0, error: "CDP_LAUNCH_FAILED" }
    }

    deadline := A_TickCount + waitMs
    while (A_TickCount < deadline) {
        if AiCompanionCDP_IsChromeCdpUp(500) {
            hwnd := minimized ? AiCompanionCDP_FindWindowHwnd(0, companionId) : AiCompanionCDP_ActivateWindow(0, companionId)
            return { ok: true, hwnd: hwnd, error: "" }
        }
        Sleep 300
    }
    return { ok: false, hwnd: 0, error: "CDP_LAUNCH_FAILED" }
}

; --- High-level Companion Operations -----------------------------------------

AiCompanionCDP_GetProcessPids() {
    pids := Map()
    try {
        wmi := ComObjGet("winmgmts:")
        procs := wmi.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name='chrome.exe' AND (CommandLine LIKE '%GeminiCDP%' OR CommandLine LIKE '%remote-debugging-port=9222%')")
        for proc in procs {
            try pid := Integer(proc.ProcessId)
            catch
                continue
            if (pid > 0)
                pids[pid] := true
        }
    } catch {
    }
    return pids
}

AiCompanionCDP_FindWindowHwnd(preferredHwnd := 0, companionId := "") {
    cdpPids := AiCompanionCDP_GetProcessPids()

    if (preferredHwnd && WinExist("ahk_id " preferredHwnd)) {
        if (cdpPids.Count = 0)
            return preferredHwnd
        try {
            winPid := WinGetPID("ahk_id " preferredHwnd)
            if cdpPids.Has(winPid)
                return preferredHwnd
        } catch {
        }
    }

    ; 1. Prefer the Chrome process launched with the dedicated GeminiCDP profile / port 9222.
    for pid, _ in cdpPids {
        try {
            for hwnd in WinGetList("ahk_pid " pid " ahk_class Chrome_WidgetWin_1") {
                title := ""
                try title := WinGetTitle("ahk_id " hwnd)
                if (title != "")
                    return hwnd
            }
        } catch {
        }
    }

    ; 2. Fallback only when WMI PID lookup returned empty: match visible Chrome_WidgetWin_1 windows by companion/chat title.
    if (cdpPids.Count > 0)
        return 0
    rawComp := StrLower(Trim(companionId))
    comp := (rawComp = "whatsapp" || rawComp = "teams") ? rawComp : AiCompanionCDP_NormalizeCompanion(companionId)
    for hwnd in WinGetList("ahk_exe chrome.exe ahk_class Chrome_WidgetWin_1") {
        title := ""
        try title := WinGetTitle("ahk_id " hwnd)
        if (title = "")
            continue
        if (comp = "enterprise") {
            if (InStr(title, "Vertex AI Search") || InStr(title, "AskBosch") || InStr(title, "Gemini Enterprise"))
                return hwnd
        } else if (comp = "whatsapp") {
            if InStr(title, "WhatsApp")
                return hwnd
        } else if (comp = "teams") {
            if InStr(title, "Teams")
                return hwnd
        } else if (InStr(title, "Gemini")) {
            return hwnd
        }
    }
    return 0
}

AiCompanionCDP_ActivateWindow(preferredHwnd := 0, companionId := "") {
    hwnd := 0
    loop 3 {
        hwnd := AiCompanionCDP_FindWindowHwnd(A_Index = 1 ? preferredHwnd : 0, companionId)
        if (!hwnd || !WinExist("ahk_id " hwnd)) {
            Sleep 150
            continue
        }
        try {
            if (WinGetMinMax("ahk_id " hwnd) = -1) {
                WinRestore("ahk_id " hwnd)
                Sleep 60
            }
            if !WinActive("ahk_id " hwnd) {
                WinActivate("ahk_id " hwnd)
                WinWaitActive("ahk_id " hwnd, , 1.0)
            }
            if !WinActive("ahk_id " hwnd) {
                DllCall("User32\ShowWindow", "Ptr", hwnd, "Int", 9)
                DllCall("User32\SetForegroundWindow", "Ptr", hwnd)
                WinActivate("ahk_id " hwnd)
                WinWaitActive("ahk_id " hwnd, , 0.8)
            }
            if WinActive("ahk_id " hwnd)
                return hwnd
        } catch {
        }
        Sleep 150
    }
    return hwnd
}

AiCompanionCDP_Ensure(companionId := "", activate := false, newTab := false) {
    comp := AiCompanionCDP_NormalizeCompanion(companionId)
    if (comp = "")
        return { ok: false, hwnd: 0, error: "Unsupported companion" }
    res := AiCompanionCDP_Post("/ensure", Map(
        "companion", comp,
        "activate", activate ? true : false,
        "new_tab", newTab ? true : false,
        "launch_if_down", true
    ), 25000, true)
    hwnd := _AiCDP_JsonGetInt(res.raw, "hwnd", 0)
    if (res.ok && activate)
        hwnd := AiCompanionCDP_ActivateWindow(hwnd, comp)
    return { ok: res.ok, hwnd: hwnd, error: res.error }
}

AiCompanionCDP_Paste(companionId, text, submit := false, files := "", enableImageGen := false, clear := false) {
    comp := AiCompanionCDP_NormalizeCompanion(companionId)
    if (comp = "")
        return { ok: false, started: false, hwnd: 0, error: "Unsupported companion" }
    fileArr := IsObject(files) ? files : []
    res := AiCompanionCDP_Post("/paste", Map(
        "companion", comp,
        "text", text,
        "files", fileArr,
        "submit", submit ? true : false,
        "clear", clear ? true : false,
        "image_gen", enableImageGen ? true : false,
        "launch_if_down", true
    ), 45000, true)
    return {
        ok: res.ok,
        started: _AiCDP_JsonGetBool(res.raw, "started", false),
        replyCountBefore: _AiCDP_JsonGetInt(res.raw, "reply_count_before", -1),
        hwnd: _AiCDP_JsonGetInt(res.raw, "hwnd", 0),
        error: res.error
    }
}

AiCompanionCDP_Submit(companionId := "") {
    comp := AiCompanionCDP_NormalizeCompanion(companionId)
    if (comp = "")
        return false
    res := AiCompanionCDP_Post("/submit", Map("companion", comp), 15000)
    return res.ok && _AiCDP_JsonGetBool(res.raw, "started", true)
}

AiCompanionCDP_Status(companionId := "") {
    comp := AiCompanionCDP_NormalizeCompanion(companionId)
    if (comp = "")
        return { ok: false, generating: false, replyCount: 0, lastReplyLen: 0, composerText: "", title: "" }
    res := AiCompanionCDP_Post("/status", Map("companion", comp), 8000)
    return {
        ok: res.ok,
        generating: _AiCDP_JsonGetBool(res.raw, "generating", false),
        replyCount: _AiCDP_JsonGetInt(res.raw, "reply_count", 0),
        lastReplyLen: _AiCDP_JsonGetInt(res.raw, "last_reply_len", 0),
        composerText: _AiCDP_JsonGetStr(res.raw, "composer_text", ""),
        title: _AiCDP_JsonGetStr(res.raw, "title", "")
    }
}

AiCompanionCDP_CopyLast(companionId := "", isCode := false, playChimeAndNotify := false) {
    comp := AiCompanionCDP_NormalizeCompanion(companionId)
    if (comp = "")
        return false
    res := AiCompanionCDP_Post("/copy_last", Map(
        "companion", comp,
        "kind", isCode ? "code" : "message"
    ), 12000)
    if (!res.ok)
        return false
    text := _AiCDP_JsonGetStr(res.raw, "text", "")
    if (Trim(text) = "") {
        global g_AiCompanionCDP_LastError
        g_AiCompanionCDP_LastError := "NO_RESPONSE_FOUND"
        return false
    }
    try A_Clipboard := text
    catch {
        return false
    }
    if (playChimeAndNotify) {
        try ScriptSoundPlay(_AiCDP_RepoRoot() . "\assets\sounds\copy.wav")
        catch {
        }
        try ShowCenteredOverlay_Utils(isCode ? "Code copied!" : "Copied!", 800, BANNER_ACCENT_SUCCESS)
        catch {
        }
    }
    return true
}

AiCompanionCDP_NewChat(companionId := "", activeTitle := "") {
    comp := AiCompanionCDP_NormalizeCompanion(companionId)
    if (comp = "")
        return false
    payload := Map("companion", comp, "launch_if_down", true)
    if (Trim(activeTitle) != "")
        payload["active_title"] := Trim(activeTitle)
    res := AiCompanionCDP_Post("/new_chat", payload, 15000, true)
    return res.ok
}

AiCompanionCDP_CloseTab(companionId := "") {
    comp := AiCompanionCDP_NormalizeCompanion(companionId)
    if (comp = "")
        return { ok: false, remaining: 0, error: "Unsupported companion" }
    res := AiCompanionCDP_Post("/close_tab", Map("companion", comp), 10000)
    return {
        ok: res.ok,
        remaining: _AiCDP_JsonGetInt(res.raw, "remaining", 0),
        error: res.error
    }
}

AiCompanionCDP_SelectModel(companionId, modelName) {
    comp := AiCompanionCDP_NormalizeCompanion(companionId)
    if (comp = "")
        return false
    res := AiCompanionCDP_Post("/select_model", Map(
        "companion", comp,
        "model", modelName
    ), 15000)
    return res.ok
}

AiCompanionCDP_ReadAloud(companionId := "") {
    comp := AiCompanionCDP_NormalizeCompanion(companionId)
    if (comp = "")
        return false
    res := AiCompanionCDP_Post("/read_aloud", Map("companion", comp), 10000)
    return res.ok
}
