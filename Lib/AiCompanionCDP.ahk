; =============================================================================
; Lib: AiCompanionCDP.ahk
; Dual-mode AI Companion manager (CDP background mode vs AHK UIA foreground mode)
; and WinHttp client for infra/cdp/ai_companion_cdp_server.py (127.0.0.1:8768).
; Included from Utils.ahk and Gemini.ahk.
; =============================================================================

global AI_COMPANION_CDP_PORT := 8768

AiCompanionMode_IniPath() {
    return A_ScriptDir . "\assets\data\ai_companion_mode.ini"
}

AiCompanionMode_Get() {
    path := AiCompanionMode_IniPath()
    mode := "cdp"
    try {
        if FileExist(path)
            mode := StrLower(Trim(IniRead(path, "Mode", "Current", "cdp")))
    } catch {
        mode := "cdp"
    }
    return (mode = "ahk") ? "ahk" : "cdp"
}

AiCompanionMode_Set(mode) {
    m := (StrLower(Trim(mode)) = "ahk") ? "ahk" : "cdp"
    path := AiCompanionMode_IniPath()
    try DirCreate(A_ScriptDir . "\assets\data")
    catch {
    }
    try IniWrite(m, path, "Mode", "Current")
    catch {
    }
    return m
}

AiCompanionMode_MacroTitle() {
    if (AiCompanionMode_Get() = "cdp")
        return "🔌 AI companion mode: CDP (background) — toggle to AHK"
    return "🖱 AI companion mode: AHK (foreground) — toggle to CDP"
}

AiCompanionMode_Toggle() {
    next := (AiCompanionMode_Get() = "cdp") ? "ahk" : "cdp"
    AiCompanionMode_Set(next)
    if (next = "cdp") {
        try ShowCenteredOverlay_Utils("🔌 AI Companion Mode: CDP (Background)", 1800, BANNER_ACCENT_SUCCESS)
        catch {
        }
        try SetTimer(() => AiCompanionCDP_Ensure("", false, false), -50)
        catch {
        }
    } else {
        try ShowCenteredOverlay_Utils("🖱 AI Companion Mode: AHK (Foreground)", 1800, BANNER_ACCENT_INFO)
        catch {
        }
    }
    return next
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
    if (AiCompanionMode_Get() != "cdp")
        return false
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
    if RegExMatch(raw, '"' . key . '"\s*:\s*"((?:[^"\\]|\\.)*)"', &m)
        return _AiCDP_JsonUnescape(m[1])
    return defaultVal
}

; --- Python / Server bootstrap -----------------------------------------------

AiCompanionCDP_FindPythonExe() {
    localApp := EnvGet("LOCALAPPDATA")
    candidates := [
        localApp . "\Python\pythoncore-3.14-64\python.exe",
        "C:\Users\fie7ca\AppData\Local\Python\pythoncore-3.14-64\python.exe",
        A_WinDir . "\py.exe"
    ]
    for c in candidates {
        if FileExist(c)
            return '"' . c . '"'
    }
    return "py"
}

AiCompanionCDP_HealthOk() {
    global AI_COMPANION_CDP_PORT
    try {
        whr := ComObject("WinHttp.WinHttpRequest.5.1")
        try whr.SetProxy(1)
        whr.Open("GET", "http://127.0.0.1:" . AI_COMPANION_CDP_PORT . "/health", false)
        whr.SetTimeouts(400, 400, 1000, 1000)
        whr.Send()
        return whr.Status = 200
    } catch {
        return false
    }
}

AiCompanionCDP_EnsureServer() {
    if AiCompanionCDP_HealthOk()
        return true
    py := A_ScriptDir . "\infra\cdp\ai_companion_cdp_server.py"
    if !FileExist(py)
        return false
    pyExe := AiCompanionCDP_FindPythonExe()
    try Run(pyExe . ' -u "' . py . '"', A_ScriptDir, "Hide")
    catch {
        return false
    }
    deadline := A_TickCount + 5000
    while (A_TickCount < deadline) {
        if AiCompanionCDP_HealthOk()
            return true
        Sleep 150
    }
    return false
}

AiCompanionCDP_Post(endpoint, payload, timeoutMs := 25000) {
    global AI_COMPANION_CDP_PORT
    if !AiCompanionCDP_EnsureServer()
        return { ok: false, raw: "", error: "CDP server unavailable" }

    body := _AiCDP_JsonEncode(payload)
    try {
        whr := ComObject("WinHttp.WinHttpRequest.5.1")
        try whr.SetProxy(1)
        whr.Open("POST", "http://127.0.0.1:" . AI_COMPANION_CDP_PORT . endpoint, false)
        whr.SetRequestHeader("Content-Type", "application/json; charset=utf-8")
        whr.SetTimeouts(2000, 2000, timeoutMs, timeoutMs)
        whr.Send(body)
        raw := whr.ResponseText
        ok := _AiCDP_JsonGetBool(raw, "ok", false)
        err := _AiCDP_JsonGetStr(raw, "error", "")
        return { ok: ok, raw: raw, error: err }
    } catch as e {
        return { ok: false, raw: "", error: e.Message }
    }
}

; --- High-level Companion Operations -----------------------------------------

AiCompanionCDP_Ensure(companionId := "", activate := false, newTab := false) {
    comp := AiCompanionCDP_NormalizeCompanion(companionId)
    if (comp = "")
        return { ok: false, hwnd: 0, error: "Unsupported companion" }
    res := AiCompanionCDP_Post("/ensure", Map(
        "companion", comp,
        "activate", activate ? true : false,
        "new_tab", newTab ? true : false
    ), 25000)
    hwnd := _AiCDP_JsonGetInt(res.raw, "hwnd", 0)
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
        "image_gen", enableImageGen ? true : false
    ), 30000)
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
        return { ok: false, generating: false, replyCount: 0, composerText: "", title: "" }
    res := AiCompanionCDP_Post("/status", Map("companion", comp), 8000)
    return {
        ok: res.ok,
        generating: _AiCDP_JsonGetBool(res.raw, "generating", false),
        replyCount: _AiCDP_JsonGetInt(res.raw, "reply_count", 0),
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
    if (Trim(text) = "")
        return false
    try A_Clipboard := text
    catch {
        return false
    }
    if (playChimeAndNotify) {
        try ScriptSoundPlay(A_ScriptDir . "\assets\sounds\copy.wav")
        catch {
        }
        try ShowCenteredOverlay_Utils(isCode ? "Code copied!" : "Copied!", 800, BANNER_ACCENT_SUCCESS)
        catch {
        }
    }
    return true
}

AiCompanionCDP_NewChat(companionId := "") {
    comp := AiCompanionCDP_NormalizeCompanion(companionId)
    if (comp = "")
        return false
    res := AiCompanionCDP_Post("/new_chat", Map("companion", comp), 15000)
    return res.ok
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
