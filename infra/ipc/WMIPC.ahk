; WindowManagement IPC client: persistent Named Pipe connection to wm_daemon.
; Feature flags: WM_USE_DAEMON, WM_USE_PIPE_IPC, WM_USE_SHM_IPC, WM_USE_EVENT_HOOK_CACHE.
; All default off until each phase is verified.

; --- Minimal JSON for protocol (no external lib) ---
_WMIPC_JsonEscape(s) {
    s := StrReplace(s, "\\", "\\\\")
    s := StrReplace(s, "`"", "\`"")
    s := StrReplace(s, "`n", "\\n")
    s := StrReplace(s, "`r", "\\r")
    s := StrReplace(s, "`t", "\\t")
    return s
}
_WMIPC_JsonEncode(val) {
    if (val is Integer || val is Float)
        return String(val)
    if (val is String)
        return "`"" . _WMIPC_JsonEscape(val) . "`""
    if (Type(val) = "Map") {
        out := ""
        for k, v in val {
            if (out != "")
                out .= ","
            out .= "`"" . _WMIPC_JsonEscape(String(k)) . "`":" . _WMIPC_JsonEncode(v)
        }
        return "{" . out . "}"
    }
    return "null"
}
_WMIPC_JsonDecode(str) {
    try {
        o := Map()
        pos := 1
        while (pos <= StrLen(str)) {
            c := SubStr(str, pos, 1)
            if (c = " " || c = "`t" || c = "`n" || c = "`r" || c = "," || c = ":") {
                pos++
                continue
            }
            if (c = "}")
                break
            if (c = "`"") {
                keyStart := pos + 1
                pos += 2
                while (pos <= StrLen(str)) {
                    if (SubStr(str, pos, 1) = "`"")
                        break
                    if (SubStr(str, pos, 2) = "\\")
                        pos += 2
                    else
                        pos++
                }
                key := SubStr(str, keyStart, pos - keyStart)
                key := StrReplace(StrReplace(StrReplace(key, "\`"", "`""), "\\n", "`n"), "\\r", "`r")
                pos += 2
                while (pos <= StrLen(str) && InStr(" `t`n`r:", SubStr(str, pos, 1)))
                    pos++
                if (pos > StrLen(str))
                    break
                c := SubStr(str, pos, 1)
                if (c = "`"") {
                    valStart := pos + 1
                    pos += 2
                    while (pos <= StrLen(str)) {
                        if (SubStr(str, pos, 1) = "`"")
                            break
                        if (SubStr(str, pos, 2) = "\\")
                            pos += 2
                        else
                            pos++
                    }
                    val := SubStr(str, valStart, pos - valStart)
                    val := StrReplace(StrReplace(StrReplace(val, "\`"", "`""), "\\n", "`n"), "\\r", "`r")
                    o[key] := val
                    pos++
                } else if (c = "t" && SubStr(str, pos, 4) = "true") {
                    o[key] := true
                    pos += 4
                } else if (c = "f" && SubStr(str, pos, 5) = "false") {
                    o[key] := false
                    pos += 5
                } else if (c = "n" && SubStr(str, pos, 4) = "null") {
                    o[key] := ""
                    pos += 4
                } else if (c = "{" || c = "[") {
                    depth := 1
                    start := pos
                    pos++
                    while (pos <= StrLen(str) && depth > 0) {
                        ch := SubStr(str, pos, 1)
                        if (ch = "`"") {
                            pos++
                            while (pos <= StrLen(str)) {
                                if (SubStr(str, pos, 1) = "`"")
                                    break
                                if (SubStr(str, pos, 2) = "\\")
                                    pos += 2
                                else
                                    pos++
                            }
                            pos++
                        } else {
                            if (ch = "{" || ch = "[")
                                depth++
                            else if (ch = "}" || ch = "]")
                                depth--
                            pos++
                        }
                    }
                    o[key] := SubStr(str, start, pos - start)
                } else {
                    numStart := pos
                    while (pos <= StrLen(str) && InStr("0123456789.-eE", SubStr(str, pos, 1)))
                        pos++
                    o[key] := SubStr(str, numStart, pos - numStart)
                }
            }
            pos++
        }
        return o
    }
    return Map()
}

; Parse JSON array of objects "[{...},{...}]" into an array of Maps.
_WMIPC_JsonDecodeArray(str) {
    out := []
    if (Type(str) != "String" || SubStr(StrReplace(StrReplace(str, " ", ""), "`n", ""), 1, 1) != "[")
        return out
    pos := 1
    while (pos <= StrLen(str)) {
        i := InStr(str, "{", false, pos)
        if (!i)
            break
        depth := 1
        start := i
        pos := i + 1
        while (pos <= StrLen(str) && depth > 0) {
            c := SubStr(str, pos, 1)
            if (c = "`"") {
                pos++
                while (pos <= StrLen(str)) {
                    if (SubStr(str, pos, 1) = "`"")
                        break
                    if (SubStr(str, pos, 2) = "\\")
                        pos += 2
                    else
                        pos++
                }
                pos++
            } else {
                if (c = "{")
                    depth++
                else if (c = "}")
                    depth--
                pos++
            }
        }
        chunk := SubStr(str, start, pos - start)
        try
            out.Push(_WMIPC_JsonDecode(chunk))
    }
    return out
}

; --- Feature flags (default off) ---
global WM_USE_DAEMON := false
global WM_USE_PIPE_IPC := false
global WM_USE_SHM_IPC := false
global WM_USE_EVENT_HOOK_CACHE := false

; Pipe name and timeouts
global WM_PIPE_NAME := "\\.\pipe\wm_automation"
global WM_IPC_TIMEOUT_MS := 2000
global WM_IPC_RECONNECT_DELAY_MS := 500
global WM_AUTOMATION_SWITCH_DEFAULT_MS := 1800

global _WMIPC_Handle := 0

global _WM_GENERIC_READ := 0x80000000, _WM_GENERIC_WRITE := 0x40000000
global _WM_OPEN_EXISTING := 3, _WM_FILE_ATTRIBUTE_NORMAL := 0x80
global _WM_INVALID_HANDLE_VALUE := -1

; Connect to daemon. Returns true if connected.
WMIPC_Connect() {
    if (!WM_USE_DAEMON && !WM_USE_PIPE_IPC)
        return false
    try {
        h := DllCall("kernel32\CreateFileW", "Str", WM_PIPE_NAME, "UInt", _WM_GENERIC_READ | _WM_GENERIC_WRITE,
            "UInt", 0, "Ptr", 0, "UInt", _WM_OPEN_EXISTING, "UInt", _WM_FILE_ATTRIBUTE_NORMAL, "Ptr", 0, "Ptr")
        if (h != _WM_INVALID_HANDLE_VALUE && h != 0) {
            global _WMIPC_Handle := h
            return true
        }
    }
    return false
}

; Harness/standalone: connect even when flags are off (for latency benchmark).
WMIPC_ConnectForHarness() {
    try {
        h := DllCall("kernel32\CreateFileW", "Str", WM_PIPE_NAME, "UInt", _WM_GENERIC_READ | _WM_GENERIC_WRITE,
            "UInt", 0, "Ptr", 0, "UInt", _WM_OPEN_EXISTING, "UInt", _WM_FILE_ATTRIBUTE_NORMAL, "Ptr", 0, "Ptr")
        if (h != _WM_INVALID_HANDLE_VALUE && h != 0) {
            global _WMIPC_Handle := h
            return true
        }
    }
    return false
}

WMIPC_Disconnect() {
    global _WMIPC_Handle
    if (_WMIPC_Handle != 0) {
        try
            DllCall("kernel32\CloseHandle", "Ptr", _WMIPC_Handle)
        _WMIPC_Handle := 0
    }
}

; Send one request and wait for response with timeout. Returns response object or empty Map.
WMIPC_SendRequest(id, op, context := "", payload := "", deadlineMs := 0) {
    global _WMIPC_Handle
    if (_WMIPC_Handle = 0) {
        if (WMIPC_Connect())
            return WMIPC_SendRequest(id, op, context, payload, deadlineMs)
        return Map()
    }
    req := Map(
        "id", id,
        "op", op,
        "context", context,
        "payload", payload ? payload : Map(),
    "ts", A_TickCount,
    "deadlineMs", deadlineMs
    )
    jsonReq := _WMIPC_JsonEncode(req)
    if (jsonReq = "")
        return Map()
    utf8 := Buffer(StrPut(jsonReq, "UTF-8"))
    len := StrPut(jsonReq, utf8, "UTF-8")
    frameLen := 4 + len
    frame := Buffer(frameLen)
    NumPut("UChar", (len >> 24) & 0xFF, frame, 0)
    NumPut("UChar", (len >> 16) & 0xFF, frame, 1)
    NumPut("UChar", (len >> 8) & 0xFF, frame, 2)
    NumPut("UChar", len & 0xFF, frame, 3)
    loop len
        NumPut("UChar", NumGet(utf8, A_Index - 1, "UChar"), frame, 3 + A_Index)
    timeOut := (deadlineMs > 0) ? deadlineMs : WM_IPC_TIMEOUT_MS
    needReconnect := false
    try {
        written := 0
        r := DllCall("kernel32\WriteFile", "Ptr", _WMIPC_Handle, "Ptr", frame.Ptr, "UInt", frameLen, "UInt*", &written,
            "Ptr", 0)
        if (!r || written != frameLen)
            needReconnect := true
        else {
            lenBuf := Buffer(4)
            read := 0
            r := DllCall("kernel32\ReadFile", "Ptr", _WMIPC_Handle, "Ptr", lenBuf.Ptr, "UInt", 4, "UInt*", &read, "Ptr",
                0)
            if (!r || read != 4)
                needReconnect := true
            else {
                respLen := (NumGet(lenBuf, 0, "UChar") << 24) | (NumGet(lenBuf, 1, "UChar") << 16) | (NumGet(lenBuf, 2,
                    "UChar") << 8) | NumGet(lenBuf, 3, "UChar")
                if (respLen > 0 && respLen <= 1048576) {
                    respBuf := Buffer(respLen)
                    read := 0
                    r := DllCall("kernel32\ReadFile", "Ptr", _WMIPC_Handle, "Ptr", respBuf.Ptr, "UInt", respLen,
                        "UInt*", &read, "Ptr", 0)
                    if (r && read = respLen) {
                        respStr := StrGet(respBuf.Ptr, respLen, "UTF-8")
                        try
                            return _WMIPC_JsonDecode(respStr)
                        catch
                            return Map()
                    }
                }
                needReconnect := true
            }
        }
    } catch
        needReconnect := true
    if (needReconnect) {
        WMIPC_Disconnect()
        Sleep(WM_IPC_RECONNECT_DELAY_MS)
        if (WMIPC_Connect())
            return WMIPC_SendRequest(id, op, context, payload, deadlineMs)
    }
    return Map()
}

WMIPC_HealthCheck() {
    resp := WMIPC_SendRequest("HealthCheck", "HealthCheck", "", "", 1000)
    return resp.Has("ok") && resp["ok"] = true
}

; --- Phase 2: cache API (use when WM_USE_DAEMON and WM_USE_EVENT_HOOK_CACHE) ---
WMIPC_GetForegroundWindowState() {
    resp := WMIPC_SendRequest("GetForegroundWindowState", "GetForegroundWindowState", "", "", 500)
    if (!resp.Has("ok") || resp["ok"] != true || !resp.Has("result"))
        return Map("hwnd", 0, "pid", 0, "title", "", "class", "", "exe", "", "suppressCursorCentering", false,
            "cursorSuppressedUntilMs", 0, "cursorSuppressionReason", "", "lastNonGeminiHwnd", 0)
    return resp["result"] is Map ? resp["result"] : Map("hwnd", 0, "pid", 0, "title", "", "class", "", "exe", "",
        "suppressCursorCentering", false, "cursorSuppressedUntilMs", 0, "cursorSuppressionReason", "",
        "lastNonGeminiHwnd", 0)
}

WMIPC_GetCursorWindows() {
    resp := WMIPC_SendRequest("GetCursorWindows", "GetCursorWindows", "", "", 500)
    if (!resp.Has("ok") || resp["ok"] != true || !resp.Has("result"))
        return []
    r := resp["result"]
    w := r.Has("windows") ? r["windows"] : []
    if (Type(w) = "String")
        return _WMIPC_JsonDecodeArray(w)
    return IsObject(w) ? w : []
}

WMIPC_GetPreviewWindows() {
    resp := WMIPC_SendRequest("GetPreviewWindows", "GetPreviewWindows", "", "", 500)
    if (!resp.Has("ok") || resp["ok"] != true || !resp.Has("result"))
        return []
    r := resp["result"]
    w := r.Has("windows") ? r["windows"] : []
    if (Type(w) = "String")
        return _WMIPC_JsonDecodeArray(w)
    return IsObject(w) ? w : []
}

WMIPC_GetVisibleWindowsByMonitor(monitorIndex) {
    payload := Map("monitorIndex", monitorIndex)
    resp := WMIPC_SendRequest("GetVisibleWindowsByMonitor", "GetVisibleWindowsByMonitor", "", payload, 1000)
    if (!resp.Has("ok") || resp["ok"] != true || !resp.Has("result"))
        return []
    r := resp["result"]
    w := r.Has("windows") ? r["windows"] : []
    if (Type(w) = "String")
        return _WMIPC_JsonDecodeArray(w)
    return IsObject(w) ? w : []
}

WMIPC_ResolveProjectWindow(projectPath) {
    payload := Map("projectPath", projectPath)
    resp := WMIPC_SendRequest("ResolveProjectWindow", "ResolveProjectWindow", "", payload, 500)
    if (!resp.Has("ok") || resp["ok"] != true || !resp.Has("result"))
        return Map("hwnd", 0, "title", "")
    return resp["result"] is Map ? resp["result"] : Map("hwnd", 0, "title", "")
}

WMIPC_GetAutomationContext() {
    resp := WMIPC_SendRequest("GetAutomationContext", "GetAutomationContext", "", "", 500)
    if (!resp.Has("ok") || resp["ok"] != true || !resp.Has("result"))
        return Map("foregroundHwnd", 0, "foregroundTitle", "", "foregroundExe", "", "lastNonGeminiHwnd", 0,
            "lastNonGeminiTitle", "", "cursorSuppressed", false, "cursorSuppressedUntilMs", 0,
            "cursorSuppressionReason", "")
    return resp["result"] is Map ? resp["result"] : Map("foregroundHwnd", 0, "foregroundTitle", "", "foregroundExe",
        "", "lastNonGeminiHwnd", 0, "lastNonGeminiTitle", "", "cursorSuppressed", false,
        "cursorSuppressedUntilMs", 0, "cursorSuppressionReason", "")
}

WMIPC_BeginAutomationSwitch(reason := "", durationMs := 0) {
    payload := Map("reason", reason, "durationMs", durationMs > 0 ? durationMs : WM_AUTOMATION_SWITCH_DEFAULT_MS)
    resp := WMIPC_SendRequest("BeginAutomationSwitch", "BeginAutomationSwitch", "", payload, 500)
    if (!resp.Has("ok") || resp["ok"] != true || !resp.Has("result"))
        return Map("cursorSuppressed", false, "cursorSuppressedUntilMs", 0, "cursorSuppressionReason", "")
    return resp["result"] is Map ? resp["result"] : Map("cursorSuppressed", false, "cursorSuppressedUntilMs", 0,
        "cursorSuppressionReason", "")
}

WMIPC_EndAutomationSwitch(reason := "") {
    payload := Map("reason", reason)
    resp := WMIPC_SendRequest("EndAutomationSwitch", "EndAutomationSwitch", "", payload, 500)
    if (!resp.Has("ok") || resp["ok"] != true || !resp.Has("result"))
        return Map("cursorSuppressed", false, "cursorSuppressedUntilMs", 0, "cursorSuppressionReason", "")
    return resp["result"] is Map ? resp["result"] : Map("cursorSuppressed", false, "cursorSuppressedUntilMs", 0,
        "cursorSuppressionReason", "")
}
