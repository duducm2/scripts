; Gemini IPC client: persistent Named Pipe connection to gemini_daemon.
; Keeps Gemini browser/TTS work out of the AHK foreground thread when enabled.

_GeminiIPC_JsonEscape(s) {
    s := StrReplace(s, "\\", "\\\\")
    s := StrReplace(s, "`"", "\`"")
    s := StrReplace(s, "`n", "\\n")
    s := StrReplace(s, "`r", "\\r")
    s := StrReplace(s, "`t", "\\t")
    return s
}

_GeminiIPC_JsonEncode(val) {
    if (val is Integer || val is Float)
        return String(val)
    if (val is String)
        return "`"" . _GeminiIPC_JsonEscape(val) . "`""
    if (Type(val) = "Map") {
        out := ""
        for k, v in val {
            if (out != "")
                out .= ","
            out .= "`"" . _GeminiIPC_JsonEscape(String(k)) . "`":" . _GeminiIPC_JsonEncode(v)
        }
        return "{" . out . "}"
    }
    if (Type(val) = "Array") {
        out := ""
        for item in val {
            if (out != "")
                out .= ","
            out .= _GeminiIPC_JsonEncode(item)
        }
        return "[" . out . "]"
    }
    if (val = true)
        return "true"
    if (val = false)
        return "false"
    return "null"
}

_GeminiIPC_JsonDecode(str) {
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

global GEMINI_PIPE_NAME := "\\.\pipe\gemini_automation"
global GEMINI_IPC_TIMEOUT_MS := 2500
global GEMINI_IPC_RECONNECT_DELAY_MS := 500
global GEMINI_DAEMON_START_WAIT_MS := 6000
global GEMINI_DAEMON_HEALTH_POLL_MS := 400
global _GeminiIPC_Handle := 0

global _Gemini_GENERIC_READ := 0x80000000, _Gemini_GENERIC_WRITE := 0x40000000
global _Gemini_OPEN_EXISTING := 3, _Gemini_FILE_ATTRIBUTE_NORMAL := 0x80
global _Gemini_INVALID_HANDLE_VALUE := -1

GeminiIPC_Connect() {
    try {
        h := DllCall("kernel32\CreateFileW", "Str", GEMINI_PIPE_NAME, "UInt", _Gemini_GENERIC_READ |
            _Gemini_GENERIC_WRITE,
            "UInt", 0, "Ptr", 0, "UInt", _Gemini_OPEN_EXISTING, "UInt", _Gemini_FILE_ATTRIBUTE_NORMAL, "Ptr", 0, "Ptr")
        if (h != _Gemini_INVALID_HANDLE_VALUE && h != 0) {
            global _GeminiIPC_Handle := h
            return true
        }
    }
    return false
}

GeminiIPC_Disconnect() {
    global _GeminiIPC_Handle
    if (_GeminiIPC_Handle != 0) {
        try
            DllCall("kernel32\CloseHandle", "Ptr", _GeminiIPC_Handle)
        _GeminiIPC_Handle := 0
    }
}

; True if client already has an open pipe handle (caller must not force GeminiIPC_EnsureReady → blocking CreateFile).
GeminiIPC_HasOpenPipe() {
    global _GeminiIPC_Handle
    return _GeminiIPC_Handle != 0
}

; Bytes currently readable from the pipe (client side). Returns 0 if unavailable or error.
GeminiIPC_PeekPipeBytesAvailable(handle) {
    avail := 0
    if !DllCall("kernel32\PeekNamedPipe", "Ptr", handle, "Ptr", 0, "UInt", 0, "Ptr", 0, "UInt*", &avail, "Ptr", 0)
        return 0
    return avail
}

; Poll until at least minBytes are available or timeoutMs elapses (avoids infinite blocking ReadFile).
GeminiIPC_WaitForPipeBytes(handle, minBytes, timeoutMs) {
    if (minBytes <= 0)
        return true
    if (timeoutMs < 1)
        return false
    t0 := A_TickCount
    while ((A_TickCount - t0) < timeoutMs) {
        if (GeminiIPC_PeekPipeBytesAvailable(handle) >= minBytes)
            return true
        Sleep 15
    }
    return false
}

GeminiIPC_SendRequest(id, op, payload := "", deadlineMs := 0) {
    global _GeminiIPC_Handle
    if (_GeminiIPC_Handle = 0) {
        if (GeminiIPC_Connect())
            return GeminiIPC_SendRequest(id, op, payload, deadlineMs)
        return Map()
    }
    req := Map(
        "id", id,
        "op", op,
        "context", "Gemini.ahk",
        "payload", payload ? payload : Map(),
    "ts", A_TickCount,
    "deadlineMs", deadlineMs
    )
    jsonReq := _GeminiIPC_JsonEncode(req)
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
    readBudget := deadlineMs > 0 ? deadlineMs : GEMINI_IPC_TIMEOUT_MS
    if (readBudget < 400)
        readBudget := 400
    if (readBudget > 120000)
        readBudget := 120000
    deadlineTick := A_TickCount + readBudget
    needReconnect := false
    try {
        written := 0
        r := DllCall("kernel32\WriteFile", "Ptr", _GeminiIPC_Handle, "Ptr", frame.Ptr, "UInt", frameLen, "UInt*",
            &written, "Ptr", 0)
        if (!r || written != frameLen)
            needReconnect := true
        else {
            remMs := deadlineTick - A_TickCount
            if (remMs < 1)
                remMs := 1
            if !GeminiIPC_WaitForPipeBytes(_GeminiIPC_Handle, 4, remMs)
                needReconnect := true
            else {
                lenBuf := Buffer(4)
                read := 0
                r := DllCall("kernel32\ReadFile", "Ptr", _GeminiIPC_Handle, "Ptr", lenBuf.Ptr, "UInt", 4, "UInt*", &
                    read, "Ptr", 0)
                if (!r || read != 4)
                    needReconnect := true
                else {
                    respLen := (NumGet(lenBuf, 0, "UChar") << 24) | (NumGet(lenBuf, 1, "UChar") << 16) | (NumGet(lenBuf,
                        2,
                        "UChar") << 8) | NumGet(lenBuf, 3, "UChar")
                    if (respLen > 0 && respLen <= 1048576) {
                        remMs := deadlineTick - A_TickCount
                        if (remMs < 1)
                            remMs := 1
                        if !GeminiIPC_WaitForPipeBytes(_GeminiIPC_Handle, respLen, remMs)
                            needReconnect := true
                        else {
                            respBuf := Buffer(respLen)
                            read := 0
                            r := DllCall("kernel32\ReadFile", "Ptr", _GeminiIPC_Handle, "Ptr", respBuf.Ptr, "UInt",
                                respLen,
                                "UInt*", &read, "Ptr", 0)
                            if (r && read = respLen) {
                                respStr := StrGet(respBuf.Ptr, respLen, "UTF-8")
                                try
                                    return _GeminiIPC_JsonDecode(respStr)
                                catch
                                    return Map()
                            }
                        }
                    }
                    if (!needReconnect)
                        needReconnect := true
                }
            }
        }
    } catch
        needReconnect := true
    if (needReconnect) {
        GeminiIPC_Disconnect()
        Sleep(GEMINI_IPC_RECONNECT_DELAY_MS)
        if (GeminiIPC_Connect())
            return GeminiIPC_SendRequest(id, op, payload, deadlineMs)
    }
    return Map()
}

GeminiIPC_ResponseOk(resp) {
    return resp.Has("ok") && resp["ok"] = true
}

GeminiIPC_ResponseResultMap(resp) {
    if (!resp.Has("result"))
        return Map()
    result := resp["result"]
    if (Type(result) = "Map")
        return result
    if (Type(result) = "String" && result != "" && SubStr(result, 1, 1) = "{")
        return _GeminiIPC_JsonDecode(result)
    return Map()
}

GeminiIPC_HealthCheck() {
    resp := GeminiIPC_SendRequest("HealthCheck", "HealthCheck", "", 1000)
    return GeminiIPC_ResponseOk(resp)
}

GeminiIPC_StartDaemon() {
    daemonPath := A_ScriptDir "\infra\python\gemini_daemon.py"
    if (!FileExist(daemonPath))
        return false
    ; Never use cmd.exe /c start here: it opens a console window. If the user clicks it, Windows
    ; QuickEdit ("Selecionar") suspends the child process and the Named Pipe daemon stops responding.
    dq := Chr(34)
    cmdLineVariants := []
    ; Prefer pythonw / pyw — no console window, cannot get stuck in QuickEdit.
    cmdLineVariants.Push("pythonw.exe " . dq . daemonPath . dq)
    cmdLineVariants.Push("pyw.exe -3 " . dq . daemonPath . dq)
    localAppData := EnvGet("LOCALAPPDATA")
    for suffix in ["313", "312", "311", "310"] {
        pw := localAppData "\Programs\Python\Python" . suffix . "\pythonw.exe"
        if (FileExist(pw))
            cmdLineVariants.Push(dq . pw . dq . " " . dq . daemonPath . dq)
    }
    ; Last resort: python.exe with Hide (may briefly allocate a console on some setups).
    cmdLineVariants.Push("python.exe " . dq . daemonPath . dq)
    cmdLineVariants.Push("py.exe -3 " . dq . daemonPath . dq)
    for cmdLine in cmdLineVariants {
        try {
            Run(cmdLine, A_ScriptDir, "Hide")
            return true
        } catch {
            continue
        }
    }
    return false
}

GeminiIPC_EnsureReady(waitMs := 0) {
    if (!GEMINI_USE_PYTHON_IPC)
        return false
    if (GeminiIPC_HealthCheck())
        return true
    GeminiIPC_StartDaemon()
    maxWaitMs := waitMs > 0 ? waitMs : GEMINI_DAEMON_HEALTH_POLL_MS
    waited := 0
    while (waited < maxWaitMs) {
        Sleep GEMINI_DAEMON_HEALTH_POLL_MS
        if (GeminiIPC_HealthCheck())
            return true
        waited += GEMINI_DAEMON_HEALTH_POLL_MS
    }
    return false
}

GeminiIPC_QueueTask(taskKind, payload := "") {
    reqPayload := payload ? payload : Map()
    if (!(reqPayload is Map))
        reqPayload := Map()
    reqPayload["taskKind"] := taskKind
    return GeminiIPC_SendRequest("QueueTask", "QueueTask", reqPayload, 2000)
}

GeminiIPC_GetTaskStatus(taskId) {
    return GeminiIPC_SendRequest("GetTaskStatus", "GetTaskStatus", Map("taskId", taskId), 2000)
}

; Returns ISO 639-1 code "pt" | "en" | "de" from gemini_daemon (lingua), or "" if IPC/daemon fails (caller may use heuristic).
GeminiIPC_DetectLang(text, fallback := "") {
    if (!GeminiIPC_EnsureReady(GEMINI_DAEMON_START_WAIT_MS)) {
        return ""
    }
    ; First lingua load can exceed 1.5s; deadline applies to full framed response wait (PeekNamedPipe + reads).
    resp := GeminiIPC_SendRequest("DetectLang", "DetectLang", Map("text", text), 12000)
    rok := GeminiIPC_ResponseOk(resp)
    if (!rok)
        return ""
    result := GeminiIPC_ResponseResultMap(resp)
    if (result.Has("language") && result["language"] != "")
        return result["language"]
    return fallback
}
