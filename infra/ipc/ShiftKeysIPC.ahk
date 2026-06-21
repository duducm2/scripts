; ShiftKeys IPC client: persistent Named Pipe connection to shiftkeys_daemon.
; Feature flags: USE_DAEMON, USE_PIPE_IPC, USE_SHM_IPC. Fallback to legacy when daemon unavailable.
;
; Phase 5 rollout: enable one at a time and validate before full cutover.
; - USE_DAEMON_CONTEXT_CHATGPT: O(1) #HotIf for ChatGPT (requires daemon + context mirror timer).
; - USE_DAEMON_MONITOR_GEMINI: non-blocking Enter/^Enter (daemon watch; chime on completion/timeout).
; KPI: predicate evaluation O(1); no long synchronous waits on AHK thread; daemon reconnect on failure.

; --- Minimal JSON for protocol (no external lib) ---
_ShiftKeysIPC_JsonEscape(s) {
    s := StrReplace(s, "\\", "\\\\")
    s := StrReplace(s, "`"", "\`"")
    s := StrReplace(s, "`n", "\\n")
    s := StrReplace(s, "`r", "\\r")
    s := StrReplace(s, "`t", "\\t")
    return s
}
_ShiftKeysIPC_JsonEncode(val) {
    if (val is Integer || val is Float)
        return String(val)
    if (val is String)
        return "`"" . _ShiftKeysIPC_JsonEscape(val) . "`""
    if (Type(val) = "Map") {
        out := ""
        for k, v in val {
            if (out != "")
                out .= ","
            out .= "`"" . _ShiftKeysIPC_JsonEscape(String(k)) . "`":" . _ShiftKeysIPC_JsonEncode(v)
        }
        return "{" . out . "}"
    }
    return "null"
}
_ShiftKeysIPC_JsonDecode(str) {
    try {
        ; AHK v2 has no built-in JSON; use ComObject("Scripting.Dictionary") or manual parse.
        ; Minimal parse: assume response is {"id":"...","ok":true,...}. We only need to read top-level keys.
        o := Map()
        pos := 1
        while (pos <= StrLen(str)) {
            c := SubStr(str, pos, 1)
            if (c = " " || c = "`t" || c = "`n" || c = "`r" || c = "," || c = ":") {
                pos++
                continue
            }
            if (c = "}") {
                break
            }
            if (c = "`"") {
                keyStart := pos + 1
                pos += 2
                while (pos <= StrLen(str)) {
                    if (SubStr(str, pos, 1) = "`"") {
                        break
                    }
                    if (SubStr(str, pos, 2) = "\\") {
                        pos += 2
                    } else {
                        pos++
                    }
                }
                key := SubStr(str, keyStart, pos - keyStart)
                key := StrReplace(StrReplace(StrReplace(key, "\`"", "`""), "\\n", "`n"), "\\r", "`r")
                pos += 2
                ; skip colon and whitespace
                while (pos <= StrLen(str) && InStr(" `t`n`r:", SubStr(str, pos, 1)))
                    pos++
                if (pos > StrLen(str)) {
                    break
                }
                c := SubStr(str, pos, 1)
                if (c = "`"") {
                    valStart := pos + 1
                    pos += 2
                    while (pos <= StrLen(str)) {
                        if (SubStr(str, pos, 1) = "`"") {
                            break
                        }
                        if (SubStr(str, pos, 2) = "\\") {
                            pos += 2
                        } else {
                            pos++
                        }
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
                    ; nested: skip until matching brace/bracket
                    depth := 1
                    start := pos
                    pos++
                    while (pos <= StrLen(str) && depth > 0) {
                        ch := SubStr(str, pos, 1)
                        if (ch = "`"") {
                            pos++
                            while (pos <= StrLen(str)) {
                                if (SubStr(str, pos, 1) = "`"") {
                                    break
                                }
                                if (SubStr(str, pos, 2) = "\\") {
                                    pos += 2
                                } else {
                                    pos++
                                }
                            }
                            pos++
                        } else {
                            if (ch = "{" || ch = "[") {
                                depth++
                            } else if (ch = "}" || ch = "]") {
                                depth--
                            }
                            pos++
                        }
                    }
                    o[key] := SubStr(str, start, pos - start)
                } else {
                    ; number
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

; --- Feature flags (set in env.ahk or before this include to override) ---
global USE_DAEMON := true
global USE_PIPE_IPC := true
global USE_SHM_IPC := false

; Pipe name and timeouts
global SHIFTKEYS_PIPE_NAME := "\\.\pipe\shiftkeys_automation"
global SHIFTKEYS_IPC_TIMEOUT_MS := 2000
global SHIFTKEYS_IPC_RECONNECT_DELAY_MS := 500

; Internal: pipe handle (0 = not connected)
global _ShiftKeysIPC_Handle := 0

; Windows constants (script-level; static is function-only in AHK v2)
global GENERIC_READ := 0x80000000, GENERIC_WRITE := 0x40000000
global OPEN_EXISTING := 3, FILE_ATTRIBUTE_NORMAL := 0x80
global INVALID_HANDLE_VALUE := -1

; Connect to daemon. Returns true if connected.
ShiftKeysIPC_Connect() {
    if (!USE_DAEMON || !USE_PIPE_IPC)
        return false
    try {
        h := DllCall("kernel32\CreateFileW", "Str", SHIFTKEYS_PIPE_NAME, "UInt", GENERIC_READ | GENERIC_WRITE,
            "UInt", 0, "Ptr", 0, "UInt", OPEN_EXISTING, "UInt", FILE_ATTRIBUTE_NORMAL, "Ptr", 0, "Ptr")
        if (h != INVALID_HANDLE_VALUE && h != 0) {
            global _ShiftKeysIPC_Handle := h
            return true
        }
    }
    return false
}

; Disconnect and clear handle.
ShiftKeysIPC_Disconnect() {
    global _ShiftKeysIPC_Handle
    if (_ShiftKeysIPC_Handle != 0) {
        try
            DllCall("kernel32\CloseHandle", "Ptr", _ShiftKeysIPC_Handle)
        _ShiftKeysIPC_Handle := 0
    }
}

; Send one request and wait for response with timeout. Returns response object or empty object on failure.
; id, op, context, payload (object), deadlineMs (optional).
ShiftKeysIPC_SendRequest(id, op, context := "", payload := "", deadlineMs := 0) {
    global _ShiftKeysIPC_Handle
    if (!USE_DAEMON || !USE_PIPE_IPC || _ShiftKeysIPC_Handle = 0) {
        if (USE_DAEMON && USE_PIPE_IPC && ShiftKeysIPC_Connect())
        ; retry once after connect
            return ShiftKeysIPC_SendRequest(id, op, context, payload, deadlineMs)
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
    jsonReq := _ShiftKeysIPC_JsonEncode(req)
    if (jsonReq = "")
        return Map()
    utf8 := Buffer(StrPut(jsonReq, "UTF-8"))
    len := StrPut(jsonReq, utf8, "UTF-8")
    ; Frame: 4-byte big-endian length + payload
    frameLen := 4 + len
    frame := Buffer(frameLen)
    ; Big-endian length (one byte each)
    NumPut("UChar", (len >> 24) & 0xFF, frame, 0)
    NumPut("UChar", (len >> 16) & 0xFF, frame, 1)
    NumPut("UChar", (len >> 8) & 0xFF, frame, 2)
    NumPut("UChar", len & 0xFF, frame, 3)
    ptr := utf8.Ptr
    loop len
        NumPut("UChar", NumGet(utf8, A_Index - 1, "UChar"), frame, 3 + A_Index)
    ; Write
    timeOut := (deadlineMs > 0) ? deadlineMs : SHIFTKEYS_IPC_TIMEOUT_MS
    needReconnect := false
    try {
        written := 0
        r := DllCall("kernel32\WriteFile", "Ptr", _ShiftKeysIPC_Handle, "Ptr", frame.Ptr, "UInt", frameLen, "UInt*", &
            written, "Ptr", 0)
        if (!r || written != frameLen)
            needReconnect := true
        else {
            ; Read 4-byte length
            lenBuf := Buffer(4)
            read := 0
            r := DllCall("kernel32\ReadFile", "Ptr", _ShiftKeysIPC_Handle, "Ptr", lenBuf.Ptr, "UInt", 4, "UInt*", &
                read, "Ptr", 0)
            if (!r || read != 4)
                needReconnect := true
            else {
                respLen := (NumGet(lenBuf, 0, "UChar") << 24) | (NumGet(lenBuf, 1, "UChar") << 16) | (NumGet(lenBuf,
                    2, "UChar") << 8) | NumGet(lenBuf, 3, "UChar")
                if (respLen > 0 && respLen <= 1048576) {
                    respBuf := Buffer(respLen)
                    read := 0
                    r := DllCall("kernel32\ReadFile", "Ptr", _ShiftKeysIPC_Handle, "Ptr", respBuf.Ptr, "UInt",
                        respLen, "UInt*", &read, "Ptr", 0)
                    if (r && read = respLen) {
                        respStr := StrGet(respBuf.Ptr, respLen, "UTF-8")
                        try {
                            return _ShiftKeysIPC_JsonDecode(respStr)
                        } catch
                            return Map()
                    }
                }
                needReconnect := true
            }
        }
    } catch
        needReconnect := true
    if (needReconnect) {
        ShiftKeysIPC_Disconnect()
        Sleep(SHIFTKEYS_IPC_RECONNECT_DELAY_MS)
        if (ShiftKeysIPC_Connect())
            return ShiftKeysIPC_SendRequest(id, op, context, payload, deadlineMs)
    }
    return Map()
}

; Heartbeat. Returns true if daemon responded ok.
ShiftKeysIPC_HealthCheck() {
    resp := ShiftKeysIPC_SendRequest("HealthCheck", "HealthCheck", "", "", 1000)
    return resp.Has("ok") && resp["ok"] = true
}

; Call once at script load to bootstrap connection (non-blocking: try connect, no wait).
ShiftKeysIPC_Bootstrap() {
    if (!USE_DAEMON || !USE_PIPE_IPC)
        return
    ShiftKeysIPC_Connect()
    ; Mirror daemon context cache periodically for O(1) #HotIf (Phase 2)
    SetTimer(ShiftKeysIPC_MirrorContext, 200)
}

; --- Phase 2: O(1) context cache mirror (WinEvent-driven in daemon) ---
global g_ShiftKeys_IsChatGPTActive := false
global g_ShiftKeys_IsGeminiActive := false
global g_ShiftKeys_IsPowerBIActive := false
global g_ShiftKeys_IsOutlookMainActive := false
global g_ShiftKeys_IsOutlookMessageActive := false
global g_ShiftKeys_IsOutlookAppointmentActive := false
global g_ShiftKeys_IsWikipediaActive := false
global g_ShiftKeys_IsChromePdfViewerActive := false

; Feature flag: use daemon-cached context for ChatGPT #HotIf (default false = legacy).
global USE_DAEMON_CONTEXT_CHATGPT := false

; Timer callback: fetch ResolveContext and mirror to globals.
ShiftKeysIPC_MirrorContext() {
    if (!USE_DAEMON || !USE_PIPE_IPC)
        return
    resp := ShiftKeysIPC_SendRequest("ResolveContext", "ResolveContext", "", "", 500)
    if (!resp.Has("ok") || resp["ok"] != true || !resp.Has("result"))
        return
    r := resp["result"]
    global g_ShiftKeys_IsChatGPTActive := r.Has("IsChatGPTActive") ? !!r["IsChatGPTActive"] : false
    global g_ShiftKeys_IsGeminiActive := r.Has("IsGeminiActive") ? !!r["IsGeminiActive"] : false
    global g_ShiftKeys_IsPowerBIActive := r.Has("IsPowerBIActive") ? !!r["IsPowerBIActive"] : false
    global g_ShiftKeys_IsOutlookMainActive := r.Has("IsOutlookMainActive") ? !!r["IsOutlookMainActive"] : false
    global g_ShiftKeys_IsOutlookMessageActive := r.Has("IsOutlookMessageActive") ? !!r["IsOutlookMessageActive"] :
        false
    global g_ShiftKeys_IsOutlookAppointmentActive := r.Has("IsOutlookAppointmentActive") ? !!r[
        "IsOutlookAppointmentActive"] : false
    global g_ShiftKeys_IsWikipediaActive := r.Has("IsWikipediaActive") ? !!r["IsWikipediaActive"] : false
    global g_ShiftKeys_IsChromePdfViewerActive := r.Has("IsChromePdfViewerActive") ? !!r["IsChromePdfViewerActive"] :
        false
}

; Returns true if ChatGPT context is active (daemon cache or legacy). Use in #HotIf.
IsChatGPTActiveForHotkey() {
    if (USE_DAEMON_CONTEXT_CHATGPT) {
        global g_ShiftKeys_IsChatGPTActive
        return g_ShiftKeys_IsChatGPTActive
    }
    hwnd := GetChatGPTWindowHwnd()
    return hwnd && WinActive("ahk_id " hwnd)
}

; --- Phase 3: Non-blocking Gemini monitor (daemon WatchUIState + poll callback) ---
; Set to true to use daemon watch instead of blocking WaitForStopResponseButton_Gemini.
global USE_DAEMON_MONITOR_GEMINI := false
global _ShiftKeysIPC_GeminiWatchId := ""
global _ShiftKeysIPC_GeminiWatchTimer := ""

; Start async watch for Gemini "Stop response" completion; callback runs when done/timeout.
ShiftKeysIPC_StartGeminiWatch(timeoutMs := 300000, onComplete := "") {
    if (!USE_DAEMON || !USE_PIPE_IPC)
        return
    payload := Map("context", "Gemini", "timeoutMs", timeoutMs)
    resp := ShiftKeysIPC_SendRequest("WatchGemini", "WatchUIState", "", payload, 5000)
    if (!resp.Has("ok") || resp["ok"] != true || !resp.Has("result"))
        return
    wid := resp["result"].Has("watchId") ? resp["result"]["watchId"] : ""
    if (wid = "")
        return
    global _ShiftKeysIPC_GeminiWatchId := wid
    global _ShiftKeysIPC_GeminiWatchTimer := SetTimer(ShiftKeysIPC_PollGeminiWatch.Bind(onComplete), 500)
}

; Timer callback: poll GetWatchStatus; when completed/timeout/error, run callback and stop timer.
ShiftKeysIPC_PollGeminiWatch(onComplete := "") {
    global _ShiftKeysIPC_GeminiWatchId, _ShiftKeysIPC_GeminiWatchTimer
    if (_ShiftKeysIPC_GeminiWatchId = "")
        return
    payload := Map("watchId", _ShiftKeysIPC_GeminiWatchId)
    resp := ShiftKeysIPC_SendRequest("GetWatchStatus", "GetWatchStatus", "", payload, 1000)
    if (!resp.Has("ok") || resp["ok"] != true || !resp.Has("result"))
        return
    status := resp["result"].Has("status") ? resp["result"]["status"] : ""
    if (status = "pending")
        return
    try SetTimer(_ShiftKeysIPC_GeminiWatchTimer, 0)
    _ShiftKeysIPC_GeminiWatchTimer := ""
    _ShiftKeysIPC_GeminiWatchId := ""
    if (IsSet(onComplete) && onComplete != "")
        try (IsObject(onComplete) ? onComplete.Call() : %onComplete%())
}
