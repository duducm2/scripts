; AppLauncher IPC client: MMF + Mutex + EventWaitHandle sync with applauncher_daemon.
; Feature flags: AL_USE_DAEMON, AL_USE_MMF_IPC (both default off).
; DllCall wrappers for OpenFileMappingW, MapViewOfFile, Mutex, Events.

; --- Constants (Windows API) ---
global AL_IPC_MMF_NAME := "Global\AppLauncherIPC"
global AL_IPC_MUTEX_NAME := "Global\AppLauncherIPCLock"
global AL_IPC_EVENT_REQUEST := "Global\AppLauncherRequestReady"
global AL_IPC_EVENT_RESPONSE := "Global\AppLauncherResponseReady"
global AL_IPC_MMF_SIZE := 4096
global AL_IPC_HEADER_SIZE := 20
global AL_IPC_PAYLOAD_MAX := AL_IPC_MMF_SIZE - AL_IPC_HEADER_SIZE

; FILE_MAP_ALL_ACCESS, PAGE_READWRITE, handle constants
global AL_IPC_FILE_MAP_ALL_ACCESS := 0xF001F
global AL_IPC_PAGE_READWRITE := 0x04
global AL_IPC_INVALID_HANDLE := 0
global AL_IPC_WAIT_OBJECT_0 := 0
global AL_IPC_WAIT_TIMEOUT := 258
global AL_IPC_WAIT_ABANDONED := 0x80
global AL_IPC_EVENT_MODIFY_STATE := 0x0002
global AL_IPC_MUTEX_MODIFY_STATE := 0x0001
global AL_IPC_SYNCHRONIZE := 0x00100000
global AL_IPC_STATUS_IDLE := 0
global AL_IPC_STATUS_REQUEST_READY := 1
global AL_IPC_STATUS_RESPONSE_READY := 2
global AL_IPC_VERSION := 1

; Handles (set by IpcInit, cleared by IpcTeardown)
global AL_IPC_hMap := 0
global AL_IPC_pView := 0
global AL_IPC_hMutex := 0
global AL_IPC_hEvtRequest := 0
global AL_IPC_hEvtResponse := 0
global AL_IPC_seq := 0

; --- DllCall wrappers ---
AL_IPC_OpenFileMapping() {
    h := DllCall("kernel32\OpenFileMappingW", "UInt", AL_IPC_FILE_MAP_ALL_ACCESS, "Int", 0, "Str", AL_IPC_MMF_NAME,
        "Ptr")
    return h
}

AL_IPC_MapViewOfFile(hMap) {
    p := DllCall("kernel32\MapViewOfFile", "Ptr", hMap, "UInt", AL_IPC_FILE_MAP_ALL_ACCESS, "UInt", 0, "UInt", 0,
        "UPtr", AL_IPC_MMF_SIZE, "Ptr")
    return p
}

AL_IPC_UnmapViewOfFile(pView) {
    return DllCall("kernel32\UnmapViewOfFile", "Ptr", pView, "Int")
}

AL_IPC_CloseHandle(h) {
    return DllCall("kernel32\CloseHandle", "Ptr", h, "Int")
}

AL_IPC_OpenMutex() {
    h := DllCall("kernel32\OpenMutexW", "UInt", AL_IPC_MUTEX_MODIFY_STATE | AL_IPC_SYNCHRONIZE, "Int", 0, "Str",
        AL_IPC_MUTEX_NAME, "Ptr")
    return h
}

AL_IPC_OpenEventRequest() {
    h := DllCall("kernel32\OpenEventW", "UInt", AL_IPC_EVENT_MODIFY_STATE | AL_IPC_SYNCHRONIZE, "Int", 0, "Str",
        AL_IPC_EVENT_REQUEST, "Ptr")
    return h
}

AL_IPC_OpenEventResponse() {
    ; Need EVENT_MODIFY_STATE so AHK can ResetEvent after reading response
    h := DllCall("kernel32\OpenEventW", "UInt", AL_IPC_EVENT_MODIFY_STATE | AL_IPC_SYNCHRONIZE, "Int", 0, "Str",
        AL_IPC_EVENT_RESPONSE, "Ptr")
    return h
}

AL_IPC_WaitForSingleObject(h, timeoutMs) {
    r := DllCall("kernel32\WaitForSingleObject", "Ptr", h, "UInt", timeoutMs, "UInt")
    return r
}

AL_IPC_SetEvent(h) {
    return DllCall("kernel32\SetEvent", "Ptr", h, "Int")
}

AL_IPC_ResetEvent(h) {
    return DllCall("kernel32\ResetEvent", "Ptr", h, "Int")
}

AL_IPC_ReleaseMutex(h) {
    return DllCall("kernel32\ReleaseMutex", "Ptr", h, "Int")
}

; --- Pack header: version(4), seq(4), payloadLen(4), status(4), errorCode(4) ---
AL_IPC_PackHeader(version, seq, payloadLen, status, errorCode) {
    buf := Buffer(AL_IPC_HEADER_SIZE, 0)
    NumPut("UInt", version, buf, 0)
    NumPut("UInt", seq, buf, 4)
    NumPut("UInt", payloadLen, buf, 8)
    NumPut("UInt", status, buf, 12)
    NumPut("UInt", errorCode, buf, 16)
    return buf
}

; --- Write bytes to MMF at offset (buf = Buffer) ---
AL_IPC_WriteToView(pView, offset, buf) {
    len := buf.Size
    if (offset + len > AL_IPC_MMF_SIZE)
        len := AL_IPC_MMF_SIZE - offset
    DllCall("RtlMoveMemory", "Ptr", pView + offset, "Ptr", buf.Ptr, "UPtr", len)
}

; --- Read bytes from MMF ---
AL_IPC_ReadFromView(pView, offset, length) {
    buf := Buffer(length, 0)
    DllCall("RtlMoveMemory", "Ptr", buf.Ptr, "Ptr", pView + offset, "UPtr", length)
    return buf
}

; --- Minimal JSON encode (request only: id, op, context, payload, ts, deadlineMs) ---
AL_IPC_JsonEncodeRequest(id, op, context, payload, deadlineMs) {
    ts := Integer(A_TickCount)
    s := "{`"id`":`"" . id . "`",`"op`":`"" . op . "`",`"context`":`"" . context . "`",`"payload`":" . (Type(payload) =
    "Map" ? AL_IPC_MapToJson(payload) : "{}") . ",`"ts`":" . ts . ",`"deadlineMs`":" . deadlineMs . "}"
    byteLen := StrPut(s, "UTF-8")
    buf := Buffer(byteLen, 0)
    StrPut(s, buf, "UTF-8")
    return buf
}

AL_IPC_MapToJson(m) {
    out := ""
    for k, v in m {
        if (out != "")
            out .= ","
        out .= "`"" . StrReplace(StrReplace(StrReplace(String(k), "\", "\\"), "`"", "\`""), "`n", "\n") . "`":"
        if (v is Integer || v is Float)
            out .= String(v)
        else if (v is String)
            out .= "`"" . StrReplace(StrReplace(StrReplace(String(v), "\", "\\"), "`"", "\`""), "`n", "\n") . "`""
        else
            out .= "{}"
    }
    return "{" . out . "}"
}

; --- Decode response JSON (simplified: extract id, ok, result) ---
AL_IPC_DecodeResponse(utf8Buf) {
    try {
        str := StrGet(utf8Buf, "UTF-8")
        o := Map()
        ; Minimal parse: "ok":true/false, "result":{...}, "id":"..."
        if InStr(str, "`"ok`":true")
            o["ok"] := true
        else
            o["ok"] := false
        if (pos := InStr(str, "`"result`":")) {
            start := pos + 10
            depth := 0
            i := start
            loop {
                c := SubStr(str, i, 1)
                if (c = "{")
                    depth++
                else if (c = "}")
                    depth--
                if (depth = -1) {
                    o["resultRaw"] := SubStr(str, start, i - start)
                    break
                }
                i++
                if (i > StrLen(str))
                    break
            }
        }
        return o
    }
    return Map("ok", false, "resultRaw", "")
}

; --- IpcInit: open MMF, map view, open mutex and events ---
AL_IPC_Init() {
    if (AL_IPC_pView)
        return true
    hMap := AL_IPC_OpenFileMapping()
    if (!hMap || hMap = AL_IPC_INVALID_HANDLE)
        return false
    pView := AL_IPC_MapViewOfFile(hMap)
    if (!pView) {
        AL_IPC_CloseHandle(hMap)
        return false
    }
    hMutex := AL_IPC_OpenMutex()
    if (!hMutex) {
        AL_IPC_UnmapViewOfFile(pView)
        AL_IPC_CloseHandle(hMap)
        return false
    }
    hReq := AL_IPC_OpenEventRequest()
    hResp := AL_IPC_OpenEventResponse()
    if (!hReq || !hResp) {
        if (hReq) AL_IPC_CloseHandle(hReq)
            if (hResp) AL_IPC_CloseHandle(hResp)
                AL_IPC_CloseHandle(hMutex)
        AL_IPC_UnmapViewOfFile(pView)
        AL_IPC_CloseHandle(hMap)
        return false
    }
    global AL_IPC_hMap := hMap
    global AL_IPC_pView := pView
    global AL_IPC_hMutex := hMutex
    global AL_IPC_hEvtRequest := hReq
    global AL_IPC_hEvtResponse := hResp
    return true
}

; --- IpcTeardown ---
AL_IPC_Teardown() {
    if (AL_IPC_hEvtResponse) {
        AL_IPC_CloseHandle(AL_IPC_hEvtResponse)
        global AL_IPC_hEvtResponse := 0
    }
    if (AL_IPC_hEvtRequest) {
        AL_IPC_CloseHandle(AL_IPC_hEvtRequest)
        global AL_IPC_hEvtRequest := 0
    }
    if (AL_IPC_hMutex) {
        AL_IPC_CloseHandle(AL_IPC_hMutex)
        global AL_IPC_hMutex := 0
    }
    if (AL_IPC_pView) {
        AL_IPC_UnmapViewOfFile(AL_IPC_pView)
        global AL_IPC_pView := 0
    }
    if (AL_IPC_hMap) {
        AL_IPC_CloseHandle(AL_IPC_hMap)
        global AL_IPC_hMap := 0
    }
}

; --- IpcSend: write request to MMF, signal request event ---
AL_IPC_Send(op, payload := Map(), deadlineMs := 5000) {
    if (!AL_IPC_pView || !AL_IPC_hMutex || !AL_IPC_hEvtRequest)
        return ""
    id := A_NowUTC . A_MSec . Random(1, 99999)
    context := ""
    buf := AL_IPC_JsonEncodeRequest(id, op, context, payload, deadlineMs)
    payloadLen := buf.Size
    if (payloadLen > AL_IPC_PAYLOAD_MAX)
        return ""
    AL_IPC_seq++
    header := AL_IPC_PackHeader(AL_IPC_VERSION, AL_IPC_seq, payloadLen, AL_IPC_STATUS_REQUEST_READY, 0)
    r := AL_IPC_WaitForSingleObject(AL_IPC_hMutex, 5000)
    if (r != AL_IPC_WAIT_OBJECT_0 && r != AL_IPC_WAIT_ABANDONED)
        return ""
    try {
        AL_IPC_WriteToView(AL_IPC_pView, 0, header)
        AL_IPC_WriteToView(AL_IPC_pView, AL_IPC_HEADER_SIZE, buf)
        AL_IPC_SetEvent(AL_IPC_hEvtRequest)
    } finally {
        AL_IPC_ReleaseMutex(AL_IPC_hMutex)
    }
    return id
}

; --- IpcAwait: wait for response event, read response from MMF ---
AL_IPC_Await(id, timeoutMs := 5000) {
    if (!AL_IPC_hEvtResponse || !AL_IPC_hMutex || !AL_IPC_pView)
        return Map("ok", false)
    r := AL_IPC_WaitForSingleObject(AL_IPC_hEvtResponse, timeoutMs)
    if (r != AL_IPC_WAIT_OBJECT_0 && r != AL_IPC_WAIT_ABANDONED)
        return Map("ok", false)
    if (AL_IPC_WaitForSingleObject(AL_IPC_hMutex, 5000) != AL_IPC_WAIT_OBJECT_0)
        return Map("ok", false)
    try {
        headerBuf := AL_IPC_ReadFromView(AL_IPC_pView, 0, AL_IPC_HEADER_SIZE)
        status := NumGet(headerBuf, 12, "UInt")
        payloadLen := NumGet(headerBuf, 8, "UInt")
        if (status != AL_IPC_STATUS_RESPONSE_READY || payloadLen = 0 || payloadLen > AL_IPC_PAYLOAD_MAX) {
            AL_IPC_ResetEvent(AL_IPC_hEvtResponse)
            return Map("ok", false)
        }
        payloadBuf := AL_IPC_ReadFromView(AL_IPC_pView, AL_IPC_HEADER_SIZE, payloadLen)
        ; Mark consumed (status = IDLE)
        statusBuf := Buffer(4, 0)
        NumPut("UInt", AL_IPC_STATUS_IDLE, statusBuf)
        AL_IPC_WriteToView(AL_IPC_pView, 12, statusBuf)
        resp := AL_IPC_DecodeResponse(payloadBuf)
    } finally {
        AL_IPC_ResetEvent(AL_IPC_hEvtResponse)
        AL_IPC_ReleaseMutex(AL_IPC_hMutex)
    }
    return resp
}

; --- Extract integer from result JSON string (e.g. "primaryHwnd":123) ---
AL_IPC_GetResultInt(resp, key) {
    if (!resp.Has("resultRaw"))
        return 0
    raw := resp["resultRaw"]
    needle := "`"" . key . "`":"
    pos := InStr(raw, needle)
    if (!pos)
        return 0
    start := pos + StrLen(needle)
    n := 0
    while (start <= StrLen(raw)) {
        c := SubStr(raw, start, 1)
        if (c >= "0" && c <= "9")
            n := n * 10 + Integer(c)
        else
            return n
        start++
    }
    return n
}

; --- One-shot request/response (Send + Await) ---
AL_IPC_Call(op, payload := Map(), timeoutMs := 5000) {
    if (!AL_USE_DAEMON || !AL_USE_MMF_IPC)
        return Map("ok", false)
    if (!AL_IPC_Init())
        return Map("ok", false)
    id := AL_IPC_Send(op, payload, timeoutMs)
    if (id = "")
        return Map("ok", false)
    return AL_IPC_Await(id, timeoutMs)
}
