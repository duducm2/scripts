; StudyLink helpers for per-study subtopic links via Google Apps Script API
; API endpoint: Google Apps Script web app (GET = read, POST = write)
global STUDY_LINKS_API_URL :=
    "https://script.google.com/macros/s/AKfycbzzkjpT_47W0TwcjwEulzkV9l5xTtqcwWJmF0h-B-11SwiL_49SPhKXnj3PTsgFUZcp/exec"

; ────────────────────────────────────────────────────────────────────────────
; HTTP helpers
; ServerXMLHTTP + WinHttp use WinHTTP (failed here with 0x80072EE7 DNS).
; XMLHTTP uses WinInet (browser/system proxy). PowerShell IWR is fallback.
; ────────────────────────────────────────────────────────────────────────────
global StudyLink_LastHttpMeta := Map()

StudyLink_HttpLogMeta(backend, status, elapsedMs, err := "") {
    global StudyLink_LastHttpMeta
    StudyLink_LastHttpMeta := Map("backend", backend, "status", status, "ms", elapsedMs, "err", err)
    ; #region agent log
    errEsc := StrReplace(SubStr(err, 1, 120), '"', "'")
    try FileAppend('{"sessionId":"d22874","timestamp":' . A_TickCount .
        ',"hypothesisId":"H7","location":"StudyLink_HttpSend","message":"http_done","data":{"backend":"' .
        backend . '","status":' . status . ',"ms":' . elapsedMs . ',"err":"' . errEsc . '"}}' "`n",
        A_ScriptDir "\debug-d22874.log", "UTF-8")
    ; #endregion
}

StudyLink_HttpSendXmlHttp(method, url, body := "") {
    xhr := ComObject("MSXML2.XMLHTTP.6.0")
    xhr.open(method, url, false)
    xhr.setRequestHeader("User-Agent", "AutoHotkey/2.0")
    if (method = "POST")
        xhr.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
    if (body != "")
        xhr.send(body)
    else
        xhr.send()
    return { status: xhr.status, text: xhr.responseText }
}

StudyLink_HttpSendPowerShell(method, url, body := "") {
    tmpBody := ""
    if (method = "POST" && body != "") {
        tmpBody := A_Temp "\studylink_body_" A_TickCount ".txt"
        try FileDelete(tmpBody)
        FileAppend(body, tmpBody, "UTF-8")
    }
    tmpOut := A_Temp "\studylink_resp_" A_TickCount ".txt"
    try FileDelete(tmpOut)
    if (method = "POST") {
        ps := "try { "
            . "$u='" . StrReplace(url, "'", "''") . "'; "
            . (tmpBody != "" ? "$b=Get-Content -LiteralPath '" . StrReplace(tmpBody, "'", "''") .
                "' -Raw; $r=Invoke-WebRequest -Uri $u -Method POST -Body $b -ContentType 'application/x-www-form-urlencoded' -UseBasicParsing; "
            : "$r=Invoke-WebRequest -Uri $u -Method POST -UseBasicParsing; ")
            . "$r.Content | Out-File -LiteralPath '" . StrReplace(tmpOut, "'", "''") .
            "' -Encoding utf8 -NoNewline } catch { $_.Exception.Message | Out-File -LiteralPath '" .
            StrReplace(tmpOut, "'", "''") . "' -Encoding utf8 }"
    } else {
        ps := "try { (Invoke-WebRequest -Uri '" . StrReplace(url, "'", "''") .
            "' -UseBasicParsing).Content | Out-File -LiteralPath '" . StrReplace(tmpOut, "'", "''") .
            "' -Encoding utf8 -NoNewline } catch { $_.Exception.Message | Out-File -LiteralPath '" .
            StrReplace(tmpOut, "'", "''") . "' -Encoding utf8 }"
    }
    RunWait('powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "' . ps . '"', , "Hide")
    if !FileExist(tmpOut)
        return { status: 0, text: "ERROR: PowerShell produced no output" }
    resp := FileRead(tmpOut, "UTF-8")
    try FileDelete(tmpOut)
    if (tmpBody != "")
        try FileDelete(tmpBody)
    if (SubStr(resp, 1, 12) = "Unable to co" || InStr(resp, "could not be resolved"))
        return { status: 0, text: "ERROR: " . resp }
    return { status: 200, text: resp }
}

StudyLink_HttpSend(method, url, body := "") {
    t0 := A_TickCount
    try {
        r := StudyLink_HttpSendXmlHttp(method, url, body)
        elapsed := A_TickCount - t0
        if (r.status >= 200 && r.status < 300) {
            StudyLink_HttpLogMeta("XMLHTTP", r.status, elapsed)
            return r.text
        }
        if (r.text != "") {
            StudyLink_HttpLogMeta("XMLHTTP", r.status, elapsed, "status " . r.status)
            return r.text
        }
        throw Error("XMLHTTP status " . r.status)
    } catch as e1 {
        err1 := e1.Message
        elapsed1 := A_TickCount - t0
        StudyLink_HttpLogMeta("XMLHTTP", 0, elapsed1, err1)
        if !(InStr(err1, "80072EE7") || InStr(err1, "could not be resolved") || InStr(err1, "status 0"))
            return "ERROR: " . err1
    }
    t1 := A_TickCount
    try {
        r2 := StudyLink_HttpSendPowerShell(method, url, body)
        elapsed2 := A_TickCount - t1
        StudyLink_HttpLogMeta("PowerShell", r2.status, elapsed2, SubStr(r2.text, 1, 6) = "ERROR:" ? r2.text : "")
        return r2.text
    } catch as e2 {
        StudyLink_HttpLogMeta("PowerShell", 0, A_TickCount - t1, e2.Message)
        return "ERROR: " . e2.Message
    }
}

StudyLink_HttpGet(params) {
    fullUrl := STUDY_LINKS_API_URL
    if (params != "")
        fullUrl .= "?" . params
    return StudyLink_HttpSend("GET", fullUrl)
}

StudyLink_HttpPost(data) {
    return StudyLink_HttpSend("POST", STUDY_LINKS_API_URL, data)
}

; ────────────────────────────────────────────────────────────────────────────
; URL-encode / URL-decode helpers
; ────────────────────────────────────────────────────────────────────────────
StudyLink_UrlEncode(str) {
    s := StrReplace(str, " ", "%20")
    s := StrReplace(s, "&", "%26")
    s := StrReplace(s, "=", "%3D")
    s := StrReplace(s, "?", "%3F")
    s := StrReplace(s, "#", "%23")
    s := StrReplace(s, "+", "%2B")
    s := StrReplace(s, "/", "%2F")
    s := StrReplace(s, ":", "%3A")
    s := StrReplace(s, '"', "%22")
    s := StrReplace(s, "'", "%27")
    return s
}

StudyLink_UrlDecode(str) {
    s := StrReplace(str, "%20", " ")
    s := StrReplace(s, "%21", "!")
    s := StrReplace(s, "%22", '"')
    s := StrReplace(s, "%23", "#")
    s := StrReplace(s, "%24", "$")
    s := StrReplace(s, "%25", "%")
    s := StrReplace(s, "%26", "&")
    s := StrReplace(s, "%27", "'")
    s := StrReplace(s, "%28", "(")
    s := StrReplace(s, "%29", ")")
    s := StrReplace(s, "%2A", "*")
    s := StrReplace(s, "%2B", "+")
    s := StrReplace(s, "%2C", ",")
    s := StrReplace(s, "%2D", "-")
    s := StrReplace(s, "%2E", ".")
    s := StrReplace(s, "%2F", "/")
    s := StrReplace(s, "%3A", ":")
    s := StrReplace(s, "%3B", ";")
    s := StrReplace(s, "%3D", "=")
    s := StrReplace(s, "%3E", ">")
    s := StrReplace(s, "%3F", "?")
    s := StrReplace(s, "%40", "@")
    s := StrReplace(s, "%5B", "[")
    s := StrReplace(s, "%5C", "\")
    s := StrReplace(s, "%5D", "]")
    s := StrReplace(s, "%5E", "^")
    s := StrReplace(s, "%5F", "_")
    s := StrReplace(s, "%60", Chr(96))
    s := StrReplace(s, "%7B", "{")
    s := StrReplace(s, "%7C", "|")
    s := StrReplace(s, "%7D", "}")
    s := StrReplace(s, "%7E", "~")
    return s
}

; ────────────────────────────────────────────────────────────────────────────
; Parse the API response into a Map.
; WARNING: The `url` value may contain literal `&` characters (e.g. &t=435s).
; StrSplit("&") would truncate those — instead we extract `url=` as the LAST
; field by taking everything after its prefix (url is always the final field).
; Simple scalar fields (key, status) use safe &-delimited extraction.
; ────────────────────────────────────────────────────────────────────────────
StudyLink_ParseFormEncoded(response) {
    result := Map()
    if RegExMatch(response, "key=([^&]+)", &m)
        result["key"] := m[1]
    if RegExMatch(response, "status=([^&]+)", &m)
        result["status"] := m[1]
    urlPos := InStr(response, "url=")
    if (urlPos) {
        rawUrl := SubStr(response, urlPos + 4)
        result["url"] := rawUrl
    }
    return result
}

; ────────────────────────────────────────────────────────────────────────────
; Public API: get a stored link for a study key
; API returns URL-encoded form data: key=subtopic&url=https%3A%2F%2F...
; Returns the decoded URL string, or "" if not found / on error.
; ────────────────────────────────────────────────────────────────────────────
StudyLink_Get(studyKey) {
    response := StudyLink_HttpGet("key=" . StudyLink_UrlEncode(studyKey))
    if (SubStr(response, 1, 6) = "ERROR:") {
        return ""
    }
    parsed := StudyLink_ParseFormEncoded(response)
    if (parsed.Has("url")) {
        rawUrl := parsed["url"]
        if (rawUrl != "")
            return StudyLink_UrlDecode(rawUrl)
    }
    return ""
}

; ────────────────────────────────────────────────────────────────────────────
; Public API: set a link for a study key
; Sends POST with key and url; accepts status=ok form data or plain-text "Saved"
; Returns true on success, false otherwise.
; ────────────────────────────────────────────────────────────────────────────
StudyLink_Set(studyKey, url) {
    ; Match MacroDroid Set_Video.macro: literal url in body (Apps Script reads everything after url=).
    data := "key=" . StudyLink_UrlEncode(studyKey) . "&url=" . url
    response := StudyLink_HttpPost(data)
    ; #region agent log
    respPreview := SubStr(StrReplace(response, '"', "'"), 1, 80)
    global StudyLink_LastHttpMeta
    httpBackend := StudyLink_LastHttpMeta.Has("backend") ? StudyLink_LastHttpMeta["backend"] : "?"
    httpStatus := StudyLink_LastHttpMeta.Has("status") ? StudyLink_LastHttpMeta["status"] : 0
    httpMs := StudyLink_LastHttpMeta.Has("ms") ? StudyLink_LastHttpMeta["ms"] : 0
    try FileAppend('{"sessionId":"d22874","timestamp":' . A_TickCount .
        ',"hypothesisId":"H4","location":"StudyLink_Set","message":"api_response","data":{"backend":"' .
        httpBackend . '","httpStatus":' . httpStatus . ',"httpMs":' . httpMs . ',"respLen":' .
        StrLen(response) . ',"preview":"' . respPreview . '","bodyLen":' . StrLen(data) . '}}' "`n",
        A_ScriptDir "\debug-d22874.log", "UTF-8")
    ; #endregion
    if (SubStr(response, 1, 6) = "ERROR:") {
        return false
    }
    parsed := StudyLink_ParseFormEncoded(response)
    if (parsed.Has("status") && (parsed["status"] = "ok" || parsed["status"] = "OK"))
        return true
    resp := Trim(response)
    if (resp = "Saved" || resp = "saved" || resp = "OK" || resp = "ok")
        return true
    if (resp != "" && InStr(resp, "Saved", false))
        return true
    if InStr(response, "ok") || InStr(response, "OK")
        return true
    return false
}

; Legacy INI path (kept for reference — not used by new API code)
StudyLink_IniPath() {
    return A_ScriptDir "\data\study_links.ini"
}

; Functional test: verifies both GET and SET operations against the live API.
StudyLink_RunFunctionalTest() {
    testKey := "subtopic_test"
    testUrl := "https://example.com/api-test-link-" . A_TickCount
    result := Map()

    setOk := StudyLink_Set(testKey, testUrl)
    result["setOk"] := setOk
    result["setMsg"] := setOk ? "PASSED" : "FAILED"

    retrieved := StudyLink_Get(testKey)
    getOk := (retrieved = testUrl)
    result["getOk"] := getOk
    result["getMsg"] := getOk ? "PASSED" : "FAILED (got: '" . retrieved . "')"

    StudyLink_Set(testKey, "")

    return result
}

StudyLink_Open(studyKey) {
    url := StudyLink_Get(studyKey)
    if (url != "") {
        try Run(url)
    } else {
        MsgBox "No link stored for this study."
    }
}
