; StudyLink helpers for per-study subtopic links via Google Apps Script API
; API endpoint: Google Apps Script web app (GET = read, POST = write)
global STUDY_LINKS_API_URL :=
    "https://script.google.com/macros/s/AKfycbzzkjpT_47W0TwcjwEulzkV9l5xTtqcwWJmF0h-B-11SwiL_49SPhKXnj3PTsgFUZcp/exec"

; ────────────────────────────────────────────────────────────────────────────
; HTTP helpers
; ServerXMLHTTP + WinHttp use WinHTTP (failed here with 0x80072EE7 DNS).
; XMLHTTP uses WinInet (browser/system proxy). PowerShell IWR is fallback.
; ────────────────────────────────────────────────────────────────────────────
StudyLink_HttpSendXmlHttp(method, url, body := "", useCustomUserAgent := false) {
    xhr := ComObject("MSXML2.XMLHTTP.6.0")
    xhr.open(method, url, false)
    if (useCustomUserAgent)
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

StudyLink_IsHttpErrorText(response) {
    return (response = "" || SubStr(response, 1, 6) = "ERROR:")
}

StudyLink_IsValidGetResponse(response) {
    if StudyLink_IsHttpErrorText(response)
        return false
    if (InStr(response, "<html") || InStr(response, "<!DOCTYPE"))
        return false
    return (InStr(response, "key=") || InStr(response, "url="))
}

StudyLink_HttpSend(method, url, body := "") {
    try {
        r := StudyLink_HttpSendXmlHttp(method, url, body, false)
        if (r.status >= 200 && r.status < 300 && r.text != "")
            return r.text
    } catch {
    }
    try {
        r := StudyLink_HttpSendPowerShell(method, url, body)
        if (r.text != "")
            return r.text
        return "ERROR: PowerShell produced no output"
    } catch as e2 {
        return "ERROR: " . e2.Message
    }
}

StudyLink_HttpSendGet(url) {
    try {
        r := StudyLink_HttpSendXmlHttp("GET", url, "", false)
        if (r.status >= 200 && r.status < 300 && StudyLink_IsValidGetResponse(r.text))
            return r.text
    } catch {
    }
    try {
        r := StudyLink_HttpSendPowerShell("GET", url)
        if (StudyLink_IsValidGetResponse(r.text))
            return r.text
        if (r.text != "" && !StudyLink_IsHttpErrorText(r.text))
            return r.text
        if (StudyLink_IsHttpErrorText(r.text))
            return r.text
    } catch as e {
        return "ERROR: " . e.Message
    }
    return "ERROR: Could not read link from API"
}

StudyLink_HttpGet(params) {
    fullUrl := STUDY_LINKS_API_URL
    if (params != "")
        fullUrl .= "?" . params
    return StudyLink_HttpSendGet(fullUrl)
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

StudyLink_ExtractUrlFromResponse(response) {
    parsed := StudyLink_ParseFormEncoded(response)
    if (parsed.Has("url")) {
        rawUrl := parsed["url"]
        if (rawUrl != "")
            return StudyLink_UrlDecode(rawUrl)
    }
    return ""
}

StudyLink_NormalizeApiError(response) {
    if (SubStr(response, 1, 6) = "ERROR:")
        return Trim(SubStr(response, 7))
    if (response = "")
        return "Empty API response"
    if (InStr(response, "<html") || InStr(response, "<!DOCTYPE"))
        return "Unexpected HTML response from API"
    return "Could not read link from API"
}

; Returns Map: ok (true=HTTP+parse succeeded), url (decoded link or ""), err (failure message).
StudyLink_GetResult(studyKey) {
    response := StudyLink_HttpGet("key=" . StudyLink_UrlEncode(studyKey))
    if StudyLink_IsHttpErrorText(response) || !StudyLink_IsValidGetResponse(response) {
        return Map("ok", false, "url", "", "err", StudyLink_NormalizeApiError(response))
    }
    url := StudyLink_ExtractUrlFromResponse(response)
    return Map("ok", true, "url", url, "err", "")
}

; Backward-compatible: returns URL string, or "" if not set or on error.
StudyLink_Get(studyKey) {
    r := StudyLink_GetResult(studyKey)
    return (r["ok"] && r["url"] != "") ? r["url"] : ""
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

    getResult := StudyLink_GetResult(testKey)
    getOk := (getResult["ok"] && getResult["url"] = testUrl)
    result["getOk"] := getOk
    if getOk {
        result["getMsg"] := "PASSED"
    } else if !getResult["ok"] {
        result["getMsg"] := "FAILED (" . getResult["err"] . ")"
    } else {
        result["getMsg"] := "FAILED (got: '" . getResult["url"] . "')"
    }

    StudyLink_Set(testKey, "")

    return result
}

StudyLink_Open(studyKey) {
    r := StudyLink_GetResult(studyKey)
    if (r["ok"] && r["url"] != "") {
        try Run(r["url"])
    } else if !r["ok"] {
        MsgBox "Could not load link: " . r["err"]
    } else {
        MsgBox "No link stored for this study."
    }
}
