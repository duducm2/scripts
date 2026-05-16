; StudyLink helpers for per-study subtopic links via Google Apps Script API
; API endpoint: Google Apps Script web app (GET = read, POST = write)
global STUDY_LINKS_API_URL :=
    "https://script.google.com/macros/s/AKfycbzzkjpT_47W0TwcjwEulzkV9l5xTtqcwWJmF0h-B-11SwiL_49SPhKXnj3PTsgFUZcp/exec"

; ────────────────────────────────────────────────────────────────────────────
; HTTP helpers (WinHttpRequest)
; ────────────────────────────────────────────────────────────────────────────
StudyLink_HttpGet(params) {
    try {
        whr := ComObject("WinHttp.WinHttpRequest.5.1")
        fullUrl := STUDY_LINKS_API_URL
        if (params != "") {
            fullUrl .= "?" . params
        }
        whr.Open("GET", fullUrl, false)
        whr.SetRequestHeader("User-Agent", "AutoHotkey/2.0")
        whr.Send()
        return whr.ResponseText
    } catch as e {
        return "ERROR: " . e.Message
    }
}

StudyLink_HttpPost(data) {
    try {
        whr := ComObject("WinHttp.WinHttpRequest.5.1")
        whr.Open("POST", STUDY_LINKS_API_URL, false)
        whr.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded")
        whr.SetRequestHeader("User-Agent", "AutoHotkey/2.0")
        whr.Send(data)
        return whr.ResponseText
    } catch as e {
        return "ERROR: " . e.Message
    }
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
    ; Simple fields before url (safe to split by &)
    if RegExMatch(response, "key=([^&]+)", &m)
        result["key"] := m[1]
    if RegExMatch(response, "status=([^&]+)", &m)
        result["status"] := m[1]
    ; url field: everything after "url=" — url is always the LAST field,
    ; so it may contain literal & that must NOT be treated as delimiters.
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
    data := "key=" . StudyLink_UrlEncode(studyKey) . "&url=" . StudyLink_UrlEncode(url)
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

; ────────────────────────────────────────────────────────────────────────────
; General helpers
; ────────────────────────────────────────────────────────────────────────────

; Legacy INI path (kept for reference — not used by new API code)
StudyLink_IniPath() {
    return A_ScriptDir "\data\study_links.ini"
}

; ────────────────────────────────────────────────────────────────────────────
; Functional test: verifies both GET and SET operations against the live API.
; Returns a Map with keys: setOk (bool), getOk (bool), setMsg (str), getMsg (str).
; ────────────────────────────────────────────────────────────────────────────
StudyLink_RunFunctionalTest() {
    testKey := "subtopic_test"
    testUrl := "https://example.com/api-test-link-" . A_TickCount
    result := Map()

    ; ── Step 1: SET operation ──
    setOk := StudyLink_Set(testKey, testUrl)
    result["setOk"] := setOk
    result["setMsg"] := setOk ? "PASSED" : "FAILED"

    ; ── Step 2: GET operation (retrieve what we just set) ──
    retrieved := StudyLink_Get(testKey)
    getOk := (retrieved = testUrl)
    result["getOk"] := getOk
    result["getMsg"] := getOk ? "PASSED" : "FAILED (got: '" . retrieved . "')"

    ; ── Step 3: Cleanup — remove test key ──
    StudyLink_Set(testKey, "")

    return result
}

; ────────────────────────────────────────────────────────────────────────────
; Open the stored link for a study (delegates to Utils.ahk scoped call or Run)
; ────────────────────────────────────────────────────────────────────────────
StudyLink_Open(studyKey) {
    url := StudyLink_Get(studyKey)
    if (url != "") {
        try Run(url)
    } else {
        MsgBox "No link stored for this study."
    }
}
