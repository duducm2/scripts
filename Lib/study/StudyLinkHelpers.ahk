; StudyLink helpers for per-study subtopic links via Google Apps Script API
; API endpoint: Google Apps Script web app (GET = read, POST = write)
global STUDY_LINKS_API_URL :=
    "https://script.google.com/macros/s/AKfycbzKDLbmzGF8iduyNpaUymONEkERi089rBjW0jrYUX4a8K9ornfGwYIOsgQP1K_dfaj5/exec"

global STUDYLINK_KEY_YOUTUBE := "subtopic"
global STUDYLINK_KEY_ARTICLE := "subtopic_article"
global STUDYLINK_KEY_FAVORITE := "subtopic_favorite"
global STUDYLINK_SENTINEL_NAME := "manage, study, set, top, link"

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

; Friendly label for loading banners (video / article / favorite).
StudyLink_LoadingLabel(studyKey) {
    if (studyKey = STUDYLINK_KEY_YOUTUBE)
        return "study video"
    if (studyKey = STUDYLINK_KEY_ARTICLE)
        return "study article"
    if (studyKey = STUDYLINK_KEY_FAVORITE)
        return "favorite"
    return "link"
}

StudyLink_ShowApiLoading(actionVerb, studyKey) {
    msg := "⏳ " . actionVerb . " " . StudyLink_LoadingLabel(studyKey) . "…"
    try StandardLoadingBar_Show(msg, BANNER_ACCENT_INTERMEDIATE, { passive: false })
    catch {
    }
}

StudyLink_HideApiLoading() {
    try StandardLoadingBar_Hide(0)
    catch {
    }
}

; Returns Map: ok (true=HTTP+parse succeeded), url (decoded link or ""), err (failure message).
StudyLink_GetResult(studyKey) {
    StudyLink_ShowApiLoading("Loading", studyKey)
    try {
        response := StudyLink_HttpGet("key=" . StudyLink_UrlEncode(studyKey))
        if StudyLink_IsHttpErrorText(response) || !StudyLink_IsValidGetResponse(response) {
            return Map("ok", false, "url", "", "err", StudyLink_NormalizeApiError(response))
        }
        url := StudyLink_ExtractUrlFromResponse(response)
        return Map("ok", true, "url", url, "err", "")
    } finally {
        StudyLink_HideApiLoading()
    }
}

; Backward-compatible: returns URL string, or "" if not set or on error.
StudyLink_Get(studyKey) {
    r := StudyLink_GetResult(studyKey)
    return (r["ok"] && r["url"] != "") ? r["url"] : ""
}

StudyLink_FormatLinkLabel(linkResult) {
    if !linkResult["ok"]
        return "(API error — " . linkResult["err"] . ")"
    return linkResult["url"] != "" ? linkResult["url"] : "(none)"
}

StudyLink_OpenUrlInChrome(url, newWindow := false) {
    if (Trim(url) = "")
        return false
    chromeCmd := newWindow ? 'chrome.exe --new-window "' url '"' : 'chrome.exe "' url '"'
    try Run(chromeCmd)
    catch
        try Run(url)
    return true
}

StudyLink_IsValidHttpUrl(url) {
    u := Trim(url)
    if (u = "")
        return false
    if (SubStr(u, 1, 11) = "javascript:")
        return false
    return (SubStr(u, 1, 7) = "http://" || SubStr(u, 1, 8) = "https://")
}

; Save the http(s) URL currently on the clipboard (no Chrome / address-bar capture).
StudyLink_SetFromClipboard(studyKey, successLabel := "link") {
    clip := ""
    try clip := Trim(A_Clipboard)
    catch {
    }
    if !StudyLink_IsValidHttpUrl(clip) {
        try ShowCenteredOverlay_Utils("❌ Clipboard has no valid http(s) URL.", 3000, BANNER_ACCENT_ERROR)
        return false
    }
    setOk := StudyLink_Set(studyKey, clip)
    if setOk
        ShowCenteredOverlay_Utils("✅ " . successLabel . " saved from clipboard.", 3000, BANNER_ACCENT_SUCCESS)
    else
        ShowCenteredOverlay_Utils("❌ Could not save the " . successLabel . " (API failed).", 3500, BANNER_ACCENT_ERROR)
    return setOk
}

; Prompt for a URL (clipboard prefill when valid) and POST via StudyLink_Set.
StudyLink_SetFromManualInput(studyKey, successLabel := "link") {
    defaultUrl := ""
    try {
        clip := Trim(A_Clipboard)
        if StudyLink_IsValidHttpUrl(clip)
            defaultUrl := clip
    } catch {
    }

    ib := InputBox("Paste or type the URL:", "Set " . successLabel . " manually", "w500 h120", defaultUrl)
    if (ib.Result != "OK") {
        try ShowCenteredOverlay_Utils("⚠ Manual set cancelled.", 2000, BANNER_ACCENT_INTERMEDIATE)
        return false
    }
    url := Trim(ib.Value)
    if (url = "") {
        try ShowCenteredOverlay_Utils("⚠ URL cannot be empty.", 2500, BANNER_ACCENT_INTERMEDIATE)
        return false
    }
    if !StudyLink_IsValidHttpUrl(url) {
        try ShowCenteredOverlay_Utils("❌ URL must start with http:// or https://", 3000, BANNER_ACCENT_ERROR)
        return false
    }

    setOk := StudyLink_Set(studyKey, url)
    if setOk
        ShowCenteredOverlay_Utils("✅ " . successLabel . " saved to study notes.", 3000, BANNER_ACCENT_SUCCESS)
    else
        ShowCenteredOverlay_Utils("❌ Could not save the " . successLabel . " (API failed).", 3500, BANNER_ACCENT_ERROR)
    return setOk
}
; True when POST response indicates the link was saved (see StudyLink_Set).
StudyLink_ResponseIndicatesSetSuccess(response) {
    if (SubStr(response, 1, 6) = "ERROR:")
        return false
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
; Public API: set a link for a study key
; Sends POST with key and url; accepts status=ok form data or plain-text "Saved"
; Returns true on success, false otherwise.
; ────────────────────────────────────────────────────────────────────────────
StudyLink_Set(studyKey, url) {
    ; Match MacroDroid Set_Video.macro: literal url in body (Apps Script reads everything after url=).
    StudyLink_ShowApiLoading("Saving", studyKey)
    try {
        data := "key=" . StudyLink_UrlEncode(studyKey) . "&url=" . url
        response := StudyLink_HttpPost(data)
        if (!StudyLink_ResponseIndicatesSetSuccess(response))
            return false
        StudyLink_PlayApiSuccessSound()
        return true
    } finally {
        StudyLink_HideApiLoading()
    }
}

; Legacy INI path (kept for reference — not used by new API code)
StudyLink_IniPath() {
    return A_ScriptDir "\assets\data\study_links.ini"
}

StudyLink_ManageSubtopicSentinelPath() {
    return A_MyDocuments "\" STUDYLINK_SENTINEL_NAME
}

StudyLink_EnsureManageSubtopicSentinel() {
    path := StudyLink_ManageSubtopicSentinelPath()
    if FileExist(path)
        return path
    try {
        if !DirExist(A_MyDocuments)
            return ""
        f := FileOpen(path, "w")
        if IsObject(f)
            f.Close()
    } catch {
        return ""
    }
    return FileExist(path) ? path : ""
}

StudyLink_RunKeyRoundTrip(testKey, testUrl, &setMsg := "", &getMsg := "") {
    setOk := StudyLink_Set(testKey, testUrl)
    setMsg := setOk ? "PASSED" : "FAILED"
    getResult := StudyLink_GetResult(testKey)
    getOk := false
    if (getResult["ok"] && getResult["url"] = testUrl) {
        getOk := true
        getMsg := "PASSED"
    } else if !getResult["ok"] {
        getMsg := "FAILED (" . getResult["err"] . ")"
    } else {
        getMsg := "FAILED (got: '" . getResult["url"] . "')"
    }
    StudyLink_Set(testKey, "")
    return { setOk: setOk, getOk: getOk }
}

; Functional test: verifies GET and SET for YouTube, article, and favorite keys against the live API.
StudyLink_RunFunctionalTest() {
    tick := A_TickCount
    result := Map()
    yt := StudyLink_RunKeyRoundTrip("subtopic_test", "https://example.com/yt-test-" . tick, &ytSetMsg, &ytGetMsg)
    art := StudyLink_RunKeyRoundTrip("subtopic_article_test", "https://example.com/article-test-" . tick, &artSetMsg,
        &artGetMsg)
    fav := StudyLink_RunKeyRoundTrip(STUDYLINK_KEY_FAVORITE, "https://example.com/favorite-test-" . tick, &favSetMsg,
        &favGetMsg)
    result["youtubeSetOk"] := yt.setOk
    result["youtubeGetOk"] := yt.getOk
    result["youtubeSetMsg"] := ytSetMsg
    result["youtubeGetMsg"] := ytGetMsg
    result["articleSetOk"] := art.setOk
    result["articleGetOk"] := art.getOk
    result["articleSetMsg"] := artSetMsg
    result["articleGetMsg"] := artGetMsg
    result["favoriteSetOk"] := fav.setOk
    result["favoriteGetOk"] := fav.getOk
    result["favoriteSetMsg"] := favSetMsg
    result["favoriteGetMsg"] := favGetMsg
    result["setOk"] := yt.setOk && art.setOk && fav.setOk
    result["getOk"] := yt.getOk && art.getOk && fav.getOk
    result["setMsg"] := result["setOk"] ? "PASSED (YouTube + article + favorite)" : "FAILED"
        result["getMsg"] := result["getOk"] ? "PASSED (YouTube + article + favorite)" : "FAILED"
            return result
}

StudyLink_Open(studyKey) {
    r := StudyLink_GetResult(studyKey)
    if (r["ok"] && r["url"] != "") {
        StudyLink_OpenUrlInChrome(r["url"], true)
    } else if !r["ok"] {
        MsgBox "Could not load link: " . r["err"]
    } else {
        MsgBox "No link stored for this study."
    }
}

; Memory Palace–matched manage menu (dark #1E1E1E, gold keys, card rows).
; items: [["1","Open","…"], …]   pairs: [["1", callback], …]
StudyLink_CenterGui(guiObj, w := 560, h := 420) {
    MonitorGetWorkArea(MonitorGetPrimary(), &L, &T, &R, &B)
    x := L + ((R - L) - w) // 2
    y := T + ((B - T) - h) // 2
    guiObj.Show("x" . x . " y" . y . " w" . w . " h" . h)
}

StudyLink_TruncateLabel(s, maxLen := 72) {
    t := Trim(s)
    if (StrLen(t) <= maxLen)
        return t
    return SubStr(t, 1, maxLen - 1) . "…"
}

StudyLink_UnbindManageMenuHotkeys() {
    loop 9 {
        try Hotkey(String(A_Index), "Off")
    }
    try Hotkey("Escape", "Off")
    try Hotkey("Backspace", "Off")
}

StudyLink_ClosePalaceManageGui(&guiRef) {
    StudyLink_UnbindManageMenuHotkeys()
    StudyTopicSelector_SafeDestroyGui(guiRef)
    guiRef := false
    StudyTopicSelector_ResumeSelectorEscapeAfterLinks()
    global g_StudyTopicSelectorActive
    if (!g_StudyTopicSelectorActive) {
        try Palace_ShowMainMenu()
        catch {
        }
    }
}

StudyLink_ShowPalaceManageGui(&guiRef, windowTitle, heading, statusLine, items, pairs, onEscape) {
    StudyLink_UnbindManageMenuHotkeys()

    guiRef := Gui("+AlwaysOnTop +ToolWindow", windowTitle)
    guiRef.SetFont("s10", "Segoe UI")
    guiRef.BackColor := "1E1E1E"
    guiRef.OnEvent("Close", onEscape)
    guiRef.OnEvent("Escape", onEscape)

    guiRef.SetFont("s16 cWhite Bold", "Segoe UI")
    guiRef.Add("Text", "x20 y16 w520", heading)
    guiRef.SetFont("s10 cC0C0C0 Norm", "Segoe UI")
    guiRef.Add("Text", "x20 y48 w520", StudyLink_TruncateLabel(statusLine, 90))
    guiRef.SetFont("s9 cF1C40F", "Segoe UI")
    guiRef.Add("Text", "x20 y72 w520", "Press 1–3 · Backspace / Esc — back")

    y0 := 108
    rowH := 72
    for it in items {
        y := y0 + (A_Index - 1) * rowH
        guiRef.SetFont("s14 cF1C40F Bold", "Segoe UI")
        guiRef.Add("Text", "x32 y" . (y + 10) . " w40 BackgroundTrans", "[" . it[1] . "]")
        guiRef.SetFont("s12 cWhite Bold", "Segoe UI")
        guiRef.Add("Text", "x78 y" . (y + 10) . " w440 BackgroundTrans", it[2])
        guiRef.SetFont("s9 cA0A0A0 Norm", "Segoe UI")
        guiRef.Add("Text", "x78 y" . (y + 36) . " w440 BackgroundTrans", it[3])
    }

    footerY := y0 + items.Length * rowH + 12
    guiRef.SetFont("s9 c808080", "Segoe UI")
    guiRef.Add("Text", "x20 y" . footerY . " w520", "Backspace / Esc — Memory Palace menu")

    for p in pairs {
        try Hotkey(p[1], p[2], "On")
    }
    try Hotkey("Escape", onEscape, "On")
    try Hotkey("Backspace", onEscape, "On")

    try DllCall("dwmapi\DwmSetWindowAttribute", "ptr", guiRef.Hwnd, "uint", 20, "int*", 1, "int", 4)
    catch {
    }
    winH := footerY + 40
    StudyLink_CenterGui(guiRef, 560, winH)
}
