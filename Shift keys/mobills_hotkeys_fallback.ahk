; =============================================================================
; Shift keys module: mobills_hotkeys_fallback.ahk
; Mobills URL-scoped month nav hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

; =============================================================================
; Mobills hotkeys fallback for mobile/device mode
; - Some Chrome "mobile device" views do not keep "Mobills" in the window title,
;   which prevents the WinActive("Mobills") context from matching.
; - This fallback scopes K/L (+Shift variants) to the Mobills Transactions URL only.
; =============================================================================
global g_MobillsUrlCacheTick := 0
global g_MobillsUrlCacheUrl := ""

Mobills_GetActiveBrowserUrl(cacheMs := 250) {
    global g_MobillsUrlCacheTick, g_MobillsUrlCacheUrl
    now := A_TickCount

    if (g_MobillsUrlCacheTick && (now - g_MobillsUrlCacheTick) < cacheMs)
        return g_MobillsUrlCacheUrl

    g_MobillsUrlCacheTick := now
    g_MobillsUrlCacheUrl := ""

    try {
        if WinActive("ahk_exe chrome.exe")
            uia := UIA_Browser("ahk_exe chrome.exe")
        else if WinActive("ahk_exe msedge.exe")
            uia := UIA_Browser("ahk_exe msedge.exe")
        else
            return ""

        if uia {
            try g_MobillsUrlCacheUrl := StrLower(uia.GetCurrentURL())
        }
    } catch {
        g_MobillsUrlCacheUrl := ""
    }

    return g_MobillsUrlCacheUrl
}

Mobills_IsTransactionsUrlActive(cacheMs := 250) {
    return InStr(Mobills_GetActiveBrowserUrl(cacheMs), "/transactions")
}

; FAB / description / sidebar keys: Mobills site only, never consumer Gemini.
Mobills_ShouldHandleAppKeys() {
    if !(WinActive("ahk_exe chrome.exe") || WinActive("ahk_exe msedge.exe"))
        return false
    try
        title := WinGetTitle("A")
    catch
        title := ""
    if IsConsumerGeminiChromeTitle(title)
        return false
    url := Mobills_GetActiveBrowserUrl()
    if InStr(url, "gemini.google.com") || GeminiEnterprise_UrlMatches(url)
        return false
    if InStr(url, "web.mobills.com.br")
        return true
    return InStr(title, "Mobills")
}

; True when Chrome/Edge UIA focus is in a text-editable control (typing must not trigger month nav).
; Omit ControlType Document: the web root is often Document and would block K/L when the list has focus.
Mobills_IsWebTextInputFocused() {
    try {
        fe := UIA.GetFocusedElement()
        if !fe
            return false
        ct := fe.GetPropertyValue(UIA.Property.ControlType)
        if (ct = UIA.Type.Edit || ct = 50004)
            return true
        if (ct = UIA.Type.ComboBox || ct = 50003)
            return true
        if (ct = 50023) ; UIA Type Spinner (number inputs)
            return true
    } catch {
    }
    return false
}

; Previous/Next month: WinActive("Mobills") OR mobile fallback (transactions URL). Never while typing in a field
; (fixes bare k/l stealing keys and +k/+l from long-press on mobile keyboards).
Mobills_ShouldHandleMonthNavKeys() {
    if Mobills_IsWebTextInputFocused()
        return false
    try
        title := WinGetTitle("A")
    catch
        title := ""
    if IsConsumerGeminiChromeTitle(title)
        return false
    url := Mobills_GetActiveBrowserUrl()
    if InStr(url, "gemini.google.com") || GeminiEnterprise_UrlMatches(url)
        return false
    if WinActive("Mobills")
        return true
    if (WinActive("ahk_exe chrome.exe") || WinActive("ahk_exe msedge.exe")) && Mobills_IsTransactionsUrlActive()
        return true
    return false
}

#HotIf Mobills_ShouldHandleMonthNavKeys()

k:: Mobills_Navigate("Prev")
l:: Mobills_Navigate("Next")
+k:: Mobills_Navigate("Prev")
+l:: Mobills_Navigate("Next")

#HotIf