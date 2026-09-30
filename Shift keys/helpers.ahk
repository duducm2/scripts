; =============================================================================
; Shift keys module: helpers.ahk
; Early globals and SafeDebugLog helpers
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

; --- Global Variables ---
global DEBUG_LOG_PATH := A_ScriptDir "\.cursor\debug.log"
; Phase 5: Gate debug I/O; set to true only when diagnosing (avoids file I/O in hot paths).
global DEBUG_SHIFTKEYS := false
global g_BlackoutSuppressedUntil

; --- Blackout Banner Suppression Integration (implementation in Utils.ahk) ---
IsBlackoutSuppressed() {
    return Blackout_IsSuppressed()
}

DisableBlackout7Min(*) {
    Blackout_Disable7Min()
}

; Helper function for safe debug logging with retry on file lock
; Handles file locking gracefully by retrying with exponential backoff
; No-op when DEBUG_SHIFTKEYS is false (production).
SafeDebugLog(text) {
    if (!DEBUG_SHIFTKEYS)
        return true
    maxRetries := 3
    retryDelay := 10
    loop maxRetries {
        try {
            FileAppend text, DEBUG_LOG_PATH
            return true
        } catch Error as err {
            ; Check if error has Number property before accessing it
            ; File lock error is typically error code 32
            hasNumber := false
            errNumber := 0
            try {
                errNumber := err.Number
                hasNumber := true
            } catch {
                ; Error doesn't have Number property, treat as non-retryable
                hasNumber := false
            }

            ; If it's a file lock error (32) and we have retries left, wait and retry
            if (hasNumber && errNumber = 32 && A_Index < maxRetries) {
                Sleep retryDelay * A_Index  ; Exponential backoff
            } else {
                ; For other errors or final retry, silently fail to not interrupt script execution
                return false
            }
        }
    }
    return false
}

; Helper: find ChatGPT chrome window by case-insensitive contains match
GetChatGPTWindowHwnd() {
    for hwnd in WinGetList("ahk_exe chrome.exe") {
        if InStr(WinGetTitle("ahk_id " hwnd), "chatgpt", false)
            return hwnd
    }
    return 0
}

; Letter from $*m / $+m (single a-z), ignoring hook and modifier prefixes.
ShiftLetterHotkey_TriggerKey() {
    hk := A_ThisHotkey
    hk := RegExReplace(hk, "^[$*~]+")
    hk := RegExReplace(hk, "[#^!+<>]+", "")
    if RegExMatch(hk, "^[a-zA-Z]$")
        return StrLower(hk)
    return ""
}

; True when this keypress is Shift + the letter, with no Ctrl, Alt, or Win.
ShiftLetterHotkey_IsBareShift() {
    return GetKeyState("Shift", "P")
    && !GetKeyState("Ctrl", "P")
    && !GetKeyState("Alt", "P")
    && !GetKeyState("LWin", "P")
    && !GetKeyState("RWin", "P")
}

; AI companion tabs must not run page UI Automation from a #HotIf. That walk holds the
; keyboard hook, so Alt+P in WindowManagement never reaches Clip Angel.
Chrome_IsAiCompanionTitle(title) {
    if (!title)
        return false
    if (GeminiEnterprise_TitleMatches(title) || CopilotWeb_TitleMatchesCopilot(title))
        return true
    return IsConsumerGeminiChromeTitle(title)
}

; Address bar only. Skips the document so a large Gemini or Copilot page cannot block the hook.
; Does not turn Chrome accessibility on.
ChromeHotIf_OmniboxUrl(hwnd) {
    if (!hwnd)
        return ""
    root := 0
    try root := UIA.ElementFromHandle(hwnd, unset, 0)
    catch
        return ""
    if (!IsObject(root))
        return ""
    return ChromeHotIf_FindOmnibox(root, 0)
}

ChromeHotIf_FindOmnibox(el, depth) {
    if (!IsObject(el) || depth > 8)
        return ""
    ct := 0
    try ct := el.ControlType
    catch
        return ""
    if (ct = 50030)
        return ""
    if (ct = 50004) {
        ak := ""
        name := ""
        id := ""
        try ak := el.AcceleratorKey
        try name := el.Name
        try id := el.AutomationId
        if (ak = "Ctrl+L" || id = "view_1012" || InStr(name, "Address and search bar", false)) {
            url := ""
            try url := el.Value
            return url
        }
    }
    kids := 0
    try kids := el.FindAll(UIA.TrueCondition, 2)
    catch
        return ""
    if (!IsObject(kids))
        return ""
    for child in kids {
        found := ChromeHotIf_FindOmnibox(child, depth + 1)
        if (found != "")
            return found
    }
    return ""
}

; Pass a normal letter or a non-bare-Shift chord through to the focused field.
; The $ hotkey prefix keeps this Send from retriggering the hotkey.
ShiftLetterHotkey_Relay() {
    key := ShiftLetterHotkey_TriggerKey()
    if (key != "")
        Send "{Blind}" . key
}

; Wait out a still-held trigger key so a later Send cannot type it.
ShiftLetterHotkey_Consume() {
    key := ShiftLetterHotkey_TriggerKey()
    if (key = "" || !GetKeyState(key, "P"))
        return
    KeyWait key, "T1"
}
