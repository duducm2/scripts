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

; Predicate timing for the hook-lag probe (infra\tools\HookLagProbe.ahk).
; Off unless that probe created the flag file. Nothing is written on the key;
; a timer flushes slow calls (>= 15ms) to the same local log as the probe.
global HOOK_TIMING_FLAG := A_Temp "\ahk-hook-lag.on"
global HOOK_TIMING_LOG := A_Temp "\ahk-hook-lag.log"
global HOOK_TIMING_MIN_MS := 15
global g_HookTimingOn := false
global g_HookTimingCheckedAt := 0
global g_HookTimingTimerOn := false
global g_HookTimingStack := []
global g_HookTimingBuf := []

HookTiming_IsOn() {
    now := A_TickCount
    if (g_HookTimingCheckedAt && (now - g_HookTimingCheckedAt) < 1000)
        return g_HookTimingOn
    g_HookTimingCheckedAt := now
    g_HookTimingOn := !!FileExist(HOOK_TIMING_FLAG)
    if (g_HookTimingOn && !g_HookTimingTimerOn) {
        g_HookTimingTimerOn := true
        SetTimer(HookTiming_Flush, 1000)
    } else if (!g_HookTimingOn && g_HookTimingTimerOn) {
        g_HookTimingTimerOn := false
        SetTimer(HookTiming_Flush, 0)
        g_HookTimingBuf := []
        g_HookTimingStack := []
    }
    return g_HookTimingOn
}

HookTiming_Begin(name) {
    if !HookTiming_IsOn()
        return
    g_HookTimingStack.Push(Map("name", name, "tick", A_TickCount))
}

HookTiming_End(name) {
    if !g_HookTimingStack.Length
        return
    item := g_HookTimingStack.Pop()
    elapsed := A_TickCount - item["tick"]
    if (elapsed < HOOK_TIMING_MIN_MS)
        return
    g_HookTimingBuf.Push(Format("tick={1}`tpred`t{2}ms`t{3}`tfg={4}", item["tick"], elapsed, item["name"], HookTiming_ForegroundExe()))
    if (g_HookTimingBuf.Length > 200)
        g_HookTimingBuf.RemoveAt(1)
}

HookTiming_Flush(*) {
    if !g_HookTimingBuf.Length
        return
    text := ""
    for line in g_HookTimingBuf
        text .= line "`n"
    g_HookTimingBuf := []
    try FileAppend(text, HOOK_TIMING_LOG, "UTF-8-RAW")
}

HookTiming_ForegroundExe() {
    hwnd := DllCall("GetForegroundWindow", "Ptr")
    if !hwnd
        return ""
    pid := 0
    DllCall("GetWindowThreadProcessId", "Ptr", hwnd, "UInt*", &pid)
    if !pid
        return ""
    h := DllCall("OpenProcess", "UInt", 0x1000, "Int", 0, "UInt", pid, "Ptr")
    if !h
        return "pid=" pid
    chars := 260
    buf := Buffer(chars * 2, 0)
    ok := DllCall("QueryFullProcessImageNameW", "Ptr", h, "UInt", 0, "Ptr", buf, "UInt*", &chars)
    DllCall("CloseHandle", "Ptr", h)
    if !ok
        return "pid=" pid
    full := StrGet(buf, "UTF-16")
    SplitPath(full, &name)
    return name
}
