; Predicate timing for the hook-lag probe (infra\tools\HookLagProbe.ahk).
; Included from Utils.ahk so every script that loads Copilot/Gemini already
; has these functions. Off unless the probe created the flag file.
; Nothing is written on the key; a timer flushes slow calls (>= 15ms).

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
    g_HookTimingBuf.Push(Format("tick={1}`tpred`t{2}ms`t{3}`tfg={4}", item["tick"], elapsed, item["name"],
        HookTiming_ForegroundExe()))
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
