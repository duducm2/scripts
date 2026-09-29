#Requires AutoHotkey v2.0
#SingleInstance Force
Persistent

; Run this after the other AutoHotkey scripts so this hook is called first.
; Time spent inside CallNextHookEx is the delay those earlier hooks added
; before the key reached the focused app. Lines go to %TEMP%\ahk-hook-lag.log
; (local disk). A flag file in %TEMP% turns on Shift keys predicate timing.
; Exit from the tray icon: the hook is removed and the flag file is deleted.

global HOOK_LAG_LOG := A_Temp "\ahk-hook-lag.log"
global HOOK_LAG_FLAG := A_Temp "\ahk-hook-lag.on"
global HOOK_LAG_MIN_MS := 15
global HOOK_LAG_TRAY_MS := 200

global g_ProbeHook := 0
global g_ProbeCb := 0
global g_ProbeBuf := []
global g_ProbeWorstMs := 0
global g_ProbeWorstText := ""
global g_ProbeOwnerPid := DllCall("GetCurrentProcessId", "UInt")

Probe_Install() {
    g_ProbeCb := CallbackCreate(Probe_LowLevelKeyboard, , 3)
    g_ProbeHook := DllCall("SetWindowsHookExW", "Int", 13, "Ptr", g_ProbeCb, "Ptr", 0, "UInt", 0, "Ptr")
    if !g_ProbeHook {
        TrayTip("Could not install the keyboard hook.", "Hook lag probe", "Iconx")
        ExitApp
    }
}

Probe_LowLevelKeyboard(nCode, wParam, lParam) {
    ; Time key-down only. CallNextHookEx(NULL) is valid if the handle is not stored yet.
    if (nCode != 0 || (wParam != 0x100 && wParam != 0x104))
        return DllCall("CallNextHookEx", "Ptr", g_ProbeHook, "Int", nCode, "Ptr", wParam, "Ptr", lParam, "Ptr")
    t0 := A_TickCount
    ret := DllCall("CallNextHookEx", "Ptr", g_ProbeHook, "Int", nCode, "Ptr", wParam, "Ptr", lParam, "Ptr")
    elapsed := A_TickCount - t0
    if (elapsed >= HOOK_LAG_MIN_MS) {
        vk := 0
        try vk := NumGet(lParam, 0, "UInt")
        fg := Probe_ForegroundExe()
        g_ProbeBuf.Push(Format("tick={1}`tprobe`t{2}ms`tvk={3}`tfg={4}", t0, elapsed, vk, fg))
        if (g_ProbeBuf.Length > 200)
            g_ProbeBuf.RemoveAt(1)
        if (elapsed >= HOOK_LAG_TRAY_MS && elapsed >= g_ProbeWorstMs) {
            g_ProbeWorstMs := elapsed
            g_ProbeWorstText := elapsed "ms  vk=" vk "  " fg
        }
    }
    return ret
}

Probe_ForegroundExe() {
    hwnd := DllCall("GetForegroundWindow", "Ptr")
    if !hwnd
        return ""
    pid := 0
    DllCall("GetWindowThreadProcessId", "Ptr", hwnd, "UInt*", &pid)
    if !pid
        return ""
    ; PROCESS_QUERY_LIMITED_INFORMATION
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

Probe_Flush(*) {
    if (g_ProbeWorstMs) {
        TrayTip(g_ProbeWorstText, "Hook lag", "Icon!")
        g_ProbeWorstMs := 0
        g_ProbeWorstText := ""
    }
    if !g_ProbeBuf.Length
        return
    text := ""
    for line in g_ProbeBuf
        text .= line "`n"
    g_ProbeBuf := []
    try FileAppend(text, HOOK_LAG_LOG, "UTF-8-RAW")
}

Probe_Exit(*) {
    ; A restarted probe may already have replaced the flag with its own pid.
    try {
        if (Trim(FileRead(HOOK_LAG_FLAG, "UTF-8")) = String(g_ProbeOwnerPid))
            FileDelete(HOOK_LAG_FLAG)
    }
    Probe_Flush()
    if (g_ProbeHook) {
        DllCall("UnhookWindowsHookEx", "Ptr", g_ProbeHook)
        g_ProbeHook := 0
    }
    if (g_ProbeCb) {
        CallbackFree(g_ProbeCb)
        g_ProbeCb := 0
    }
}

try FileDelete(HOOK_LAG_LOG)
try FileDelete(HOOK_LAG_FLAG)
FileAppend("hook lag probe started " FormatTime(, "yyyy-MM-dd HH:mm:ss") "`n", HOOK_LAG_LOG, "UTF-8")
FileAppend(g_ProbeOwnerPid, HOOK_LAG_FLAG)
OnExit(Probe_Exit)
Probe_Install()
SetTimer(Probe_Flush, 1000)
