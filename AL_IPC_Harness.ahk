; AppLauncher IPC sandbox harness: verifies MMF roundtrip with daemon.
; Start daemon first: python python\applauncher_daemon.py
; Run this script to measure roundtrip latency and sequence integrity.

#Requires AutoHotkey v2.0+
#include %A_ScriptDir%\AppLauncherIPC.ahk

global AL_USE_DAEMON := true
global AL_USE_MMF_IPC := true

results := []
loop 20 {
    start := A_TickCount
    resp := AL_IPC_Call("Ping", Map("clientTs", A_TickCount), 3000)
    elapsed := A_TickCount - start
    results.Push(elapsed)
    if (resp.Has("ok") && resp["ok"])
        continue
    MsgBox("Harness: Ping failed or daemon not running. Start: python python\applauncher_daemon.py")
    ExitApp 1
}

; Sort for p50/p95
results.Sort()
n := results.Length
p50 := n >= 1 ? results[Max(1, Integer(n * 0.5))] : 0
p95 := n >= 1 ? results[Max(1, Integer(n * 0.95))] : 0
avg := 0
for t in results
    avg += t
avg := avg / n
MsgBox("Roundtrip (ms): avg=" . Round(avg, 2) . " p50=" . p50 . " p95=" . p95 . "`nDaemon OK.")
AL_IPC_Teardown()