#Requires AutoHotkey v2.0+
; Standalone AHK latency harness for WindowManagement daemon IPC.
; Connects to \\.\pipe\wm_automation, sends Ping/echo requests, records round-trip times,
; reports p50/p95 and jitter. Run with daemon started: python python\wm_daemon.py

#Include "%A_ScriptDir%\WMIPC.ahk"

global WM_PIPE_NAME := "\\.\pipe\wm_automation"
global WM_IPC_TIMEOUT_MS := 5000

; Ensure harness can connect (ignore feature flags)
if (!WMIPC_ConnectForHarness()) {
    MsgBox("Harness: Could not connect to pipe " WM_PIPE_NAME "`. Start the daemon: python python\wm_daemon.py",
        "WM IPC Harness", "Icon!")
    ExitApp(1)
}

; Number of round-trips for latency stats
iterations := 100
latencies := []
payload := Map("clientTs", A_TickCount, "n", 0)

loop iterations {
    payload["n"] := A_Index
    payload["clientTs"] := A_TickCount
    t0 := A_TickCount
    resp := WMIPC_SendRequest("Ping-" A_Index, "Ping", "", payload, WM_IPC_TIMEOUT_MS)
    t1 := A_TickCount
    rtt := t1 - t0
    latencies.Push(rtt)
    if (!resp.Has("ok") || resp["ok"] != true) {
        MsgBox("Harness: Ping failed at iteration " A_Index ". RTT=" rtt " ms", "WM IPC Harness", "Icon!")
        break
    }
}

WMIPC_Disconnect()

; Sort for percentiles (copy and sort)
sorted := []
for rtt in latencies
    sorted.Push(rtt)
; Bubble sort (simple for small n)
loop sorted.Length - 1
    loop sorted.Length - A_Index
        if (sorted[A_Index] > sorted[A_Index + 1]) {
            t := sorted[A_Index]
            sorted[A_Index] := sorted[A_Index + 1]
            sorted[A_Index + 1] := t
        }

n := sorted.Length
p50 := sorted[Max(1, Round(n * 0.5))]
p95 := sorted[Max(1, Round(n * 0.95))]
sum := 0
for r in sorted
    sum += r
mean := sum / n
; Jitter: mean absolute deviation of consecutive differences (or simple std dev approximation)
jitterSum := 0
loop n - 1
    jitterSum += Abs(sorted[A_Index + 1] - sorted[A_Index])
jitter := (n > 1) ? (jitterSum / (n - 1)) : 0

msg := "WM IPC latency (Ping echo, " n " rounds)`n`n"
msg .= "p50: " p50 " ms`np95: " p95 " ms`nmean: " Round(mean, 2) " ms`njitter (avg diff): " Round(jitter, 2) " ms"
MsgBox(msg, "WM IPC Harness", "Iconi")
ExitApp(0)