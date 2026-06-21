#Requires AutoHotkey v2.0
#SingleInstance Force
; One-shot snapshot: AHK MonitorGet numbering + same left-to-right sort as WindowManagement GetMonitorIndexByOrder.
; Output: tools\MonitorEnumerationSnapshot-out.txt
; Pair with: infra\python\compare_monitor_enumeration.py

outPath := A_ScriptDir "\MonitorEnumerationSnapshot-out.txt"
count := MonitorGetCount()
lines := []
lines.Push("AHK MonitorGetCount: " count)
lines.Push("")
lines.Push("Raw MonitorGet i (AHK index) -> work rect l,t,r,b -> center cx,cy:")
loop count {
    i := A_Index
    MonitorGet i, &l, &t, &r, &b
    cx := (l + r) // 2
    cy := (t + b) // 2
    lines.Push(Format("  i={}  l={} t={} r={} b={}  cx={} cy={}", i, l, t, r, b, cx, cy))
}

monitors := []
loop count {
    i := A_Index
    MonitorGet i, &l, &t, &r, &b
    cx := (l + r) // 2
    cy := (t + b) // 2
    monitors.Push({ idx: i, cx: cx, cy: cy })
}
n := monitors.Length
loop n - 1 {
    i := A_Index
    loop n - i {
        j := A_Index
        a := monitors[j]
        b := monitors[j + 1]
        if (a.cx > b.cx || (a.cx == b.cx && a.cy > b.cy)) {
            monitors[j] := b
            monitors[j + 1] := a
        }
    }
}

lines.Push("")
lines.Push("Sorted left-to-right (ordinal -> AHK idx, same bubble sort as WindowManagement.ahk):")
loop monitors.Length {
    ord := A_Index
    m := monitors[A_Index]
    lines.Push(Format("  ordinal {} -> AHK idx {}  cx={} cy={}", ord, m.idx, m.cx, m.cy))
}

lines.Push("")
if (monitors.Length >= 1)
    lines.Push("Ordinal 1 resolves to AHK monitor index: " monitors[1].idx)
else
    lines.Push("No monitors sorted (MonitorGetCount was 0).")
try {
    if FileExist(outPath)
        FileDelete outPath
} catch {
}
FileAppend StrJoin(lines, "`n") "`n", outPath

StrJoin(arr, sep) {
    s := ""
    for i, x in arr {
        if (i > 1)
            s .= sep
        s .= x
    }
    return s
}
