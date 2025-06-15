#Requires AutoHotkey v2.0

lastHwnd := 0
SetTimer(CheckActiveWindow, 300)

CheckActiveWindow() {
    global lastHwnd
    hwnd := WinGetID("A")
    if !hwnd || hwnd = lastHwnd
        return
    lastHwnd := hwnd

    ; Get the window title to avoid processing system windows
    title := WinGetTitle("ahk_id " hwnd)
    if (title = "" || title = "Program Manager")
        return

    ; Find an available monitor
    targetMonitor := FindAvailableMonitor()
    if !targetMonitor
        return

    ; Move and maximize the window to the target monitor
    WinMove(targetMonitor.Left, targetMonitor.Top,
        targetMonitor.Right - targetMonitor.Left,
        targetMonitor.Bottom - targetMonitor.Top,
        "ahk_id " hwnd)
}

; Function to get all monitors and their work areas
GetMonitorWorkAreas() {
    monitors := Map()
    loop MonitorGetCount() {
        MonitorGetWorkArea(A_Index, &Left, &Top, &Right, &Bottom)
        monitors[A_Index] := { Left: Left, Top: Top, Right: Right, Bottom: Bottom }
    }
    return monitors
}

; Function to check if a monitor has any active windows
MonitorHasActiveWindow(monitor) {
    for hwnd in WinGetList() {
        if WinExist("ahk_id " hwnd) {
            WinGetPos(&X, &Y, &W, &H, "ahk_id " hwnd)
            if (X >= monitor.Left && X < monitor.Right &&
                Y >= monitor.Top && Y < monitor.Bottom) {
                return true
            }
        }
    }
    return false
}

; Function to find the first available monitor
FindAvailableMonitor() {
    monitors := GetMonitorWorkAreas()
    for monitorNum, monitor in monitors {
        if !MonitorHasActiveWindow(monitor) {
            return monitor
        }
    }
    return monitors[1]  ; If all monitors are busy, return the first one
}
