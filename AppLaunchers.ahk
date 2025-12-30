#Requires AutoHotkey v2.0+
#SingleInstance Force

; -----------------------------------------------------------------------------
; This script consolidates all Application/Website launcher hotkeys.
; -----------------------------------------------------------------------------

; --- Includes ----------------------------------------------------------------
#include %A_ScriptDir%\env.ahk
#include UIA-v2\Lib\UIA.ahk
#include UIA-v2\Lib\UIA_Browser.ahk
#include %A_ScriptDir%\Utils.ahk

; --- Global Variables ---
global DEBUG_LOG_PATH := A_ScriptDir "\.cursor\debug.log"

; Helper function for safe debug logging with retry on file lock
; Handles file locking gracefully by retrying with exponential backoff
SafeDebugLog(text) {
    maxRetries := 3
    retryDelay := 10
    loop maxRetries {
        try {
            FileAppend text, DEBUG_LOG_PATH
            return true
        } catch Error as err {
            ; If it's a file lock error (32) and we have retries left, wait and retry
            if (err.Number = 32 && A_Index < maxRetries) {
                Sleep retryDelay * A_Index  ; Exponential backoff
            } else {
                ; For other errors or final retry, silently fail to not interrupt script execution
                return false
            }
        }
    }
    return false
}

; --- Hotkeys & Functions -----------------------------------------------------

; =============================================================================
; Open/Activate Cursor with specific window requirements
; Hotkey: Win+Alt+Shift+N
; Original File: OneNote - Open.ahk (Updated for Cursor)
; =============================================================================
#!+n::
{
    ; Look for Cursor windows with specific names: habits, home, punctual, or work
    ; Exclude windows containing "preview"
    targetWindow := ""
    fallbackWindow := ""

    ; Get all Cursor windows (support both Cursor.exe and Code.exe just in case)
    for proc in ["ahk_exe Cursor.exe", "ahk_exe Code.exe"] {
        for hwnd in WinGetList(proc) {
            try {
                winTitle := WinGetTitle("ahk_id " hwnd)
                winTitleLower := StrLower(winTitle)

                ; Skip any preview windows
                if InStr(winTitleLower, "preview")
                    continue

                ; Store the first valid Cursor window as fallback
                if (!fallbackWindow)
                    fallbackWindow := "ahk_id " hwnd

                ; Check if window contains any of the target names
                if (InStr(winTitleLower, "habits")
                || InStr(winTitleLower, "home")
                || InStr(winTitleLower, "punctual")
                || InStr(winTitleLower, "work")) {
                    targetWindow := "ahk_id " hwnd
                    break
                }
            } catch {
                ; Silently skip invalid windows
            }
        }
        if (targetWindow)
            break
    }

    if (targetWindow) {
        ; Found a matching Cursor window
        WinActivate(targetWindow)
        if WinWaitActive(targetWindow, , 2) {
            CenterMouse()
            Sleep(100)  ; Small delay to ensure window is fully active
            Send("^t")  ; Send Ctrl+Shift+O
        }
    } else if (fallbackWindow) {
        ; No specific window found, but found a general Cursor window - use fallback
        WinActivate(fallbackWindow)
        if WinWaitActive(fallbackWindow, , 2) {
            CenterMouse()
            Sleep(100)  ; Small delay to ensure window is fully active
            Send("^t")  ; Send Ctrl+Shift+O
        }
    } else {
        ; No Cursor window found at all - show fallback panel
        ShowCursorFallbackPanel()
    }
}

; =============================================================================
; Show fallback panel when no matching Cursor window is found
; =============================================================================
ShowCursorFallbackPanel() {
    ; Create a notification panel similar to ChatGPT.ahk style
    fallbackGui := Gui()
    fallbackGui.Opt("+AlwaysOnTop -Caption +ToolWindow")
    fallbackGui.BackColor := "3772FF"  ; Blue background like ChatGPT notifications
    fallbackGui.SetFont("s20 cFFFFFF Bold", "Segoe UI")

    ; Create the message
    message := "No Cursor window found with names: habits, home, punctual, or work"
    fallbackGui.Add("Text", "w600 Center", message)

    ; Center on active monitor
    activeWin := WinGetID("A")
    if (activeWin) {
        WinGetPos(&winX, &winY, &winW, &winH, activeWin)
    } else {
        MonitorGetWorkArea(1, &l, &t, &r, &b)
        winX := l
        winY := t
        winW := r - l
        winH := b - t
    }

    fallbackGui.Show("AutoSize Hide")
    guiW := 0, guiH := 0
    fallbackGui.GetPos(, , &guiW, &guiH)

    guiX := winX + (winW - guiW) / 2
    guiY := winY + (winH - guiH) / 2
    fallbackGui.Show("x" . Round(guiX) . " y" . Round(guiY) . " NA")
    WinSetTransparent(178, fallbackGui)

    ; Auto-hide after 3 seconds
    SetTimer(() => fallbackGui.Destroy(), -3000)
}

; =============================================================================
; Open/Activate Desktop in Explorer
; Hotkey: Shift+Win+E
; Original File: Open Desktop.ahk
; =============================================================================
+#e::
{
    SetTitleMatchMode 2
    if WinExist("Área de Trabalho ahk_class CabinetWClass") || WinExist("Desktop ahk_class CabinetWClass") {
        WinActivate
        WinWaitActive("ahk_class CabinetWClass", , 2)  ; Wait up to 2 seconds for activation
        Sleep(100)  ; Small additional delay to ensure window is ready
        CenterMouse()
    } else {
        target := IS_WORK_ENVIRONMENT ? "C:\Users\fie7ca\Desktop" : "C:\Users\eduev\OneDrive\Desktop"
        Run 'explorer.exe "' target '"'
        ; Wait until Explorer window appears AND becomes active
        WinWait("ahk_class CabinetWClass")
        WinWaitActive("ahk_class CabinetWClass", , 2)  ; Wait up to 2 seconds for activation
        Sleep(100)  ; Small additional delay to ensure window is ready
        CenterMouse()
    }
}

; =============================================================================
; Open/Activate Google Chrome
; Hotkey: Win+Alt+Shift+F
; Original File: Open Google.ahk
; =============================================================================
#!+f::
{
    Run "chrome.exe"
    WinWaitActive("ahk_exe chrome.exe")
    CenterMouse()
}

; =============================================================================
; Open/Activate WhatsApp
; Hotkey: Win+Alt+Shift+Z
; Original File: Open WhatsApp.ahk
; =============================================================================
#!+z::
{
    SetTitleMatchMode(2)
    if WinExist("WhatsApp") {
        WinActivate("WhatsApp")
        CenterMouse()
    } else {
        if (IS_WORK_ENVIRONMENT) {
            Run "C:\Users\fie7ca\Documents\Shortcuts\WhatsApp.lnk"
        } else {
            Run "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\apps do Chrome\WhatsApp Web.lnk"
        }
        WinWaitActive("WhatsApp")
        CenterMouse()
    }
}

; =============================================================================
; Open/Activate YouTube
; Hotkey: Win+Alt+Shift+H
; Original File: Youtube - Activate.ahk
; =============================================================================
#!+h::
{
    SetTitleMatchMode 2
    if WinExist("YouTube") {
        WinActivate
        CenterMouse()
    } else {
        Run "chrome.exe --new-window https://www.youtube.com/feed/playlists"
        WinWaitActive("YouTube")
        CenterMouse()
    }
}

; =============================================================================
; Open/Activate Gmail
; Hotkey: Win+Alt+Shift+W
; =============================================================================
#!+w::
{
    SetTitleMatchMode 2
    if WinExist("Gmail ahk_exe chrome.exe") {
        WinActivate
        CenterMouse()
    } else {
        target := IS_WORK_ENVIRONMENT ?
            "C:\Users\fie7ca\Documents\Shortcuts\Gmail.lnk" :
                "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Chrome Apps\Gmail.lnk"
        Run target
        WinWaitActive("Gmail ahk_exe chrome.exe")
        CenterMouse()
    }
}

; =============================================================================
; Open/Activate Cursor
; Hotkey: Win+Alt+Shift+,
; =============================================================================
#!+,::
{
    SetTitleMatchMode 2
    if WinExist("ahk_exe Cursor.exe") {
        WinActivate
        CenterMouse()
    } else {
        target := IS_WORK_ENVIRONMENT ? "C:\\Users\\fie7ca\\AppData\\Local\\Programs\\cursor\\Cursor.exe" :
            "C:\\Users\\eduev\\AppData\\Local\\Programs\\cursor\\Cursor.exe"
        Run target
        WinWaitActive("ahk_exe Cursor.exe")
        CenterMouse()
    }
}

; =============================================================================
; Wikipedia Selector with Character Shortcuts
; Hotkey: Win+Alt+Shift+K
; Shows a GUI with Wikipedia article options. Pressing a character (1-5)
; immediately opens the corresponding article or performs the action.
; =============================================================================

; Global variables for Wikipedia selector
global g_WikipediaSelectorGui := false
global g_WikipediaSelectorActive := false
global g_WikipediaSelectorHandlers := []  ; Store hotkey handlers for cleanup

; Global variables for Wikipedia scroll position save/restore
global g_WikipediaScrollPositionsFile := A_ScriptDir "\data\wikipedia_scroll_positions.ini"

; Wikipedia article items configuration
; Item 1: Taoism
; Item 2: Claude Debussy
; Items 3-5: Placeholders (no action)
global g_WikipediaItems := [{ char: "1", title: "Taoism", url: "https://en.wikipedia.org/wiki/Taoism" }, { char: "2",
    title: "Claude Debussy", url: "https://en.wikipedia.org/wiki/Claude_Debussy" }, { char: "3", title: "Placeholder",
        url: "" }, { char: "4", title: "Placeholder",
            url: "" }, { char: "5", title: "Placeholder", url: "" }
]

; =============================================================================
; Wikipedia Scroll Position Storage Functions
; =============================================================================

; Check if the active window is on Monitor 3
IsWindowOnMonitor3() {
    hwnd := WinExist("A")

    if (!hwnd) {
        return false
    }

    rect := Buffer(16, 0)
    if (!DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect)) {
        return false
    }

    left := NumGet(rect, 0, "int")
    top := NumGet(rect, 4, "int")
    right := NumGet(rect, 8, "int")
    bottom := NumGet(rect, 12, "int")

    centerX := left + (right - left) // 2
    centerY := top + (bottom - top) // 2

    monitorCount := MonitorGetCount()
    loop monitorCount {
        MonitorGet(A_Index, &ml, &mt, &mr, &mb)
        if (centerX >= ml && centerX <= mr && centerY >= mt && centerY <= mb) {
            isMonitor3 := (A_Index = 3)
            return isMonitor3
        }
    }

    return false
}

; Get current Wikipedia article URL from the active Chrome window
GetWikipediaURL() {
    try {
        if (!WinActive("ahk_exe chrome.exe") || !InStr(WinGetTitle("A"), "Wikipedia")) {
            return ""
        }
        uia := UIA_Browser("ahk_exe chrome.exe")
        url := uia.GetCurrentURL()
        if (url = "" || !InStr(url, "wikipedia.org")) {
            return ""
        }
        ; Normalize URL - remove trailing slashes and fragments
        url := RegExReplace(url, "/#.*$", "")
        url := RegExReplace(url, "/+$", "")
        return url
    } catch Error as err {
        return ""
    }
}

; Save scroll position for a Wikipedia article URL
; Now saves as percentage (0.0 to 1.0) instead of absolute pixels
SaveWikipediaScrollPosition(url, scrollPercentage) {
    global g_WikipediaScrollPositionsFile
    try {
        if (url = "" || scrollPercentage = "" || scrollPercentage < 0 || scrollPercentage > 1) {
            return false
        }
        ; Ensure directory exists
        SplitPath(g_WikipediaScrollPositionsFile, , &dir)
        if (dir != "" && !DirExist(dir)) {
            DirCreate(dir)
        }
        IniWrite(scrollPercentage, g_WikipediaScrollPositionsFile, "Positions", url)
        return true
    } catch Error as err {
        return false
    }
}

; Load saved scroll position for a Wikipedia article URL
; Returns percentage (0.0 to 1.0) instead of absolute pixels
LoadWikipediaScrollPosition(url) {
    global g_WikipediaScrollPositionsFile
    try {
        if (url = "") {
            return 0.0
        }
        ; Normalize URL to match save format - remove trailing slashes and fragments
        url := RegExReplace(url, "/#.*$", "")
        url := RegExReplace(url, "/+$", "")
        scrollPos := IniRead(g_WikipediaScrollPositionsFile, "Positions", url, "0")
        scrollPercentage := Float(scrollPos)
        return scrollPercentage
    } catch Error as err {
        return 0.0
    }
}

; Handler for character key press
HandleWikipediaChar(char) {
    global g_WikipediaSelectorActive, g_WikipediaItems

    ; Only process if selector is active
    if (!g_WikipediaSelectorActive) {
        return
    }

    ; Find the item for this character
    item := ""
    for i, itm in g_WikipediaItems {
        if (itm.char = char) {
            item := itm
            break
        }
    }

    if (item) {
        ; Cleanup first (closes GUI, disables hotkeys)
        CleanupWikipediaSelector()

        ; If item has a URL, open it in Chrome in a new window
        if (item.url != "") {
            Run "chrome.exe --new-window " item.url
            ; Wait for the window to appear and become active
            WinWait("ahk_exe chrome.exe", , 5)
            Sleep(500)  ; Give the page a moment to start loading

            ; Wait for the page to load (check for Wikipedia in title)
            SetTitleMatchMode 2
            if WinWait("Wikipedia", , 10) {
                WinActivate("Wikipedia")
                WinWaitActive("Wikipedia", , 5)
                Sleep(1500)  ; Additional wait for page to stabilize

                ; Try to restore scroll position (only if on Monitor 3)
                restoreBanner := ""
                try {
                    if (!IsWindowOnMonitor3()) {
                        return
                    }
                    savedPercentage := LoadWikipediaScrollPosition(item.url)
                    if (savedPercentage > 0.0) {
                        ; Show banner to inform user that scroll position is being restored
                        restoreBanner := CreateCenteredBanner_Launchers("Restoring scroll position... Please wait",
                            "3772FF", "FFFFFF", 10, 178, 180)

                        uia := UIA_Browser("ahk_exe chrome.exe")
                        ; Wait a bit more for page to be fully ready
                        Sleep(500)
                        ; Get current document height to calculate pixel position
                        docHeight := uia.JSReturnThroughClipboard("document.documentElement.scrollHeight")
                        if (docHeight != "" && docHeight != "undefined" && docHeight != "null") {
                            docHeightFloat := Float(docHeight)
                            if (docHeightFloat > 0) {
                                targetScrollY := savedPercentage * docHeightFloat
                                uia.JSExecute("window.scrollTo(0, " . Round(targetScrollY) . ");")
                                Sleep(500)  ; Longer wait after scroll to check result

                                ; Update banner to show success
                                try {
                                    if (IsObject(restoreBanner) && restoreBanner.Hwnd) {
                                        restoreBanner.Controls[1].Text := "Scroll position restored!"
                                        Sleep(1000)  ; Show success message for 1 second
                                    }
                                } catch {
                                }

                                ; Hide banner after restoration
                                try {
                                    if (IsObject(restoreBanner) && restoreBanner.Hwnd) {
                                        Sleep(500)  ; Brief delay before hiding
                                        restoreBanner.Destroy()
                                    }
                                } catch {
                                }

                                Sleep(200)  ; Brief wait after scroll
                            }
                        }
                    }
                } catch Error as err {
                    ; Hide banner on error
                    try {
                        if (IsObject(restoreBanner) && restoreBanner.Hwnd) {
                            restoreBanner.Destroy()
                        }
                    } catch {
                    }
                }
            }
        }
        ; Items 2-5 have no URL, so no action is taken
    }
}

; Factory function to create a handler that properly captures the character
CreateWikipediaCharHandler(char) {
    ; Return a function that captures the char value at creation time
    return (*) => HandleWikipediaChar(char)
}

; Handler for Escape key
HandleWikipediaEscape(*) {
    global g_WikipediaSelectorActive
    if (g_WikipediaSelectorActive) {
        CleanupWikipediaSelector()
    }
}

; Cleanup Wikipedia selector
CleanupWikipediaSelector() {
    global g_WikipediaSelectorActive, g_WikipediaSelectorGui, g_WikipediaSelectorHandlers

    ; Disable active flag
    g_WikipediaSelectorActive := false

    ; Disable all character hotkeys
    for handler in g_WikipediaSelectorHandlers {
        try {
            char := handler.char
            Hotkey(char, "Off")
        } catch {
            ; Silently ignore errors
        }
    }

    ; Disable Escape hotkey
    try {
        Hotkey("Escape", "Off")
    } catch {
        ; Ignore
    }

    ; Clear handlers array
    g_WikipediaSelectorHandlers := []

    ; Close and destroy GUI
    if (IsObject(g_WikipediaSelectorGui)) {
        try {
            g_WikipediaSelectorGui.Destroy()
        } catch {
            ; Ignore
        }
        g_WikipediaSelectorGui := false
    }
}

; Show Wikipedia selector GUI
ShowWikipediaSelector() {
    global g_WikipediaSelectorGui, g_WikipediaSelectorActive, g_WikipediaSelectorHandlers
    global g_WikipediaItems

    ; Close existing GUI if open
    if (g_WikipediaSelectorActive && IsObject(g_WikipediaSelectorGui)) {
        CleanupWikipediaSelector()
        Sleep 50
    }

    ; Get monitor dimensions early for responsive sizing
    activeWin := 0
    try {
        activeWin := WinGetID("A")
    } catch {
        activeWin := 0
    }

    ; Default to primary monitor work area
    MonitorGetWorkArea(1, &monitorLeft, &monitorTop, &monitorRight, &monitorBottom)
    monitorWidth := monitorRight - monitorLeft
    monitorHeight := monitorBottom - monitorTop

    ; If we have an active window, find which monitor contains its center
    if (activeWin && activeWin != 0) {
        rect := Buffer(16, 0)
        if (DllCall("GetWindowRect", "ptr", activeWin, "ptr", rect)) {
            ; Calculate window center
            winLeft := NumGet(rect, 0, "int")
            winTop := NumGet(rect, 4, "int")
            winRight := NumGet(rect, 8, "int")
            winBottom := NumGet(rect, 12, "int")

            centerX := winLeft + (winRight - winLeft) // 2
            centerY := winTop + (winBottom - winTop) // 2

            ; Find which monitor contains the window center
            monitorCount := MonitorGetCount()
            loop monitorCount {
                idx := A_Index
                MonitorGetWorkArea(idx, &l, &t, &r, &b)
                if (centerX >= l && centerX <= r && centerY >= t && centerY <= b) {
                    monitorLeft := l
                    monitorTop := t
                    monitorRight := r
                    monitorBottom := b
                    monitorWidth := r - l
                    monitorHeight := b - t
                    break
                }
            }
        }
    }

    ; Create GUI
    ; Create non-activating GUI so PowerToys Command Palette stays open
    g_WikipediaSelectorGui := Gui("+AlwaysOnTop +ToolWindow +E0x08000000", "Wikipedia Articles")
    ; Use slightly smaller font for better fit on small monitors
    fontSize := (monitorHeight < 800) ? 9 : 10
    g_WikipediaSelectorGui.SetFont("s" . fontSize, "Segoe UI")
    g_WikipediaSelectorGui.MarginX := 10
    g_WikipediaSelectorGui.MarginY := 5

    ; Build display text
    displayText := ""
    for i, item in g_WikipediaItems {
        displayText .= "[" . item.char . "] > " . item.title . "`n"
    }
    displayText .= "`nPress Escape to cancel."

    ; Calculate text control height based on actual content (number of lines)
    lineCount := 1  ; Start at 1 (first line doesn't have a newline before it)
    loop parse, displayText, "`n" {
        lineCount++
    }
    ; Calculate height: ~16 pixels per line
    lineHeight := 16
    textControlHeight := lineCount * lineHeight
    ; Ensure minimum and maximum bounds
    minHeight := 150
    maxHeightPercent := (monitorHeight < 800) ? 0.90 : 0.75
    maxHeight := Floor(monitorHeight * maxHeightPercent)
    if (textControlHeight < minHeight)
        textControlHeight := minHeight
    if (textControlHeight > maxHeight)
        textControlHeight := maxHeight

    ; Make width responsive to monitor size
    baseWidth := (monitorWidth < 1200) ? 500 : 600
    textControlWidth := baseWidth - 20  ; Account for margins

    ; Enable vertical scrolling for long content
    g_WikipediaSelectorGui.AddEdit("w" . textControlWidth . " h" . textControlHeight . " ReadOnly VScroll", displayText
    )

    ; Add Close button (set as default so it gets focus, not the Edit control)
    closeBtn := g_WikipediaSelectorGui.AddButton("w100 Default Center", "Close")
    closeBtn.OnEvent("Click", (*) => CleanupWikipediaSelector())

    ; Calculate total height: margins + text control + button + spacing
    totalHeight := 10 + textControlHeight + 40 + 10  ; margins + content + button + spacing
    guiWidth := baseWidth

    ; Calculate center position for the GUI with margins
    marginX := 20  ; Horizontal margin from screen edges
    marginY := 20  ; Vertical margin from screen edges
    guiX := monitorLeft + (monitorWidth - guiWidth) // 2
    guiY := monitorTop + (monitorHeight - totalHeight) // 2

    ; Ensure the GUI stays within monitor bounds with margins
    if (guiX < monitorLeft + marginX)
        guiX := monitorLeft + marginX
    if (guiY < monitorTop + marginY)
        guiY := monitorTop + marginY
    if (guiX + guiWidth > monitorLeft + monitorWidth - marginX)
        guiX := monitorLeft + monitorWidth - guiWidth - marginX
    if (guiY + totalHeight > monitorTop + monitorHeight - marginY)
        guiY := monitorTop + monitorHeight - totalHeight - marginY

    ; Show GUI centered on the active window's monitor
    g_WikipediaSelectorGui.Show("NA w" . guiWidth . " h" . totalHeight . " x" . guiX . " y" . guiY)

    ; Set active flag
    g_WikipediaSelectorActive := true

    ; Clear handlers array
    g_WikipediaSelectorHandlers := []

    ; Enable hotkeys for characters 1-5
    for item in g_WikipediaItems {
        char := item.char
        ; Use factory function to create handler with properly captured char value
        handler := CreateWikipediaCharHandler(char)

        ; Store handler for cleanup
        g_WikipediaSelectorHandlers.Push({ char: char, handler: handler })

        ; Enable hotkey
        try {
            Hotkey(char, handler, "On")
        } catch {
            ; Silently ignore if we can't create hotkey for this character
        }
    }

    ; Enable Escape hotkey
    Hotkey("Escape", HandleWikipediaEscape, "On")
}

; =============================================================================
; Open/Activate Wikipedia
; Hotkey: Win+Alt+Shift+K
; =============================================================================
#!+k::
{
    SetTitleMatchMode 2
    if WinExist("Wikipedia") {
        WinActivate
        WinWaitActive("Wikipedia", , 2)
        CenterMouse()

        ; Enable focus mode to darken other monitors
        EnableFocusMode()

        ; Wait briefly for page to stabilize
        Sleep(500)

        ; Try to restore scroll position (only if on Monitor 3)
        restoreBanner := ""
        try {
            if (!IsWindowOnMonitor3()) {
                return
            }
            url := GetWikipediaURL()
            if (url != "") {
                savedPercentage := LoadWikipediaScrollPosition(url)
                if (savedPercentage > 0.0) {
                    ; Show banner to inform user that scroll position is being restored
                    restoreBanner := CreateCenteredBanner_Launchers("Restoring scroll position... Please wait",
                        "3772FF", "FFFFFF", 10, 178, 180)

                    uia := UIA_Browser("ahk_exe chrome.exe")
                    Sleep(500)  ; Brief wait for page to be ready
                    ; Get current document height to calculate pixel position
                    docHeight := uia.JSReturnThroughClipboard("document.documentElement.scrollHeight")
                    if (docHeight != "" && docHeight != "undefined" && docHeight != "null") {
                        docHeightFloat := Float(docHeight)
                        if (docHeightFloat > 0) {
                            targetScrollY := savedPercentage * docHeightFloat
                            uia.JSExecute("window.scrollTo(0, " . Round(targetScrollY) . ");")
                            Sleep(500)  ; Longer wait after scroll to check result

                            ; Update banner to show success
                            try {
                                if (IsObject(restoreBanner) && restoreBanner.Hwnd) {
                                    restoreBanner.Controls[1].Text := "Scroll position restored!"
                                    Sleep(1000)  ; Show success message for 1 second
                                }
                            } catch {
                            }

                            ; Hide banner after restoration
                            try {
                                if (IsObject(restoreBanner) && restoreBanner.Hwnd) {
                                    Sleep(500)  ; Brief delay before hiding
                                    restoreBanner.Destroy()
                                }
                            } catch {
                            }

                            Sleep(200)  ; Brief wait after scroll
                        }
                    }
                }
            }
        } catch Error as err {
            ; Hide banner on error
            try {
                if (IsObject(restoreBanner) && restoreBanner.Hwnd) {
                    restoreBanner.Destroy()
                }
            } catch {
            }
        }
    } else {
        ShowWikipediaSelector()
    }
}

; =============================================================================
; Pomodoro Timer System - Local Timer with CSV Logging
; Hotkey: Win+Alt+Shift+9
; =============================================================================

; Global variables for Pomodoro timer management
global g_PomodoroTimer := false
global g_ChimeTimer := false
global g_ChimeStopTimer := false
global g_PomodoroOverlay := false
global g_PomodoroLogFile := A_ScriptDir "\data\pomodoro_log.csv"

; Show water bottle image overlay as hydration reminder
ShowWaterBottleOverlay() {
    imagePath := ""
    for name in ["water-bottle.jpg"] {
        candidate := A_ScriptDir "\pictures\" name
        if FileExist(candidate) {
            imagePath := candidate
            break
        }
    }

    overlay := Gui()
    overlay.Opt("+AlwaysOnTop -Caption +ToolWindow +E0x20")
    if (imagePath != "") {
        overlay.Add("Picture", "w240 h240 Center", imagePath)
    } else {
        overlay.SetFont("s36", "Segoe UI")
        overlay.Add("Text", "Center cRed", "Pomodoro")
        overlay.BackColor := "FFFFFF"
    }
    overlay.Show("AutoSize Center")
    return overlay
}

; Log Pomodoro session to CSV file
LogPomodoroSession() {
    global g_PomodoroLogFile
    ; Ensure data directory exists
    SplitPath(g_PomodoroLogFile, , &dir)
    if (dir != "" && !DirExist(dir)) {
        DirCreate(dir)
    }

    ; Check if file exists, if not create with headers
    if (!FileExist(g_PomodoroLogFile)) {
        FileAppend("Date,Time`n", g_PomodoroLogFile)
    }

    ; Get current date and time
    currentDate := FormatTime(, "yyyy/MM/dd")
    currentTime := FormatTime(, "HH:mm")

    ; Append entry to CSV
    FileAppend(currentDate . "," . currentTime . "`n", g_PomodoroLogFile)
}

; Play chime callback - plays sound every 1 second with multiple methods for maximum audibility
PomodoroChimeCallback(*) {
    ; Play multiple sounds simultaneously for maximum audibility
    ; Method 1: Primary - SoundBeep with high frequency and longer duration (most reliable and audible)
    try {
        SoundBeep(2000, 300)  ; High frequency (2000 Hz) and longer duration (300 ms) for better audibility
    } catch {
    }

    ; Method 2: Also try MessageBeep as additional sound
    try {
        DllCall("User32\MessageBeep", "UInt", 0xFFFFFFFF)
    } catch {
    }

    ; Method 3: Also try system sound as additional alert
    try {
        SoundPlay("*16")  ; System asterisk sound
    } catch {
    }
}

; Stop chime callback - stops both chime timers
PomodoroStopChimeCallback(*) {
    global g_ChimeTimer, g_ChimeStopTimer
    if (g_ChimeTimer) {
        SetTimer(g_ChimeTimer, 0)
        g_ChimeTimer := false
    }
    if (g_ChimeStopTimer) {
        SetTimer(g_ChimeStopTimer, 0)
        g_ChimeStopTimer := false
    }
}

; Auto-hide Pomodoro overlay after 5 seconds
PomodoroHideOverlayCallback(*) {
    global g_PomodoroOverlay
    if (g_PomodoroOverlay && IsObject(g_PomodoroOverlay) && g_PomodoroOverlay.Hwnd) {
        try {
            g_PomodoroOverlay.Destroy()
        } catch {
        }
        g_PomodoroOverlay := false
    }
}

; Play completion chime for specified duration
PlayCompletionChime(durationMs) {
    global g_ChimeTimer, g_ChimeStopTimer

    ; Cancel any existing chime timers
    if (g_ChimeTimer) {
        SetTimer(g_ChimeTimer, 0)
        g_ChimeTimer := false
    }
    if (g_ChimeStopTimer) {
        SetTimer(g_ChimeStopTimer, 0)
        g_ChimeStopTimer := false
    }

    ; Play immediate sound when timer completes (before starting periodic chime)
    PomodoroChimeCallback()

    ; Start chime timer (every 1 second for better audibility)
    g_ChimeTimer := PomodoroChimeCallback
    SetTimer(g_ChimeTimer, 1000)

    ; Set timer to stop chime after duration
    g_ChimeStopTimer := PomodoroStopChimeCallback
    SetTimer(g_ChimeStopTimer, -durationMs)
}

; Handler when Pomodoro timer completes (25 minutes)
OnPomodoroComplete() {
    global g_PomodoroTimer, g_PomodoroOverlay, g_ChimeTimer, g_ChimeStopTimer

    ; Cancel the main timer
    if (g_PomodoroTimer) {
        SetTimer(g_PomodoroTimer, 0)
        g_PomodoroTimer := false
    }

    ; Don't show water bottle image when timer completes - removed per user request

    ; Play 30-second completion chime (plays immediate sound, then every 1 second for 30 seconds)
    ; The chime will continue playing even while the message box is shown
    PlayCompletionChime(30000)

    ; Show completion message box immediately (this is blocking, but chime continues in background)
    result := MsgBox("Pomodoro session complete!`n`nTrigger another Pomodoro?", "Pomodoro Complete",
        "YesNo Icon?")

    ; No overlay to clean up (we don't show it when timer completes)

    ; Stop chime when message box is dismissed (works for both Yes and No)
    if (g_ChimeTimer) {
        SetTimer(g_ChimeTimer, 0)
        g_ChimeTimer := false
    }
    if (g_ChimeStopTimer) {
        SetTimer(g_ChimeStopTimer, 0)
        g_ChimeStopTimer := false
    }

    ; If user wants to trigger another Pomodoro, start it
    if (result = "Yes") {
        StartPomodoroTimer()
    }
}

; Start a new Pomodoro timer session
StartPomodoroTimer() {
    global g_PomodoroTimer, g_PomodoroOverlay
    ; Cancel any existing timer
    if (g_PomodoroTimer) {
        SetTimer(g_PomodoroTimer, 0)
        g_PomodoroTimer := false
    }

    ; Log the session start
    LogPomodoroSession()

    ; Show water bottle image overlay
    if (g_PomodoroOverlay && IsObject(g_PomodoroOverlay) && g_PomodoroOverlay.Hwnd) {
        try {
            g_PomodoroOverlay.Destroy()
        } catch {
        }
    }
    g_PomodoroOverlay := ShowWaterBottleOverlay()

    ; Auto-hide overlay after 5 seconds (like original behavior)
    SetTimer(PomodoroHideOverlayCallback, -5000)

    ; Set up 25-minute (1,500,000 ms) completion timer
    g_PomodoroTimer := OnPomodoroComplete
    SetTimer(g_PomodoroTimer, -1500000)
}

; =============================================================================
; Pomodoro Timer - Hotkey: Win+Alt+Shift+9
; =============================================================================
#!+9::
{
    StartPomodoroTimer()
}

; =============================================================================
; Manual Wikipedia Scroll Position Save Function
; Can be called manually to save the current scroll position
; =============================================================================
SaveWikipediaScrollPositionManually() {
    ; Check if Wikipedia window is currently active
    if (!WinActive("ahk_exe chrome.exe") || !InStr(WinGetTitle("A"), "Wikipedia")) {
        return false
    }

    ; Check if window is on Monitor 3
    isOnMonitor3 := IsWindowOnMonitor3()
    if (!isOnMonitor3) {
        return false
    }

    ; Show banner to inform user that scroll position is being saved
    saveBanner := CreateCenteredBanner_Launchers("Saving scroll position... Please wait", "3772FF", "FFFFFF", 24, 178)

    try {
        url := GetWikipediaURL()
        if (url = "") {
            return false
        }

        uia := UIA_Browser("ahk_exe chrome.exe")
        scrollY := uia.JSReturnThroughClipboard("window.pageYOffset")

        ; Get document height to calculate percentage
        docHeight := uia.JSReturnThroughClipboard("document.documentElement.scrollHeight")

        ; Convert to numbers and calculate percentage
        if (scrollY != "" && scrollY != "undefined" && scrollY != "null" && docHeight != "" && docHeight !=
            "undefined" && docHeight != "null") {
            scrollYFloat := Float(scrollY)
            docHeightFloat := Float(docHeight)
            if (scrollYFloat >= 0 && docHeightFloat > 0) {
                scrollPercentage := scrollYFloat / docHeightFloat
                ; Clamp to valid range
                if (scrollPercentage > 1.0) {
                    scrollPercentage := 1.0
                }
                saved := SaveWikipediaScrollPosition(url, scrollPercentage)
                if (saved) {
                    ; Update banner to show success
                    try {
                        if (IsObject(saveBanner) && saveBanner.Hwnd) {
                            saveBanner.Controls[1].Text := "Scroll position saved!"
                            Sleep(1000)  ; Show success message for 1 second
                        }
                    } catch {
                    }
                    return true
                }
            }
        }
    } catch Error as err {
        ; Silent fail
    } finally {
        ; Always hide the banner after save operation completes
        try {
            if (IsObject(saveBanner) && saveBanner.Hwnd) {
                Sleep(500)  ; Brief delay before hiding
                saveBanner.Destroy()
            }
        } catch {
        }
    }
    return false
}

; =============================================================================
; Helper function to center mouse on the active window
; =============================================================================
CenterMouse() {
    Sleep(200)
    Send("#!+q")
}

; =============================================================================
; Send specific key combinations
; Hotkey: Win+Alt+Shift+.
; =============================================================================
#!+.::
{
    Sleep(100)
    Send("^c")
    Sleep(200)
    Send("!v")
    Sleep(700)
    Send("!q")
    Sleep(200)
    Send("{Esc}")
}

; =============================================================================
; Centered banner helper (AppLaunchers aesthetic)
; =============================================================================
CreateCenteredBanner_Launchers(message, bgColor := "be4747", fontColor := "FFFFFF", fontSize := 24, alpha := 178, width :=
    500) {
    bGui := Gui()
    bGui.Opt("+AlwaysOnTop -Caption +ToolWindow")
    bGui.BackColor := bgColor
    bGui.SetFont("s" . fontSize . " c" . fontColor . " Bold", "Segoe UI")
    bGui.Add("Text", "w" . width . " Center", message)

    activeWin := WinGetID("A")
    if (activeWin) {
        WinGetPos(&winX, &winY, &winW, &winH, activeWin)
    } else {
        workArea := SysGet.MonitorWorkArea(SysGet.MonitorPrimary)
        winX := workArea.Left, winY := workArea.Top, winW := workArea.Right - workArea.Left, winH := workArea.Bottom -
            workArea.Top
    }

    bGui.Show("AutoSize Hide")
    guiW := 0, guiH := 0
    bGui.GetPos(, , &guiW, &guiH)

    guiX := winX + (winW - guiW) / 2
    guiY := winY + (winH - guiH) / 2
    bGui.Show("x" . Round(guiX) . " y" . Round(guiY) . " NA")
    WinSetTransparent(alpha, bGui)
    return bGui
}

; =============================================================================
; Initialize Wikipedia scroll position auto-save timer - REMOVED
; =============================================================================
; Auto-save timer removed - now using manual save via Shift keys.ahk shortcut
