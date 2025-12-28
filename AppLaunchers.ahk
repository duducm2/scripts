#Requires AutoHotkey v2.0+
#SingleInstance Force

; -----------------------------------------------------------------------------
; This script consolidates all Application/Website launcher hotkeys.
; -----------------------------------------------------------------------------

; --- Includes ----------------------------------------------------------------
#include %A_ScriptDir%\env.ahk
#include UIA-v2\Lib\UIA.ahk

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

; Wikipedia article items configuration
; Item 1: Taoism
; Items 2-5: Placeholders (no action)
global g_WikipediaItems := [{ char: "1", title: "Taoism", url: "https://en.wikipedia.org/wiki/Taoism" }, { char: "2",
    title: "Placeholder", url: "" }, { char: "3", title: "Placeholder", url: "" }, { char: "4", title: "Placeholder",
        url: "" }, { char: "5", title: "Placeholder", url: "" }
]

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
        CenterMouse()
    } else {
        ShowWikipediaSelector()
    }
}

; =============================================================================
; Open Google Keep Pomodoro list, show overlay, then send keys
; Hotkey: Win+Alt+Shift+9
; =============================================================================
#!+9::
{
    SoundBeep()
    url := "https://keep.google.com/u/0/#LIST/18_9CP4EZyO6i88cnsRR6HgRNlSEohvkKX0PEZXh8mKnj8H_TpfRhG6aUbrCdxQ"
    Run "chrome.exe " url
    WinWait("ahk_exe chrome.exe")
    WinActivate("ahk_exe chrome.exe")

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
    Sleep(5000)

    WinActivate("ahk_exe chrome.exe")
    Sleep(100)
    Send("{Enter}")
    Sleep(250)
    Send("{Up}")
    Sleep(250)
    ClipSaved := ClipboardAll()
    current := FormatTime(, "dd/MM/yyyy HH:mm")
    A_Clipboard := "tomato " current
    ClipWait(1)
    Send("^v")
    Sleep(50)
    A_Clipboard := ClipSaved

    Send("{Escape}")
    Sleep(4000)
    Send("^w")

    ; Open Windows Clock (Alarms & Clock) and start focus session (Pomodoro)
    try {
        Run "ms-clock:"
        ; The Clock window is hosted by ApplicationFrameHost.exe, title typically contains "Relógio" or "Clock"
        SetTitleMatchMode 2
        WinWait("ahk_exe ApplicationFrameHost.exe")
        WinActivate("ahk_exe ApplicationFrameHost.exe")
        Sleep(2000)

        ; Tab through elements until we find the "Iniciar" button, then press Enter
        try {
            ; Start from the beginning of the application
            Send("^{Home}")
            Sleep(300)

            ; Tab through elements until we find "Iniciar" button
            maxTabs := 50  ; Safety limit to prevent infinite loop
            foundButton := false

            loop maxTabs {
                ; Check if current focused element has AutomationId "TimerPlayPauseButton"
                try {
                    focusedElement := UIA.GetFocusedElement()
                    if (focusedElement && focusedElement.AutomationId == "TimerPlayPauseButton") {
                        foundButton := true
                        break
                    }
                } catch {
                    ; Continue if UIA check fails
                }

                ; Tab to next element
                Send("{Tab}")
                Sleep(200)
            }

            ; If we found the button, press Enter to activate it
            if (foundButton) {
                Send("{Enter}")
                Sleep(200)
            }

            ; Close the Clock window (ApplicationFrameHost.exe)
            try {
                WinClose("ahk_exe ApplicationFrameHost.exe")
            }

        } catch {
            ; Fallback: do nothing if UIA not available
        }

    } catch {
        ; Fail silently if UIA not available here
    }

    ; Quick motivation banner
    goOverlay := CreateCenteredBanner_Launchers("GO!", "3772FF", "FFFFFF", 48, 178)
    Sleep(1200)
    try goOverlay.Destroy()

    SoundBeep()
    overlay.Destroy()
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
CreateCenteredBanner_Launchers(message, bgColor := "be4747", fontColor := "FFFFFF", fontSize := 24, alpha := 178) {
    bGui := Gui()
    bGui.Opt("+AlwaysOnTop -Caption +ToolWindow")
    bGui.BackColor := bgColor
    bGui.SetFont("s" . fontSize . " c" . fontColor . " Bold", "Segoe UI")
    bGui.Add("Text", "w500 Center", message)

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
