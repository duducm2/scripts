; =============================================================================
; Shift keys module: hotif_clipangel.ahk
; ClipAngel hotkeys and filter selector
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

;-------------------------------------------------------------------
; ClipAngel Shortcuts
;-------------------------------------------------------------------

; Global variables for ClipAngel filter selector
global g_ClipAngelFilterSelectorGui := false
global g_ClipAngelFilterSelectorActive := false
global g_ClipAngelFilterCharMap := Map()  ; Maps character to file type info
global g_ClipAngelFilterHotkeyHandlers := []  ; Store hotkey handlers for cleanup

; File type list (filtered to essential types only)
global g_ClipAngelFileTypes := [{ name: "img", index: 1, navKey: "I", navCount: 1 }, { name: "file", index: 2, navKey: "F",
    navCount: 1 }, { name: "text", index: 3, navKey: "T", navCount: 1 }, { name: "*url", index: 10, navKey: "*",
        navCount: 5 }, { name: "*filename", index: 13, navKey: "*", navCount: 8 }
]

; Character sequence for file type assignment (5 types)
global g_ClipAngelFilterCharSequence := ["1", "2", "3", "4", "5"]

#HotIf WinActive("ahk_exe ClipAngel.exe")

; Shift + C : Select filtered content and copy
+c:: {
    Send "{Tab}"
    Sleep 100
    Send "{Tab}"
    Sleep 300
    Send "^a"  ; Select all
    Sleep 100
    Send "^c"  ; Copy
    Sleep 100
    Send "{F10}"
    Sleep 50
    Send "{Escape}"
    Sleep 50
    Send "^v"
}

; Shift + T : Switch focus between list and text (Toggle) (F10)
+t:: Send "{F10}"

; Shift + D : Delete all non-favorite (Ctrl+Alt+K)
+d:: Send "^!k"

; Shift + X : Clear filters (F7)
+x:: Send "{F7}"

; Shift + F : Mark as favorite (Alt+Q)
+f:: Send "!q"

; Shift + U : Unmark as favorite (Unmark) (Alt+W)
+u:: Send "!w"

; Shift + E : Edit text (F4)
+e:: Send "{F4}"

; Shift + S : Save as file (Ctrl+S)
+s:: Send "^s"

; Shift + M : Merge clips
+m:: Send "^!j"

; Shift + I : Clip > Import clips (opens file picker; keep ClipAngel open)
+i:: {
    ClipAngel_WaitChordModifiersReleased()
    ClipAngel_ReleaseChordModifiersForSend()
    if !ClipAngel_InvokeImportClips()
        ShowCenteredOverlay_Utils("❌ Clip Angel Import clips failed", 1500, BANNER_ACCENT_ERROR)
}

; Alt + 1–5 : Move down (N-1), Enter to paste selected clip into prior text field, then minimize.
!1:: ClipAngel_SelectClipPasteThenMinimize(0)
!2:: ClipAngel_SelectClipPasteThenMinimize(1)
!3:: ClipAngel_SelectClipPasteThenMinimize(2)
!4:: ClipAngel_SelectClipPasteThenMinimize(3)
!5:: ClipAngel_SelectClipPasteThenMinimize(4)

; Enter : paste focused list clip into prior text field, then minimize (skip when editing filters/preview).
$Enter:: {
    if ClipAngel_IsListPasteEnterContext()
        ClipAngel_SelectClipPasteThenMinimize(0)
    else
        Send "{Enter}"
}

; Alt + Enter : Paste file — Clip > Paste > Paste file; then minimize (process stays running).
!Enter:: {
    ClipAngel_WaitChordModifiersReleased()
    ClipAngel_ReleaseChordModifiersForSend()
    try {
        if !ClipAngel_InvokePasteEnter()
            ShowCenteredOverlay_Utils("❌ Clip Angel Paste file failed", 1500, BANNER_ACCENT_ERROR)
    } finally {
        ClipAngel_CloseAndRestoreFocus(0)
    }
}

; Ctrl+Alt+B / Ctrl+Alt+V : native paste + move selection; then minimize
~^!b:: {
    Sleep 100
    ClipAngel_CloseAndRestoreFocus(0)
}
~^!v:: {
    Sleep 100
    ClipAngel_CloseAndRestoreFocus(0)
}

; Ctrl+Enter : native Clip Angel action; then minimize
~^Enter:: {
    Sleep 100
    ClipAngel_CloseAndRestoreFocus(0)
}

; Ctrl + 1–5 : Down N, F10, Select All, Copy, then minimize.
^1:: ClipAngel_SelectClipCopyThenMinimize(0)
^2:: ClipAngel_SelectClipCopyThenMinimize(1)
^3:: ClipAngel_SelectClipCopyThenMinimize(2)
^4:: ClipAngel_SelectClipCopyThenMinimize(3)
^5:: ClipAngel_SelectClipCopyThenMinimize(4)

; =============================================================================
; ClipAngel Filter Selector Functions
; =============================================================================

; Navigate ClipAngel ComboBox to select a specific file type
NavigateClipAngelComboBox(typeIndex) {
    try {
        ; Ensure ClipAngel window is active
        win := WinExist("A")
        WinActivate("ahk_id " win)
        WinWaitActive("ahk_id " win, , 1)
        Sleep 50 ; Brief settle after activation

        root := UIA.ElementFromHandle(win)
        Sleep 50 ; Brief settle for UIA

        ; Find the file type filter ComboBox
        typeFilterCombo := root.FindFirst({ AutomationId: "TypeFilter", Type: 50003 })

        ; Fallback strategies
        if !typeFilterCombo {
            typeFilterCombo := root.FindFirst({ AutomationId: "TypeFilter", Type: "ComboBox" })
        }
        if !typeFilterCombo {
            typeFilterCombo := root.FindFirst({ Name: "all types", Type: 50003 })
        }
        if !typeFilterCombo {
            typeFilterCombo := root.FindFirst({ AutomationId: "TypeFilter" })
        }

        if !typeFilterCombo {
            return ; Silently fail if element not found
        }

        ; Get file type info first to show banner
        if (typeIndex < 0 || typeIndex >= g_ClipAngelFileTypes.Length) {
            return ; Invalid index
        }

        fileType := g_ClipAngelFileTypes[typeIndex + 1] ; +1 because arrays are 1-indexed

        ; Format display name for banner
        displayName := fileType.name
        if (displayName = "img") {
            displayName := "Image"
        } else if (displayName = "file") {
            displayName := "File"
        } else if (displayName = "text") {
            displayName := "Text"
        } else if (displayName = "*url") {
            displayName := "URL"
        } else if (displayName = "*filename") {
            displayName := "Filename"
        }

        ; Show banner notification
        ShowCenteredOverlay_Utils("📌 Selecting: " . displayName, 800, BANNER_ACCENT_INTERMEDIATE)

        ; Set focus and click to open dropdown
        try {
            typeFilterCombo.SetFocus()
            Sleep 30
        } catch {
            ; Continue if SetFocus fails
        }

        typeFilterCombo.Click()
        Sleep 150 ; Reduced delay after click

        ; Ensure focus is maintained
        try {
            typeFilterCombo.SetFocus()
            Sleep 100 ; Reduced delay after SetFocus
        } catch {
            typeFilterCombo.Click()
            Sleep 100 ; Reduced delay for retry
        }

        ; Reset to "all types" first - use Home key to physically anchor at top
        Send "{Home}"
        Sleep 200 ; Reduced delay after Home

        ; Navigate to target type
        if (fileType.navKey = "*") {
            ; For asterisk items, send asterisk multiple times
            loop fileType.navCount {
                Send "*"
                Sleep 100 ; Reduced delay between keystrokes
            }
        } else {
            ; For letter-based items, send the letter
            Send fileType.navKey
            Sleep 100 ; Reduced delay after navigation character
        }

        Sleep 100 ; Reduced pause after navigation

        ; Send Tab to confirm selection (NOT Enter)
        Send "{Tab}"
        Sleep 100 ; Reduced delay after TAB

        ; Return focus to clipboard list using Shift+T logic
        Send "{F10}"
        Sleep 100 ; Reduced delay after F10

        ; Send CTRL+HOME to force focus to very first item in sidebar list
        Send "^{Home}"
        Sleep 50 ; Reduced delay after CTRL+HOME
    } catch Error as e {
        ; Silently fail on error
    }
}

; Handler for character key press in filter selector
HandleClipAngelFilterChar(char) {
    global g_ClipAngelFilterSelectorActive, g_ClipAngelFilterCharMap

    ; Only process if selector is active
    if (!g_ClipAngelFilterSelectorActive) {
        return
    }

    ; Get file type info for this character
    fileTypeInfo := g_ClipAngelFilterCharMap.Get(char, "")
    if (fileTypeInfo = "") {
        ; Try lowercase if uppercase
        fileTypeInfo := g_ClipAngelFilterCharMap.Get(StrLower(char), "")
    }

    if (fileTypeInfo != "") {
        ; Cleanup selector first (closes GUI, disables hotkeys)
        CleanupClipAngelFilterSelector()

        ; Small delay to ensure ClipAngel window has focus
        Sleep 150

        ; Navigate to selected file type - use array index, not ComboBox index
        ; Find the array index by searching for matching fileType
        arrayIndex := -1
        for idx, ft in g_ClipAngelFileTypes {
            if (ft.name = fileTypeInfo.name) {
                arrayIndex := idx - 1 ; Convert to 0-based index
                break
            }
        }

        if (arrayIndex >= 0) {
            NavigateClipAngelComboBox(arrayIndex)
        }
    }
}

; Factory function to create a handler that properly captures the character
CreateClipAngelFilterCharHandler(char) {
    return (*) => HandleClipAngelFilterChar(char)
}

; Handler for Escape key
HandleClipAngelFilterEscape(*) {
    global g_ClipAngelFilterSelectorActive
    if (g_ClipAngelFilterSelectorActive) {
        CleanupClipAngelFilterSelector()
    }
}

; Cleanup function for ClipAngel filter selector
CleanupClipAngelFilterSelector() {
    global g_ClipAngelFilterSelectorActive, g_ClipAngelFilterSelectorGui, g_ClipAngelFilterHotkeyHandlers
    global g_ClipAngelFilterCharMap

    ; Disable active flag
    g_ClipAngelFilterSelectorActive := false

    ; Disable all character hotkeys
    for handler in g_ClipAngelFilterHotkeyHandlers {
        try {
            char := handler.char
            ; Handle special VK codes
            if (char = ",") {
                Hotkey("vkBC", "Off")
            } else if (char = ".") {
                Hotkey("vkBE", "Off")
            } else {
                Hotkey(char, "Off")
                ; Also disable uppercase for lowercase letters
                if (RegExMatch(char, "^[a-z]$")) {
                    Hotkey(StrUpper(char), "Off")
                }
            }
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
    g_ClipAngelFilterHotkeyHandlers := []

    ; Close and destroy GUI
    if (IsObject(g_ClipAngelFilterSelectorGui)) {
        try {
            g_ClipAngelFilterSelectorGui.Destroy()
        } catch {
            ; Ignore
        }
        g_ClipAngelFilterSelectorGui := false
    }

    ; Clear char map
    g_ClipAngelFilterCharMap := Map()
}

; Show ClipAngel filter selector GUI
ShowClipAngelFilterSelector() {
    global g_ClipAngelFilterSelectorGui, g_ClipAngelFilterSelectorActive, g_ClipAngelFilterCharMap
    global g_ClipAngelFilterHotkeyHandlers, g_ClipAngelFileTypes, g_ClipAngelFilterCharSequence

    ; Close existing GUI if open
    if (g_ClipAngelFilterSelectorActive && IsObject(g_ClipAngelFilterSelectorGui)) {
        CleanupClipAngelFilterSelector()
        Sleep 50
    }

    ; Build character mapping
    g_ClipAngelFilterCharMap := Map()
    charIndex := 0
    for fileType in g_ClipAngelFileTypes {
        if (charIndex < g_ClipAngelFilterCharSequence.Length) {
            char := g_ClipAngelFilterCharSequence[charIndex + 1] ; +1 for 1-indexed array
            g_ClipAngelFilterCharMap[char] := fileType
            charIndex++
        }
    }

    if (g_ClipAngelFilterCharMap.Count = 0) {
        TrayTip("ClipAngel Filter Selector", "No file types found.", "IconX")
        SetTimer(() => TrayTip(), -5000)
        return
    }

    ; Get monitor dimensions
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
            winLeft := NumGet(rect, 0, "int")
            winTop := NumGet(rect, 4, "int")
            winRight := NumGet(rect, 8, "int")
            winBottom := NumGet(rect, 12, "int")

            centerX := winLeft + (winRight - winLeft) // 2
            centerY := winTop + (winBottom - winTop) // 2

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
    g_ClipAngelFilterSelectorGui := Gui("+AlwaysOnTop +ToolWindow +E0x08000000", "ClipAngel File Type Filter")
    fontSize := (monitorHeight < 800) ? 9 : 10
    g_ClipAngelFilterSelectorGui.SetFont("s" . fontSize, "Segoe UI")
    g_ClipAngelFilterSelectorGui.MarginX := 10
    g_ClipAngelFilterSelectorGui.MarginY := 5

    ; Build display text
    displayText := "ClipAngel File Type Filter`n`n"
    charIndex := 0
    for fileType in g_ClipAngelFileTypes {
        if (charIndex < g_ClipAngelFilterCharSequence.Length) {
            char := g_ClipAngelFilterCharSequence[charIndex + 1]
            ; Format display name for better readability
            displayName := fileType.name
            if (displayName = "img") {
                displayName := "Image"
            } else if (displayName = "file") {
                displayName := "File"
            } else if (displayName = "text") {
                displayName := "Text"
            } else if (displayName = "*url") {
                displayName := "URL"
            } else if (displayName = "*filename") {
                displayName := "Filename"
            }
            displayText .= "[" . char . "] " . displayName . "`n"
            charIndex++
        }
    }

    ; Calculate GUI size
    baseWidth := 300
    textControlHeight := Min(400, (g_ClipAngelFileTypes.Length * 20) + 60)
    textControlWidth := baseWidth - 20

    ; Add text control with display
    g_ClipAngelFilterSelectorGui.AddEdit("w" . textControlWidth . " h" . textControlHeight . " ReadOnly VScroll",
        displayText)

    ; Add Close button
    closeBtn := g_ClipAngelFilterSelectorGui.AddButton("w100 Default Center", "Close")
    closeBtn.OnEvent("Click", (*) => CleanupClipAngelFilterSelector())

    ; Calculate total height
    totalHeight := 10 + textControlHeight + 40 + 10
    guiWidth := baseWidth

    ; Calculate center position
    marginX := 20
    marginY := 20
    guiX := monitorLeft + (monitorWidth - guiWidth) // 2
    guiY := monitorTop + (monitorHeight - totalHeight) // 2

    ; Ensure GUI stays within monitor bounds
    if (guiX < monitorLeft + marginX)
        guiX := monitorLeft + marginX
    if (guiY < monitorTop + marginY)
        guiY := monitorTop + marginY
    if (guiX + guiWidth > monitorLeft + monitorWidth - marginX)
        guiX := monitorLeft + monitorWidth - guiWidth - marginX
    if (guiY + totalHeight > monitorTop + monitorHeight - marginY)
        guiY := monitorTop + monitorHeight - totalHeight - marginY

    ; Show GUI centered on the active window's monitor
    g_ClipAngelFilterSelectorGui.Show("NA w" . guiWidth . " h" . totalHeight . " x" . guiX . " y" . guiY)

    ; Set active flag
    g_ClipAngelFilterSelectorActive := true

    ; Clear handlers array
    g_ClipAngelFilterHotkeyHandlers := []

    ; Enable hotkeys for all assigned characters
    charIndex := 0
    for fileType in g_ClipAngelFileTypes {
        if (charIndex < g_ClipAngelFilterCharSequence.Length) {
            char := g_ClipAngelFilterCharSequence[charIndex + 1]

            ; Create handler
            handler := CreateClipAngelFilterCharHandler(char)

            ; Store handler for cleanup
            g_ClipAngelFilterHotkeyHandlers.Push({ char: char, handler: handler })

            ; Enable hotkey (both uppercase and lowercase)
            try {
                if (char = ",") {
                    Hotkey("vkBC", handler, "On")
                } else if (char = ".") {
                    Hotkey("vkBE", handler, "On")
                } else {
                    Hotkey(char, handler, "On")
                    if (RegExMatch(char, "^[a-z]$")) {
                        Hotkey(StrUpper(char), handler, "On")
                    }
                }
            } catch {
                ; Silently ignore if we can't create hotkey
            }

            charIndex++
        }
    }

    ; Enable Escape hotkey
    Hotkey("Escape", HandleClipAngelFilterEscape, "On")
}

; Shift + Y : Open file type filter selector (Quick Wizard)
+y:: {
    ; Only show selector if ClipAngel is active
    if WinActive("ahk_exe ClipAngel.exe") {
        ShowClipAngelFilterSelector()
    }
}

#HotIf
; Only Shift keys owns this hook (Utils is shared across processes).
ClipAngel_InitAutoMinimizeOnDeactivate()