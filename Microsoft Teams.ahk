#Requires AutoHotkey v2.0+
#SingleInstance Force

; -----------------------------------------------------------------------------
; This script consolidates all Microsoft Teams related hotkeys and functions.
; -----------------------------------------------------------------------------

; --- Includes ----------------------------------------------------------------
#include UIA-v2\Lib\UIA.ahk

; --- Helper Functions --------------------------------------------------------

ActivateWindowWithRetry(hwnd, attempts := 6, waitMs := 500) {
    ; Get original window state to preserve size and prevent unwanted maximization
    originalState := ""
    try {
        originalState := WinGetMinMax(hwnd)
        ; Validate the state value (-1=minimized, 0=normal, 1=maximized)
        if !(originalState = -1 || originalState = 0 || originalState = 1) {
            originalState := ""  ; Reset if invalid state
        }
    } catch {
        originalState := ""  ; Reset on error
    }
    
    ; Multiple strategies to restore and activate window
    Loop attempts {
        ; Strategy 1: Standard restore + activate (only if minimized)
        try {
            if (originalState = -1) {  ; Only restore if window was minimized (-1=minimized, 0=normal, 1=maximized)
                WinRestore(hwnd)
                Sleep 100
            }
            WinActivate(hwnd)
            if WinWaitActive("ahk_id " hwnd, , waitMs/1000) {
                return true
            }
        }
        
        ; Strategy 2: Show window using ShowWindow API (only if minimized)
        try {
            if (originalState = -1) {  ; Only restore if window was minimized (-1=minimized, 0=normal, 1=maximized)
                DllCall("ShowWindow", "Ptr", hwnd, "Int", 9)  ; SW_RESTORE
                Sleep 100
            }
            DllCall("SetForegroundWindow", "Ptr", hwnd)
            if WinWaitActive("ahk_id " hwnd, , waitMs/1000) {
                return true
            }
        }
        
        ; Strategy 3: Force to front using BringWindowToTop
        try {
            DllCall("BringWindowToTop", "Ptr", hwnd)
            Sleep 100
            WinActivate(hwnd)
            if WinWaitActive("ahk_id " hwnd, , waitMs/1000) {
                return true
            }
        }
        
        ; Strategy 4: Alt+Tab simulation to bring window up
        if A_Index = attempts {
            try {
                Send "!{Tab}"
                Sleep 200
                WinActivate(hwnd)
                if WinWaitActive("ahk_id " hwnd, , waitMs/1000) {
                    return true
                }
            }
        }
        
        Sleep 300
    }
    return false
}

ActivateTeamsMeetingWindow() {
    static processes := ["ms-teams.exe", "Teams.exe", "MSTeams.exe"]
    ; Debug: collect all Teams windows for troubleshooting
    allTitles := ""
    foundMeetingWindow := false
    
    for proc in processes {
        for hwnd in WinGetList("ahk_exe " proc) {
            title := WinGetTitle(hwnd)
            allTitles .= "- " . title . " (hwnd: " . hwnd . ")`n"
            if IsTeamsMeetingTitle(title) {
                foundMeetingWindow := true
                allTitles .= "  → MEETING WINDOW DETECTED, attempting activation...`n"
                if ActivateWindowWithRetry(hwnd) {
                    allTitles .= "  → SUCCESS: Window activated`n"
                    return true
                } else {
                    allTitles .= "  → FAILED: Could not activate window`n"
                }
            }
        }
    }
    
    ; Try regex fallback
    if hwnd := WinExist("RegEx)^.*\| Microsoft Teams$") {
        foundMeetingWindow := true
        title := WinGetTitle(hwnd)
        allTitles .= "- REGEX MATCH: " . title . " (hwnd: " . hwnd . ")`n"
        allTitles .= "  → Attempting activation via regex fallback...`n"
        if ActivateWindowWithRetry(hwnd) {
            allTitles .= "  → SUCCESS: Window activated via regex`n"
            return true
        } else {
            allTitles .= "  → FAILED: Could not activate via regex`n"
        }
    }
    
    ; Final fallback: Click on Teams taskbar button
    if foundMeetingWindow {
        allTitles .= "`n→ FINAL ATTEMPT: Clicking Teams taskbar button...`n"
        try {
            ; Try to find and click Teams in taskbar
            if WinExist("ahk_exe ms-teams.exe") || WinExist("ahk_exe Teams.exe") {
                ; Send Win+T to cycle through taskbar and look for Teams
                Send "#t"
                Sleep 200
                ; Try clicking where Teams might be
                Loop 10 {
                    Send "{Right}"
                    Sleep 50
                    if InStr(WinGetTitle("A"), "Teams") {
                        Send "{Enter}"
                        Sleep 500
                        ; Check if we now have an active Teams window
                        if WinActive("ahk_exe ms-teams.exe") || WinActive("ahk_exe Teams.exe") {
                            allTitles .= "  → SUCCESS: Activated via taskbar`n"
                            return true
                        }
                        break
                    }
                }
            }
        }
        allTitles .= "  → FAILED: Taskbar activation failed`n"
    }
    
    ; Show error as banner overlay
    debugMsg := foundMeetingWindow ? 
        "MEETING WINDOW FOUND BUT COULD NOT ACTIVATE" : 
        "NO MEETING WINDOW FOUND"
    ShowCenteredOverlay(WinGetID("A"), debugMsg, 3000)
    return false
}

ActivateTeamsChatWindow() {
    static processes := ["ms-teams.exe", "Teams.exe", "MSTeams.exe"]
    for proc in processes {
        for hwnd in WinGetList("ahk_exe " proc) {
            if IsTeamsChatTitle(title := WinGetTitle(hwnd)) {
                WinActivate(hwnd)
                return true
            }
        }
    }
    if hwnd := WinExist("RegEx)^Chat \| .* \| Microsoft Teams$") {
        WinActivate(hwnd)
        return true
    }
    ; No message box here - just return false
    return false
}

FindListItemContaining(root, text) {
    items := root.FindAll(UIA.CreateCondition({ ControlType: "ListItem" }))
    for item in items {
        if InStr(item.Name, text)
            return item
    }
    return false
}

FindListItemContainingMultiLang(root, textArray) {
    items := root.FindAll(UIA.CreateCondition({ ControlType: "ListItem" }))
    for item in items {
        for text in textArray {
            if InStr(item.Name, text)
                return item
        }
    }
    return false
}

WaitListItem(root, partialName, timeout := 3000) {
    start := A_TickCount
    while (A_TickCount - start < timeout) {
        item := FindListItemContaining(root, partialName)
        if item
            return item
        Sleep 100
    }
    return false
}

WaitListItemMultiLang(root, partialNameArray, timeout := 3000) {
    start := A_TickCount
    while (A_TickCount - start < timeout) {
        item := FindListItemContainingMultiLang(root, partialNameArray)
        if item
            return item
        Sleep 100
    }
    return false
}

IsTeamsMeetingTitle(title) {
    if InStr(title, "Chat |") || InStr(title, "Sharing control bar |")
        return false
    ; Support both English and Portuguese meeting indicators
    if InStr(title, "Microsoft Teams meeting") || InStr(title, "Reunião do Microsoft Teams")
        return true
    return RegExMatch(title, "i)^.*\| Microsoft Teams.*$")
}

IsTeamsChatTitle(title) {
    if InStr(title, "Sharing control bar |") || InStr(title, "Microsoft Teams meeting")
        return false
    return InStr(title, "Chat |") && RegExMatch(title, "i)\| Microsoft Teams$")
}

; --- NEW helper --------------------------------------------------------------
ShowCenteredOverlay(hwndTarget, text, duration := 1500) {
    ; High-contrast centered banner (consistent with other scripts)
    ; Validate target window, fall back to active window, then screen center
    target := hwndTarget
    if !(IsSet(target) && target && WinExist("ahk_id " target)) {
        target := WinGetID("A")
    }
    hasWindow := false
    if target && WinExist("ahk_id " target) {
        try {
            WinGetPos(&wx, &wy, &ww, &wh, target)
            hasWindow := (ww > 0 && wh > 0)
        } catch {
            hasWindow := false
        }
    }

    ov := Gui("+AlwaysOnTop -Caption +ToolWindow")
    ov.BackColor := "3772FF"          ; strong blue
    ov.SetFont("s24 cFFFFFF Bold", "Segoe UI")
    msg := ov.Add("Text", "w500 Center", text)
    ov.Show("AutoSize Hide")          ; measure the GUI first
    ov.GetPos(&gx, &gy, &gw, &gh)

    if hasWindow {
        cx := wx + (ww - gw)//2
        cy := wy + (wh - gh)//2
        ov.Show("x" . cx . " y" . cy . " NA")
    } else {
        ; Screen center fallback (virtual screen across monitors)
        vx := SysGet(76)  ; SM_XVIRTUALSCREEN
        vy := SysGet(77)  ; SM_YVIRTUALSCREEN
        vw := SysGet(78)  ; SM_CXVIRTUALSCREEN
        vh := SysGet(79)  ; SM_CYVIRTUALSCREEN
        cx := vx + (vw - gw)//2
        cy := vy + (vh - gh)//2
        ov.Show("x" . cx . " y" . cy . " NA")
    }

    WinSetTransparent(178, ov)        ; ~70% opacity for visibility
    Sleep duration
    ov.Destroy()
}


; --- Hotkeys & Functions -----------------------------------------------------

; --- Audio feedback helper ---
PlayMicrophoneBeep() {
    ; Play a single short beep to indicate microphone action
    SoundBeep(800, 150)
}

; --- Microphone state verification ---
GetMicrophoneState(hwndTeams, maxRetries := 3) {
    ; Returns: "muted", "unmuted", or "unknown"
    Loop maxRetries {
        try {
            root := UIA.ElementFromHandle(hwndTeams)
            if !root {
                if A_Index < maxRetries
                    Sleep 150
                continue
            }
            
            ; Try automation ID first
            micBtn := root.FindFirst(UIA.CreateCondition({ AutomationId: "microphone-button" }))
            
            ; If automation ID fails, try finding by name patterns
            if !micBtn {
                ; English and Portuguese name patterns for microphone button
                micNamePatterns := [
                    "Microphone", "Mic", "Mute", "Unmute",
                    "Microfone", "Mudo", "Desativar mudo", "Ativar mudo",
                    "Turn on microphone", "Turn off microphone",
                    "Ligar microfone", "Desligar microfone"
                ]
                
                for pattern in micNamePatterns {
                    micBtn := root.FindFirst(UIA.CreateCondition({ Name: pattern }))
                    if micBtn
                        break
                }
            }
            
            if micBtn {
                ; Prefer ToggleState when available
                try {
                    state := micBtn.ToggleState  ; 0=Off, 1=On, 2=Indeterminate
                    if (state = UIA.ToggleState.On)
                        return "muted"          ; Toggle ON => mute active
                    if (state = UIA.ToggleState.Off)
                        return "unmuted"
                }
                ; Fallback: infer from Name (action-based text)
                try name := micBtn.Name
                if name {
                    ; Portuguese patterns
                    if InStr(name, "Desativar mudo") || InStr(name, "Desligar microfone") ; "Disable mute" => currently muted
                        return "muted"
                    if InStr(name, "Ativar mudo") || InStr(name, "Ligar microfone")    ; "Enable mute" => currently unmuted
                        return "unmuted"
                    ; English patterns
                    if InStr(name, "Unmute") || InStr(name, "Turn on microphone")
                        return "muted"
                    if InStr(name, "Mute") || InStr(name, "Turn off microphone")
                        return "unmuted"
                }
            }
        }
        if A_Index < maxRetries
            Sleep 200  ; Wait before retry
    }
    return "unknown"
}

; --- Camera state verification ---
GetCameraState(hwndTeams, maxRetries := 3) {
    ; Returns: "on", "off", or "unknown"
    Loop maxRetries {
        try {
            root := UIA.ElementFromHandle(hwndTeams)
            if !root {
                if A_Index < maxRetries
                    Sleep 150
                continue
            }
            
            ; Try automation ID first
            camBtn := root.FindFirst(UIA.CreateCondition({ AutomationId: "video-button" }))
            
            ; If automation ID fails, try finding by name patterns
            if !camBtn {
                ; English and Portuguese name patterns for camera button
                camNamePatterns := [
                    "Camera", "Video", "Turn on camera", "Turn off camera", "Turn camera on", "Turn camera off",
                    "Câmera", "Vídeo", "Ativar câmera", "Desativar câmera",
                    "Start video", "Stop video", "Iniciar vídeo", "Parar vídeo"
                ]
                
                for pattern in camNamePatterns {
                    camBtn := root.FindFirst(UIA.CreateCondition({ Name: pattern }))
                    if camBtn
                        break
                }
            }
            
            if camBtn {
                ; Prefer ToggleState when available
                try {
                    state := camBtn.ToggleState  ; 0=Off, 1=On, 2=Indeterminate
                    if (state = UIA.ToggleState.On)
                        return "on"
                    if (state = UIA.ToggleState.Off)
                        return "off"
                }
                ; Fallback: infer from Name (action-based text)
                try name := camBtn.Name
                if name {
                    ; Portuguese patterns
                    if InStr(name, "Desativar câmera") || InStr(name, "Parar vídeo") ; "Disable camera" => currently on
                        return "on"
                    if InStr(name, "Ativar câmera") || InStr(name, "Iniciar vídeo")    ; "Enable camera" => currently off
                        return "off"
                    ; English patterns
                    if InStr(name, "Turn off camera") || InStr(name, "Turn camera off") || InStr(name, "Stop video")
                        return "on"
                    if InStr(name, "Turn on camera") || InStr(name, "Turn camera on") || InStr(name, "Start video")
                        return "off"
                }
            }
        }
        if A_Index < maxRetries
            Sleep 200
    }
    return "unknown"
}

; =============================================================================
; Meeting: Toggle Mute
; Hotkey: Win+Alt+Shift+5
; Original File: Microsoft Teams - meeting shortcuts.ahk
; =============================================================================
#!+5:: {
    prev := WinGetID("A")                     ; window you were in
    if !ActivateTeamsMeetingWindow()
        return

    hwndTeams := WinGetID("A")
    ; Get initial state
    initialState := GetMicrophoneState(hwndTeams)
    
    ; Toggle microphone once
    Send "^+m"
    Sleep 600
    
    ; Verify the state changed (check only; do not re-toggle)
    finalState := "unknown"
    Loop 3 {
        Sleep 250
        finalState := GetMicrophoneState(hwndTeams)
        if (finalState != "unknown" && finalState != initialState)
            break
    }
    
    ; On success, play single beep and show overlay
    if (finalState != "unknown" && finalState != initialState) {
        PlayMicrophoneBeep()
        WinActivate(prev)
        if finalState = "muted"
            ShowCenteredOverlay(prev, "MIC MUTED")
        else
            ShowCenteredOverlay(prev, "MIC UNMUTED")
        return
    }
    
    ; On failure, show an error banner and do not beep
    WinActivate(prev)
    ShowCenteredOverlay(prev, "MICROPHONE STATE UNKNOWN", 3000)
}

; =============================================================================
; Meeting: Toggle Camera
; Hotkey: Win+Alt+Shift+4
; Original File: Microsoft Teams - meeting shortcuts.ahk
; =============================================================================
#!+4:: {
    prev := WinGetID("A")
    if !ActivateTeamsMeetingWindow()
        return

    hwndTeams := WinGetID("A")

    ; Get initial camera state
    initialState := GetCameraState(hwndTeams)

    ; Toggle camera once
    Send "^+o"
    Sleep 600

    ; Verify the state changed (check only; do not re-toggle)
    finalState := "unknown"
    Loop 3 {
        Sleep 250
        finalState := GetCameraState(hwndTeams)
        if (finalState != "unknown")
            break
    }

    WinActivate(prev)

    if (finalState = "on" || finalState = "off") {
        PlayMicrophoneBeep()
        if finalState = "on"
            ShowCenteredOverlay(prev, "CAMERA ON")
        else
            ShowCenteredOverlay(prev, "CAMERA OFF")
        return
    }

    ShowCenteredOverlay(prev, "CAMERA STATE UNKNOWN", 3000)
}


; =============================================================================
; Meeting: Toggle Screen Share  (Win Alt Shift T)
; =============================================================================
#!+t:: {
    prev := WinGetID("A")                 ; remember the window you were in
    if !ActivateTeamsMeetingWindow()
        return

    hwndTeams := WinGetID("A")            ; Teams meeting window
    root := UIA.ElementFromHandle(hwndTeams)
    if !root
        return

    ; --- perform the normal sharing workflow ---
    ; Support both English and Portuguese
    windowListTexts := ["Opens list of", "Abre a lista de"]
    listItem := FindListItemContainingMultiLang(root, windowListTexts)
    if listItem {
        listItem.Invoke()
    } else {
        ; Try automation ID first
        shareBtn := root.FindFirst(UIA.CreateCondition({ AutomationId: "share-button" }))
        
        ; If automation ID fails, try finding by name patterns
        if !shareBtn {
            ; English and Portuguese name patterns for share button
            shareNamePatterns := [
                "Share", "Share content", "Share screen", "Start sharing",
                "Compartilhar", "Compartilhar conteúdo", "Compartilhar tela", "Iniciar compartilhamento",
                "Present", "Present screen", "Apresentar", "Apresentar tela"
            ]
            
            for pattern in shareNamePatterns {
                shareBtn := root.FindFirst(UIA.CreateCondition({ Name: pattern }))
                if shareBtn
                    break
            }
        }
        
        if !shareBtn
            return
        shareBtn.Invoke()
        Sleep 1000
        if li := WaitListItemMultiLang(root, windowListTexts)
            li.Invoke()
    }

    ; --- Wait for the action to complete and ensure Teams window is activated ---
    Sleep 2000  ; Give Teams time to process the sharing toggle
    
    ; Re-activate the Teams window while preserving its size
    if ActivateWindowWithRetry(hwndTeams, 3, 300) {
        PlayMicrophoneBeep()
        ShowCenteredOverlay(hwndTeams, "SHARING TOGGLED")
    } else {
        ; Fallback: show overlay on previous window if Teams activation fails
        PlayMicrophoneBeep()
        ShowCenteredOverlay(prev, "SHARING TOGGLED")
    }
}

; =============================================================================
; Meeting: Exit Meeting
; Hotkey: Win+Alt+Shift+2
; Original File: Microsoft Teams - meeting shortcuts.ahk
; =============================================================================
#!+2:: {
    if !ActivateTeamsMeetingWindow()
        return
    response := MsgBox("Tem certeza de que deseja sair da reunião?", "Sair da reunião?", "YesNo Icon!")
    if response = "Yes"
        Send "^+h"
}

; =============================================================================
; Activate Chat Window
; Hotkey: Win+Alt+Shift+E
; Original File: Microsoft Teams - meeting shortcuts.ahk
; =============================================================================
#!+E:: {
    if !ActivateTeamsChatWindow() {
        RunTeams()
    }
}

RunTeams() {
    ; Example for Microsoft Store Teams
    ; Run("shell:AppsFolder\MicrosoftTeams_8wekyb3d8bbwe!App")
    
    ; Example for desktop Teams
    Run("c:\Users\fie7ca\Documents\Atalhos\Microsoft Teams - Shortcut.lnk")
}

; =============================================================================
; Activate Meeting Window
; Hotkey: Win+Alt+Shift+3
; Original File: Microsoft Teams - meeting shortcuts.ahk
; =============================================================================
#!+3:: {
    if !ActivateTeamsMeetingWindow()
        ShowCenteredOverlay(WinGetID("A"), "NO ACTIVE MEETING WINDOW", 3000)
}

; =============================================================================
; Start New Conversation
; Hotkey: Win+Alt+Shift+R
; Original File: Microsoft Teams - New conversation.ahk
; =============================================================================
#!+r::
{
    contact := Trim(InputBox("Enter a Teams contact name:", "Jump to Chat").Value)
    if contact = ""
        return
    SetWinDelay 0
    SetKeyDelay 0, 0
    SetControlDelay 0
    teamsWindow := "Microsoft Teams"
    if !WinExist("ahk_exe ms-teams.exe") && !WinExist("ahk_exe Teams.exe") {
        Run "ms-teams:"
        WinWait(teamsWindow, , 15)
    }
    WinActivate(teamsWindow)
    WinWaitActive(teamsWindow, , 5)
    Send "^g"
    Sleep 100
    ; Save current clipboard content
    ClipboardOld := ClipboardAll()
    
    ; Ensure clipboard contains the correct contact name with retry logic
    Loop 5 {
        A_Clipboard := contact
        ; Wait for clipboard to contain data and verify it's the correct content
        if ClipWait(2) && (A_Clipboard = contact) {
            break
        }
        if A_Index = 5 {
            ; Restore clipboard and show error if we couldn't set it correctly
            A_Clipboard := ClipboardOld
            ShowCenteredOverlay(WinGetID("A"), "CLIPBOARD ERROR - TRY AGAIN", 3000)
            return
        }
        Sleep 100
    }
    
    Send "^v"
    Sleep 200  ; Give more time for the paste operation
    
    ; Restore original clipboard content
    A_Clipboard := ClipboardOld
    Sleep 600
    Send "{Enter}"
    Sleep 300
    Send "^r"
}
