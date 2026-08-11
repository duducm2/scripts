#Requires AutoHotkey v2.0+
#SingleInstance Force

; -----------------------------------------------------------------------------
; This script consolidates all Microsoft Teams related hotkeys and functions.
; -----------------------------------------------------------------------------

; --- Config and feature flags (Phase 1 / Phase 3 rollout) --------------------
; Rollout order: 1) HWND cache 2) activation/wait params 3) UIA cache (always on)
TEAMS_USE_HWND_CACHE := true
TEAMS_ACTIVATION_ATTEMPTS := 3
TEAMS_ACTIVATION_WAIT_MS := 300
WS_VISIBLE := 0x10000000

; --- Includes ----------------------------------------------------------------
#include vendor\UIA-v2\Lib\UIA.ahk
#include %A_ScriptDir%\Lib\TeamsContext.ahk
#include %A_ScriptDir%\Utils.ahk
try Hotkey("#!+Y", "Off")
try Hotkey("#!+X", "Off")

; --- Singleton HWND cache (Phase 1.3) ----------------------------------------
class TeamsHwndCache {
    static MeetingHwnd := 0
    static ChatHwnd := 0
    static IsValid(hwnd) {
        if (!hwnd || hwnd <= 0)
            return false
        try {
            if !WinExist("ahk_id " hwnd)
                return false
            style := WinGetStyle("ahk_id " hwnd)
            return (style & WS_VISIBLE) != 0
        } catch {
            return false
        }
    }
    static InvalidateMeeting() {
        TeamsHwndCache.MeetingHwnd := 0
    }
    static InvalidateChat() {
        TeamsHwndCache.ChatHwnd := 0
    }
}

; --- Helper Functions --------------------------------------------------------

; Returns meeting window HWND (integer) or 0. Uses cache when TEAMS_USE_HWND_CACHE.
ResolveTeamsMeetingHwnd() {
    if (TEAMS_USE_HWND_CACHE && TeamsHwndCache.IsValid(TeamsHwndCache.MeetingHwnd)) {
        if TeamsIsMeetingHwnd(TeamsHwndCache.MeetingHwnd)
            return TeamsHwndCache.MeetingHwnd
        TeamsHwndCache.InvalidateMeeting()
    }
    for proc in TEAMS_PROCESSES {
        for hwnd in WinGetList("ahk_exe " proc) {
            if TeamsIsMeetingHwnd(hwnd) {
                if (TEAMS_USE_HWND_CACHE)
                    TeamsHwndCache.MeetingHwnd := hwnd
                return hwnd
            }
        }
    }
    if (TEAMS_USE_HWND_CACHE)
        TeamsHwndCache.InvalidateMeeting()
    return 0
}

ActivateWindowWithRetry(hwnd, attempts := 3, waitMs := 300) {
    if (!hwnd || hwnd <= 0 || !WinExist("ahk_id " hwnd)) {
        try ShowCenteredOverlay(WinGetID("A"), "❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
        catch {
        }
        return false
    }
    originalState := ""
    try {
        originalState := WinGetMinMax(hwnd)
        if !(originalState = -1 || originalState = 0 || originalState = 1)
            originalState := ""
    } catch {
        originalState := ""
    }
    loop attempts {
        try {
            if (originalState = -1) {
                WinRestore(hwnd)
                Sleep 100
            }
            WinActivate(hwnd)
            if WinWaitActive("ahk_id " hwnd, , waitMs / 1000)
                return true
        } catch {
            try ShowCenteredOverlay(WinGetID("A"), "❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
        }
        try {
            if (originalState = -1) {
                DllCall("ShowWindow", "Ptr", hwnd, "Int", 9)
                Sleep 100
            }
            DllCall("SetForegroundWindow", "Ptr", hwnd)
            if WinWaitActive("ahk_id " hwnd, , waitMs / 1000)
                return true
        } catch {
            try ShowCenteredOverlay(WinGetID("A"), "❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
        }
        try {
            DllCall("BringWindowToTop", "Ptr", hwnd)
            Sleep 100
            WinActivate(hwnd)
            if WinWaitActive("ahk_id " hwnd, , waitMs / 1000)
                return true
        } catch {
            try ShowCenteredOverlay(WinGetID("A"), "❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
        }
        Sleep 200
    }
    return false
}

ActivateTeamsMeetingWindow() {
    hwnd := ResolveTeamsMeetingHwnd()
    if (hwnd <= 0) {
        TeamsHwndCache.InvalidateMeeting()
        ShowCenteredOverlay(WinGetID("A"), "NO MEETING WINDOW FOUND", 3000, BANNER_ACCENT_ERROR)
        return false
    }
    if ActivateWindowWithRetry(hwnd, TEAMS_ACTIVATION_ATTEMPTS, TEAMS_ACTIVATION_WAIT_MS)
        return true
    TeamsHwndCache.InvalidateMeeting()
    ShowCenteredOverlay(WinGetID("A"), "MEETING WINDOW FOUND BUT COULD NOT ACTIVATE", 3000, BANNER_ACCENT_ERROR)
    return false
}

ActivateTeamsChatWindow() {
    activeHwnd := WinExist("A")
    if (activeHwnd && TeamsIsChatHwnd(activeHwnd)) {
        WinActivate(activeHwnd)
        return true
    }
    for proc in TEAMS_PROCESSES {
        for hwnd in WinGetList("ahk_exe " proc) {
            if TeamsIsChatHwnd(hwnd) {
                if (WinExist("ahk_id " hwnd)) {
                    WinActivate(hwnd)
                    return true
                }
            }
        }
    }
    return false
}

; Returns a visible Teams HWND suitable for generic navigation (chat/search), or 0.
ResolveTeamsMainHwnd() {
    for proc in TEAMS_PROCESSES {
        for hwnd in WinGetList("ahk_exe " proc) {
            try {
                title := WinGetTitle(hwnd)
                if (!title)
                    continue
                if InStr(title, "Sharing control bar |")
                    continue
                if (RegExMatch(title, "i)\| Microsoft Teams"))
                    return hwnd
            } catch {
                continue
            }
        }
    }
    return 0
}

; Phase 2.3: single list-item finder (string or array of strings).
FindListItemByNames(root, nameOrNames) {
    names := Type(nameOrNames) = "Array" ? nameOrNames : [nameOrNames]
    items := root.FindAll(UIA.CreateCondition({ ControlType: "ListItem" }))
    for item in items {
        for text in names {
            if InStr(item.Name, text)
                return item
        }
    }
    return false
}

; Single wait helper; poll interval 150 ms.
WaitListItemByNames(root, nameOrNames, timeout := 3000) {
    start := A_TickCount
    while (A_TickCount - start < timeout) {
        item := FindListItemByNames(root, nameOrNames)
        if item
            return item
        Sleep 150
    }
    return false
}

; --- Standard overlay (uses Utils) -------------------------------------------
ShowCenteredOverlay(hwndTarget, text, duration := 1500, bgColor := BANNER_ACCENT_INTERMEDIATE) {
    ; hwndTarget kept for API compatibility; placement uses foreground monitor (centerOnHwnd 0) after WinActivate patterns.
    StandardLoadingBar_Show(text, bgColor, { passive: true, centerOnHwnd: 0, textWidth: 500, fontSize: 17,
        passiveBgColor: bgColor })
    StandardLoadingBar_Hide(duration)
}

; --- Hotkeys & Functions -----------------------------------------------------

; --- Audio feedback helper ---
PlayMicrophoneBeep() {
    ; Play a single short beep to indicate microphone action (if enabled)
    ScriptSoundBeep(800, 150)
}

; Phase 2.1/2.2: one cache request for toggle buttons; generic state resolver.
TEAMS_TOGGLE_CACHE := UIA.CreateCacheRequest(["Name", "AutomationId"], ["Toggle"])

; Returns state string from mapping, or "unknown". Uses cached UIA (Phase 2.1).
GetTeamsToggleState(hwndTeams, automationId, namePatterns, stateFromToggle, stateFromName, maxRetries := 3) {
    loop maxRetries {
        try {
            root := UIA.ElementFromHandleBuildCache(TEAMS_TOGGLE_CACHE, hwndTeams)
            if !root {
                if A_Index < maxRetries
                    Sleep 150
                continue
            }
            btn := root.FindFirstBuildCache(TEAMS_TOGGLE_CACHE, UIA.CreateCondition({ AutomationId: automationId }))
            if !btn {
                for pattern in namePatterns {
                    btn := root.FindFirstBuildCache(TEAMS_TOGGLE_CACHE, UIA.CreateCondition({ Name: pattern }))
                    if btn
                        break
                }
            }
            if btn {
                try {
                    state := btn.CachedToggleState
                    if (state is "Integer" && stateFromToggle.Has(state))
                        return stateFromToggle[state]
                } catch {
                }
                try {
                    name := btn.CachedName
                    if name
                        for needle, result in stateFromName
                            if InStr(name, needle)
                                return result
                } catch {
                }
            }
        } catch {
            if A_Index < maxRetries
                Sleep 200
        }
    }
    return "unknown"
}

; stateFromToggle: map ToggleState (0/1/2) to label; stateFromName: map Name substring to label (or use Has + InStr).
GetMicrophoneState(hwndTeams, maxRetries := 3) {
    stateFromToggle := Map(UIA.ToggleState.On, "muted", UIA.ToggleState.Off, "unmuted")
    stateFromName := Map(
        "Desativar mudo", "muted", "Desligar microfone", "muted", "Ativar mudo", "unmuted", "Ligar microfone",
        "unmuted",
        "Unmute", "muted", "Turn on microphone", "muted", "Mute", "unmuted", "Turn off microphone", "unmuted")
    return GetTeamsToggleState(hwndTeams, "microphone-button", [
        "Microphone", "Mic", "Mute", "Unmute", "Microfone", "Mudo",
        "Turn on microphone", "Turn off microphone", "Ligar microfone", "Desligar microfone"
    ], stateFromToggle, stateFromName, maxRetries)
}

GetCameraState(hwndTeams, maxRetries := 3) {
    stateFromToggle := Map(UIA.ToggleState.On, "on", UIA.ToggleState.Off, "off")
    stateFromName := Map(
        "Desativar câmera", "on", "Parar vídeo", "on", "Ativar câmera", "off", "Iniciar vídeo", "off",
        "Turn off camera", "on", "Turn camera off", "on", "Stop video", "on",
        "Turn on camera", "off", "Turn camera on", "off", "Start video", "off")
    return GetTeamsToggleState(hwndTeams, "video-button", [
        "Camera", "Video", "Turn on camera", "Turn off camera", "Turn camera on", "Turn camera off",
        "Câmera", "Vídeo", "Ativar câmera", "Desativar câmera", "Start video", "Stop video", "Iniciar vídeo",
        "Parar vídeo"
    ], stateFromToggle, stateFromName, maxRetries)
}

; =============================================================================
; Meeting: Toggle Mute
; Hotkey: Win+Alt+Shift+5
; Original File: Microsoft Teams - meeting shortcuts.ahk
; =============================================================================
#!+5:: {
    ; Check if Teams is closed and prompt to open if needed
    if (!CheckAndOpenOutlookTeams(false, true)) {
        return  ; User cancelled opening Teams
    }

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
    loop 3 {
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
            ShowCenteredOverlay(prev, "🔇 MIC MUTED")
        else
            ShowCenteredOverlay(prev, "🔊 MIC UNMUTED")
        return
    }

    ; On failure, show an error banner and do not beep
    WinActivate(prev)
    ShowCenteredOverlay(prev, "❓ MICROPHONE STATE UNKNOWN", 3000)
}

; =============================================================================
; Meeting: Toggle Camera
; Hotkey: Win+Alt+Shift+4
; Original File: Microsoft Teams - meeting shortcuts.ahk
; =============================================================================
#!+4:: {
    ; Check if Teams is closed and prompt to open if needed
    if (!CheckAndOpenOutlookTeams(false, true)) {
        return  ; User cancelled opening Teams
    }

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
    loop 3 {
        Sleep 250
        finalState := GetCameraState(hwndTeams)
        if (finalState != "unknown")
            break
    }

    WinActivate(prev)

    if (finalState = "on" || finalState = "off") {
        PlayMicrophoneBeep()
        if finalState = "on"
            ShowCenteredOverlay(prev, "📷 CAMERA ON")
        else
            ShowCenteredOverlay(prev, "📷 CAMERA OFF")
        return
    }

    ShowCenteredOverlay(prev, "❓ CAMERA STATE UNKNOWN", 3000)
}

; =============================================================================
; Meeting: Toggle Screen Share  (Win Alt Shift T)
; =============================================================================
#!+t:: {
    ; Check if Teams is closed and prompt to open if needed
    if (!CheckAndOpenOutlookTeams(false, true)) {
        return  ; User cancelled opening Teams
    }

    prev := WinGetID("A")                 ; remember the window you were in
    if !ActivateTeamsMeetingWindow()
        return

    hwndTeams := WinGetID("A")            ; Teams meeting window
    root := UIA.ElementFromHandle(hwndTeams)
    if !root
        return

    ; --- perform the normal sharing workflow ---
    windowListTexts := ["Opens list of", "Abre a lista de"]
    listItem := FindListItemByNames(root, windowListTexts)
    if listItem {
        listItem.Invoke()
    } else {
        shareBtn := root.FindFirst(UIA.CreateCondition({ AutomationId: "share-button" }))
        if !shareBtn {
            for pattern in ["Share", "Share content", "Share screen", "Start sharing", "Compartilhar",
                "Compartilhar conteúdo", "Compartilhar tela", "Iniciar compartilhamento", "Present", "Present screen",
                "Apresentar", "Apresentar tela"] {
                shareBtn := root.FindFirst(UIA.CreateCondition({ Name: pattern }))
                if shareBtn
                    break
            }
        }
        if !shareBtn
            return
        shareBtn.Invoke()
        Sleep 1000
        if li := WaitListItemByNames(root, windowListTexts)
            li.Invoke()
    }

    ; --- Wait for the action to complete and ensure Teams window is activated ---
    Sleep 2000  ; Give Teams time to process the sharing toggle

    ; Re-activate the Teams window while preserving its size
    if ActivateWindowWithRetry(hwndTeams, 3, 300) {
        PlayMicrophoneBeep()
        ShowCenteredOverlay(hwndTeams, "🖥 SHARING TOGGLED")
    } else {
        ; Fallback: show overlay on previous window if Teams activation fails
        PlayMicrophoneBeep()
        ShowCenteredOverlay(prev, "🖥 SHARING TOGGLED")
    }
}

; =============================================================================
; Meeting: Exit Meeting
; Hotkey: Win+Alt+Shift+2
; Original File: Microsoft Teams - meeting shortcuts.ahk
; =============================================================================
#!+2:: {
    ; Check if Teams is closed and prompt to open if needed
    if (!CheckAndOpenOutlookTeams(false, true)) {
        return  ; User cancelled opening Teams
    }

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
    ; Check if Teams is closed and prompt to open if needed
    if (!CheckAndOpenOutlookTeams(false, true)) {
        return  ; User cancelled opening Teams
    }

    if !ActivateTeamsChatWindow() {
        RunTeams()
    }
}

; Resolve Teams launch path: env/registry then fallback to ms-teams.exe (Phase 1.4).
ResolveTeamsPath() {
    global IS_WORK_ENVIRONMENT
    ; 1) Environment-based: WindowsApps alias or Start Menu
    localAppData := EnvGet("LOCALAPPDATA")
    if (localAppData) {
        teamsExe := localAppData "\Microsoft\WindowsApps\ms-teams.exe"
        if (FileExist(teamsExe))
            return teamsExe
    }
    appData := EnvGet("APPDATA")
    if (appData) {
        shortcut := appData "\Microsoft\Windows\Start Menu\Programs\Microsoft Teams.lnk"
        if (FileExist(shortcut))
            return shortcut
    }
    ; 2) Registry: current user packages for MSTeams
    try {
        loop reg "HKCU\SOFTWARE\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\Repository\Packages",
            "K" {
            if InStr(A_LoopRegName, "MSTeams_") && InStr(A_LoopRegName, "8wekyb3d8bbwe") {
                basePath := "C:\Program Files\WindowsApps\" A_LoopRegName "\ms-teams.exe"
                if (FileExist(basePath))
                    return basePath
            }
        }
    } catch {
    }
    ; 3) Fallback: protocol or executable alias
    return "ms-teams:"
}

RunTeams() {
    path := ResolveTeamsPath()
    try {
        if (path = "ms-teams:")
            Run "ms-teams:"
        else
            Run path
    } catch {
        Run "ms-teams.exe"
    }
}

; =============================================================================
; Activate Meeting Window
; Hotkey: Win+Alt+Shift+3
; Original File: Microsoft Teams - meeting shortcuts.ahk
; =============================================================================
#!+3:: {
    ; Check if Teams is closed and prompt to open if needed
    if (!CheckAndOpenOutlookTeams(false, true)) {
        return  ; User cancelled opening Teams
    }

    if !ActivateTeamsMeetingWindow()
        ShowCenteredOverlay(WinGetID("A"), "⚠ NO ACTIVE MEETING WINDOW", 3000, BANNER_ACCENT_ERROR)
}

; =============================================================================
; Start New Conversation
; Hotkey: Win+Alt+Shift+R
; 1× = Teams jump to chat; 2× within 400ms = WhatsApp jump to chat
; Original File: Microsoft Teams - New conversation.ahk
; =============================================================================
global Teams_R_DoubleTapArmed := false
global Teams_R_LastPressTick := 0
global Teams_R_DoubleTapThresholdMs := 400  ; Matches ZMK tapping-term-ms for tap-dance
global Teams_R_DoubleTapTimer := 0
global Teams_R_DoubleTapOverlayGui := ""

Teams_R_DoubleTap_ShowOverlay() {
    global Teams_R_DoubleTapOverlayGui
    if (Teams_R_DoubleTapOverlayGui) {
        Teams_R_DoubleTapOverlayGui.Destroy()
        Teams_R_DoubleTapOverlayGui := ""
    }
    Teams_R_DoubleTapOverlayGui := Gui("+ToolWindow +AlwaysOnTop -Caption +E0x20 -DPIScale", "Teams_R_DoubleTapOverlay"
    )
    Teams_R_DoubleTapOverlayGui.BackColor := "2980B9"
    WinSetTransColor("2980B9", Teams_R_DoubleTapOverlayGui.Hwnd)
    Teams_R_DoubleTapOverlayGui.SetFont("s11 w700 cWhite", "Segoe UI")
    ovlW := 300, ovlH := 36
    Teams_R_DoubleTapOverlayGui.Add("Text", "x12 y8 w276 h20 0x200 Background2980B9 Center",
        "Tap again → WhatsApp")
    mon := 0
    try mon := MonitorGetPrimary()
    MonitorGetWorkArea(mon, &mlx, &mty, &mrx, &mby)
    ovlX := mlx + ((mrx - mlx) - ovlW) // 2
    ovlY := mty + 48
    Teams_R_DoubleTapOverlayGui.Show("NoActivate x" ovlX " y" ovlY " w" ovlW " h" ovlH)
}

Teams_R_DoubleTap_HideOverlay() {
    global Teams_R_DoubleTapOverlayGui
    if (Teams_R_DoubleTapOverlayGui) {
        Teams_R_DoubleTapOverlayGui.Destroy()
        Teams_R_DoubleTapOverlayGui := ""
    }
}

class Teams_R_DoubleTapTimerObj {
    static OnSingleTapTimeout() {
        global Teams_R_DoubleTapArmed, Teams_R_DoubleTapTimer
        if (!Teams_R_DoubleTapArmed)
            return
        Teams_R_DoubleTapArmed := false
        Teams_R_DoubleTapTimer := 0
        Teams_R_DoubleTap_HideOverlay()
        Teams_StartNewConversationJump()
    }
}

Teams_StartNewConversationJump() {
    if (!CheckAndOpenOutlookTeams(false, true))
        return
    contact := Trim(InputBox("Enter a Teams contact name:", "Jump to Chat").Value)
    if contact = ""
        return
    clipSaved := ClipboardAll()
    try {
        TeamsJumpToChat(contact)
    } finally {
        A_Clipboard := clipSaved
        if (ClipWait(1)) {
        }
    }
}

WhatsApp_StartNewConversationJump() {
    contact := Trim(InputBox("Enter a WhatsApp contact name:", "Jump to Chat").Value)
    if contact = ""
        return
    clipSaved := ClipboardAll()
    try {
        WhatsAppJumpToChat(contact)
    } finally {
        A_Clipboard := clipSaved
        if (ClipWait(1)) {
        }
    }
}

#!+r:: {
    global Teams_R_DoubleTapArmed, Teams_R_LastPressTick, Teams_R_DoubleTapThresholdMs,
        Teams_R_DoubleTapTimer

    now := A_TickCount
    elapsed := (Teams_R_LastPressTick > 0) ? (now - Teams_R_LastPressTick) : 9999

    if (Teams_R_DoubleTapArmed && elapsed >= 0 && elapsed < Teams_R_DoubleTapThresholdMs) {
        Teams_R_DoubleTapArmed := false
        Teams_R_LastPressTick := 0
        if (Teams_R_DoubleTapTimer) {
            SetTimer(Teams_R_DoubleTapTimer, 0)
            Teams_R_DoubleTapTimer := 0
        }
        Teams_R_DoubleTap_HideOverlay()
        WhatsApp_StartNewConversationJump()
        return
    }

    Teams_R_LastPressTick := now
    Teams_R_DoubleTapArmed := true
    Teams_R_DoubleTap_ShowOverlay()
    Teams_R_DoubleTapTimer := ObjBindMethod(Teams_R_DoubleTapTimerObj, "OnSingleTapTimeout")
    SetTimer(Teams_R_DoubleTapTimer, -Teams_R_DoubleTapThresholdMs)
}
