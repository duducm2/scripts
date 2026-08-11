; Teams meeting/chat context detection (UIA-first).
; Requires vendor\UIA-v2\Lib\UIA.ahk loaded before this file.

TEAMS_PROCESSES := ["ms-teams.exe", "Teams.exe", "MSTeams.exe"]
TEAMS_CONTEXT_CACHE_MS := 400
TEAMS_DEBUG_CONTEXT := false

TEAMS_MEETING_TOOLBAR_IDS := ["microphone-button", "video-button", "reaction-menu-button", "prejoin-join-button",
    "prejoin-audiosettings-button"]

class TeamsContextCache {
    static Hwnd := 0
    static Mode := ""
    static Tick := 0

    static Invalidate() {
        TeamsContextCache.Hwnd := 0
        TeamsContextCache.Mode := ""
        TeamsContextCache.Tick := 0
    }

    static Get(hwnd) {
        if (!hwnd || hwnd <= 0)
            return ""
        now := A_TickCount
        if (hwnd = TeamsContextCache.Hwnd && (now - TeamsContextCache.Tick) < TEAMS_CONTEXT_CACHE_MS)
            return TeamsContextCache.Mode
        mode := TeamsResolveContextUncached(hwnd)
        TeamsContextCache.Hwnd := hwnd
        TeamsContextCache.Mode := mode
        TeamsContextCache.Tick := now
        return mode
    }
}

IsTeamsProcessHwnd(hwnd) {
    if (!hwnd || hwnd <= 0)
        return false
    try {
        exe := WinGetProcessName("ahk_id " hwnd)
        for proc in TEAMS_PROCESSES {
            if (StrLower(exe) = StrLower(proc))
                return true
        }
    } catch {
    }
    return false
}

IsTeamsSharingBarTitle(title) {
    return InStr(title, "Sharing control bar |")
}

IsTeamsChatTitlePrefix(title) {
    return InStr(title, "Chat |") || InStr(title, "Bate-papo |")
}

; Title-only meeting hints (no broad catch-all regex).
IsTeamsMeetingTitle(title) {
    if !title
        return false
    if IsTeamsSharingBarTitle(title)
        return false
    if IsTeamsChatTitlePrefix(title)
        return false
    if InStr(title, "Microsoft Teams meeting") || InStr(title, "Reunião do Microsoft Teams")
        return true
    if InStr(title, "Modo de exibição compacto da reunião")
        return true
    ; Pre-join / lobby (e.g. "Meeting join | … | Microsoft Teams")
    if InStr(title, "Meeting join |") || InStr(title, "Ingressar na reunião |") || InStr(title,
        "Entrar na reunião |")
        return true
    return false
}

; Title-only chat hints (channel tabs, detached chat, main-window conversations).
IsTeamsChatTitle(title) {
    if !title || IsTeamsSharingBarTitle(title)
        return false
    if IsTeamsMeetingTitle(title)
        return false
    if IsTeamsChatTitlePrefix(title) && RegExMatch(title, "i)\| Microsoft Teams$")
        return true
    if RegExMatch(title, "i)\| Microsoft Teams$")
        return true
    return title = "Microsoft Teams"
}

TeamsHasMeetingToolbar(hwnd) {
    if (!hwnd || hwnd <= 0)
        return false
    try {
        root := UIA.ElementFromHandle(hwnd)
        if !root
            return false
        for id in TEAMS_MEETING_TOOLBAR_IDS {
            if root.FindFirst({ AutomationId: id })
                return true
        }
    } catch {
    }
    return false
}

IsTeamsWebViewHwnd(hwnd) {
    try {
        return WinGetClass("ahk_id " hwnd) = "TeamsWebView"
    } catch {
        return false
    }
}

TeamsTitleLooksLikeChat(title) {
    if !title || IsTeamsSharingBarTitle(title)
        return false
    if IsTeamsMeetingTitle(title)
        return false
    if RegExMatch(title, "i)\| Microsoft Teams")
        return true
    return title = "Microsoft Teams"
}

TeamsResolveContextUncached(hwnd) {
    if !IsTeamsProcessHwnd(hwnd)
        return ""
    try {
        title := WinGetTitle("ahk_id " hwnd)
    } catch {
        return ""
    }
    if IsTeamsSharingBarTitle(title)
        return ""
    ; Title first: pre-join is TeamsWebView and would otherwise fall through as chat
    ; when the deep UIA toolbar probe misses/times out.
    if IsTeamsMeetingTitle(title)
        return "meeting"
    if TeamsHasMeetingToolbar(hwnd)
        return "meeting"
    if IsTeamsWebViewHwnd(hwnd) || TeamsTitleLooksLikeChat(title)
        return "chat"
    return ""
}

TeamsResolveContext(hwnd) {
    return TeamsContextCache.Get(hwnd)
}

IsTeamsMeetingActive() {
    return TeamsResolveContext(WinExist("A")) = "meeting"
}

IsTeamsChatActive() {
    return TeamsResolveContext(WinExist("A")) = "chat"
}

; HWND-level checks for window enumeration (Microsoft Teams.ahk activation).
TeamsIsMeetingHwnd(hwnd) {
    if !IsTeamsProcessHwnd(hwnd)
        return false
    mode := TeamsResolveContextUncached(hwnd)
    if (mode = "meeting")
        return true
    try {
        title := WinGetTitle("ahk_id " hwnd)
    } catch {
        return false
    }
    return IsTeamsMeetingTitle(title)
}

TeamsIsChatHwnd(hwnd) {
    if !IsTeamsProcessHwnd(hwnd)
        return false
    return TeamsResolveContextUncached(hwnd) = "chat"
}

TeamsDebugShowContext(hwnd := 0) {
    if !TEAMS_DEBUG_CONTEXT
        return
    hwnd := hwnd || WinExist("A")
    mode := TeamsResolveContext(hwnd)
    try {
        title := WinGetTitle("ahk_id " hwnd)
    } catch {
        title := ""
    }
    try {
        className := WinGetClass("ahk_id " hwnd)
    } catch {
        className := ""
    }
    MsgBox "Teams context: " mode "`nTitle: " title "`nClass: " className, "TeamsContext debug", "Iconi T1"
}
