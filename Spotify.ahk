#Requires AutoHotkey v2.0+
#SingleInstance Force
#UseHook  ; Ensure Volume hotkeys are captured before the OS processes them

; -----------------------------------------------------------------------------
; This script consolidates all Spotify related hotkeys and functions.
; -----------------------------------------------------------------------------

; --- Includes ----------------------------------------------------------------
#include UIA-v2\Lib\UIA.ahk
#include UIA-v2\Lib\UIA_Browser.ahk
#include %A_ScriptDir%\env.ahk
#include %A_ScriptDir%\SpotifyWASAPI.ahk

; --- Feature: use WASAPI for Ctrl+Volume (no window activation). Set false to use legacy activate+send.
global AL_USE_WASAPI := true

; --- Browser group for YouTube targeting (Chrome, Edge, Firefox, Brave) -------
GroupAdd "YouTubeBrowsers", "ahk_exe chrome.exe"
GroupAdd "YouTubeBrowsers", "ahk_exe msedge.exe"
GroupAdd "YouTubeBrowsers", "ahk_exe firefox.exe"
GroupAdd "YouTubeBrowsers", "ahk_exe brave.exe"

; --- Hotkeys & Functions -----------------------------------------------------

; =============================================================================
; Open or Activate Spotify
; Hotkey: Win+Alt+Shift+S
; Original File: Spotify - Open.ahk
; =============================================================================
#!+s:: OpenSpotify()

OpenSpotify() {
    ; 1) If Spotify is already running, just activate it. Exact process only; no title substring.
    if WinExist("ahk_exe Spotify.exe") {
        WinActivate("ahk_exe Spotify.exe")
        WinWaitActive("ahk_exe Spotify.exe", , 2)
        return
    }

    ; 2) Resolve launch command: dynamic path then Store fallback.
    link := GetSpotifyShortcutPath()
    if (link != "") {
        Run(link)
        WinWaitActive("ahk_exe Spotify.exe", , 5)
        return
    }
    ; Store / UWP fallback
    Run("explorer.exe shell:AppsFolder\SpotifyAB.SpotifyMusic_zpdnekdrzrea0!Spotify")
    WinWaitActive("ahk_exe Spotify.exe", , 5)
}

; Returns path to Spotify shortcut (Start Menu or Programs) or "" for Store-only. No hardcoded user paths.
GetSpotifyShortcutPath() {
    ; Prefer user Start Menu Programs (works for current user on any machine).
    path := A_AppData "\Microsoft\Windows\Start Menu\Programs\Spotify.lnk"
    if FileExist(path)
        return path
    ; Common alternate: All Users Start Menu (if installed for all users).
    try {
        path := A_ProgramsCommon "\Spotify.lnk"
        if FileExist(path)
            return path
    } catch {
        ;
    }
    ; Optional: resolve Store package from registry (HKCU ... AppModel\Repository\Packages).
    ; For now we rely on shell:AppsFolder fallback in OpenSpotify; registry traversal can be added here.
    return ""
}

*Volume_Down:: HandleVolumeDelta(-1)
*Volume_Up:: HandleVolumeDelta(1)

HandleVolumeDelta(deltaStep) {
    prevHwnd := WinGetID("A")
    if GetKeyState("Ctrl", "P") {
        ; Ctrl held: adjust Spotify volume (WASAPI = silent; else legacy activate+send)
        hwnd := GetSpotifyHwnd()
        if !(hwnd is Integer) || (hwnd <= 0) {
            ToolTip("Spotify not running")
            SetTimer(() => ToolTip(), -2000)
            return
        }
        if (AL_USE_WASAPI) {
            try {
                pid := WinGetPID("ahk_id " hwnd)
                if (pid is Integer) && (pid > 0) && AdjustProcessVolumeByPid(pid, deltaStep > 0 ? 5 : -5)
                    return
            } catch {
                ; Fall through to legacy
            }
        }
        wasMinimized := (WinGetMinMax(hwnd) == -1)
        WinActivate("ahk_id " hwnd)
        if WinWaitActive("ahk_id " hwnd, , 2) {
            Send(deltaStep > 0 ? "^{Up}" : "^{Down}")
        }
        if wasMinimized {
            SetTimer(VerifiedMinimize.Bind(hwnd), -3500)
        } else {
            SetTimer(RestoreFocus.Bind(prevHwnd), -800)
        }
    } else if GetKeyState("Alt", "P") {
        ; Alt held: adjust YouTube volume
        hwnd := GetYouTubeTabHwnd()
        if (hwnd is Integer) && (hwnd > 0) {
            wasMinimized := (WinGetMinMax(hwnd) == -1)
            try {
                WinActivate(hwnd)
                WinWaitActive(hwnd, , 2)
                Send(deltaStep > 0 ? "{Up}" : "{Down}")
            } catch {
                Send(deltaStep > 0 ? "{Volume_Up}" : "{Volume_Down}")
                return
            }
            if wasMinimized {
                SetTimer(VerifiedMinimize.Bind(hwnd), -3500)
            } else {
                SetTimer(RestoreFocus.Bind(prevHwnd), -800)
            }
        } else {
            Send(deltaStep > 0 ? "{Volume_Up}" : "{Volume_Down}")
        }
    } else {
        Send(deltaStep > 0 ? "{Volume_Up}" : "{Volume_Down}")
    }
}

; Restore focus to previous window only if it still exists. Deterministic; no Alt+Tab.
RestoreFocus(prevHwnd) {
    if (prevHwnd is Integer) && (prevHwnd > 0) && WinExist("ahk_id " prevHwnd)
        WinActivate("ahk_id " prevHwnd)
}

; Minimize only if window still exists. Avoids mutating wrong window after HWND reuse.
VerifiedMinimize(hwnd) {
    if (hwnd is Integer) && (hwnd > 0) && WinExist("ahk_id " hwnd)
        WinMinimize("ahk_id " hwnd)
}

; Returns Spotify window HWND (integer) or 0 if not found. Strict sentinel contract.
GetSpotifyHwnd() {
    hwnd := WinExist("ahk_exe Spotify.exe")
    return (hwnd) ? hwnd : 0
}

; Returns first browser window HWND whose current tab URL is a YouTube domain, or 0. Uses UIA + ahk_group.
GetYouTubeTabHwnd() {
    winList := WinGetList("ahk_group YouTubeBrowsers")
    for win in winList {
        try {
            uia := UIA_Browser("ahk_id " win)
            url := uia.GetCurrentURL()
            if IsYouTubeDomain(url)
                return win
        } catch {
            continue
        }
    }
    return 0
}

; True only if url contains a YouTube domain (www.youtube.com, m.youtube.com, youtube.com, etc.).
IsYouTubeDomain(url) {
    if (url = "" || Type(url) != "String")
        return false
    return InStr(url, "youtube.com") > 0
}
