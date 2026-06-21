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

; Minimum width/height for the main Spotify UI (skip hidden Electron helper windows).
global SPOTIFY_MIN_WINDOW_SIZE := 200
global SPOTIFY_OPEN_WAIT_SEC := 12
global SPOTIFY_OPEN_RETRIES := 3
global SPOTIFY_VERIFY_SETTLE_MS := 200

OpenSpotify() {
    loop SPOTIFY_OPEN_RETRIES {
        if SpotifyOpenAttempt(A_Index)
            return
        if (A_Index < SPOTIFY_OPEN_RETRIES)
            Sleep 400
    }
    SpotifyShowOpenFailure()
}

; One open cycle: launch/restore as needed, then quality-gate on a usable foreground window.
SpotifyOpenAttempt(attempt := 1) {
    hwnd := GetSpotifyMainHwnd()
    if hwnd > 0 && SpotifyActivateAndVerify(hwnd)
        return true

    SpotifyLaunchForAttempt(attempt)
    return SpotifyWaitActivateAndVerify(SPOTIFY_OPEN_WAIT_SEC)
}

; Quality gate: main Spotify window exists, is usable, and is the active foreground window.
SpotifyIsOpenedAndActive(hwnd := 0) {
    if !hwnd
        hwnd := GetSpotifyMainHwnd()
    if !(hwnd is Integer) || hwnd <= 0 || !WinExist("ahk_id " hwnd)
        return false
    if !SpotifyIsUsableMainWindow(hwnd)
        return false
    return WinActive("ahk_id " hwnd)
}

SpotifyActivateAndVerify(hwnd) {
    if !(hwnd is Integer) || hwnd <= 0
        return false
    if !SpotifyActivateWindow(hwnd)
        return false
    Sleep SPOTIFY_VERIFY_SETTLE_MS
    current := GetSpotifyMainHwnd()
    if current
        hwnd := current
    return SpotifyIsOpenedAndActive(hwnd)
}

SpotifyWaitActivateAndVerify(timeoutSec := 12) {
    hwnd := SpotifyWaitForMainHwnd(timeoutSec)
    if hwnd > 0 && SpotifyActivateAndVerify(hwnd)
        return true
    ; Window may appear slightly after ProcessWait; one short re-check before failing the attempt.
    Sleep 400
    hwnd := GetSpotifyMainHwnd()
    return hwnd > 0 && SpotifyActivateAndVerify(hwnd)
}

; Escalating launch methods per retry (shortcut/exe path may move between installs).
SpotifyLaunchForAttempt(attempt := 1) {
    if (attempt = 1) {
        if SpotifyProcessExists()
            SpotifyLaunchRestore()
        else
            SpotifyLaunchFresh()
        return
    }
    if (attempt = 2) {
        Run("spotify:")
        if !SpotifyProcessExists()
            SpotifyLaunchFresh()
        return
    }
    exe := SpotifyResolveExePath()
    if exe != ""
        Run('"' exe '"')
    else if (link := GetSpotifyShortcutPath()) != ""
        Run(link)
    else
        Run("shell:AppsFolder\SpotifyAB.SpotifyMusic_zpdnekdrzrea0!Spotify")
    try
        ProcessWait("Spotify.exe", Min(SPOTIFY_OPEN_WAIT_SEC, 8))
    catch {
        ;
    }
}

SpotifyProcessExists() {
    return ProcessExist("Spotify.exe") > 0
}

SpotifyLaunchFresh() {
    link := GetSpotifyShortcutPath()
    if (link != "")
        Run(link)
    else if (exe := SpotifyResolveExePath()) != ""
        Run('"' exe '"')
    else
        Run("shell:AppsFolder\SpotifyAB.SpotifyMusic_zpdnekdrzrea0!Spotify")
    try
        ProcessWait("Spotify.exe", SPOTIFY_OPEN_WAIT_SEC)
    catch {
        ;
    }
}

; Process running but no main window (tray-only / crashed UI): restore via shortcut or spotify: URI.
SpotifyLaunchRestore() {
    link := GetSpotifyShortcutPath()
    if (link != "")
        Run(link)
    else if (exe := SpotifyResolveExePath()) != ""
        Run('"' exe '"')
    else
        Run("spotify:")
}

; Resolve Spotify.exe from shortcut target or common install locations (handles moved installs).
SpotifyResolveExePath() {
    link := GetSpotifyShortcutPath()
    if (link != "") {
        try {
            FileGetShortcut link, &target
            if (target != "") {
                if FileExist(target)
                    return target
                fixed := StrReplace(target, "Program Files (x86)", "Program Files")
                if FileExist(fixed)
                    return fixed
            }
        } catch {
            ;
        }
    }
    localAppData := EnvGet("LocalAppData")
    for candidate in [A_AppData "\Spotify\Spotify.exe", localAppData "\Microsoft\WindowsApps\Spotify.exe"] {
        if FileExist(candidate)
            return candidate
    }
    return ""
}

SpotifyWaitForMainHwnd(timeoutSec := 12) {
    deadline := A_TickCount + (timeoutSec * 1000)
    loop {
        hwnd := GetSpotifyMainHwnd()
        if hwnd > 0
            return hwnd
        if (A_TickCount >= deadline)
            break
        Sleep 200
    }
    return 0
}

SpotifyShowOpenFailure() {
    ToolTip("Could not open Spotify")
    SetTimer(() => ToolTip(), -2000)
}

SpotifyActivateWindow(hwnd, attempts := 3, waitMs := 300) {
    if !(hwnd is Integer) || hwnd <= 0 || !WinExist("ahk_id " hwnd)
        return false
    try {
        pid := WinGetPID("ahk_id " hwnd)
        if (pid is Integer) && pid > 0
            DllCall("AllowSetForegroundWindow", "UInt", pid)
    } catch {
        ;
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
            if WinActive("ahk_id " hwnd) || WinWaitActive("ahk_id " hwnd, , waitMs / 1000)
                return true
        } catch {
        }
        try {
            if (originalState = -1) {
                DllCall("ShowWindow", "Ptr", hwnd, "Int", 9)
                Sleep 100
            }
            DllCall("SetForegroundWindow", "Ptr", hwnd)
            if WinActive("ahk_id " hwnd) || WinWaitActive("ahk_id " hwnd, , waitMs / 1000)
                return true
        } catch {
        }
        try {
            DllCall("BringWindowToTop", "Ptr", hwnd)
            Sleep 100
            WinActivate(hwnd)
            if WinActive("ahk_id " hwnd) || WinWaitActive("ahk_id " hwnd, , waitMs / 1000)
                return true
        } catch {
        }
        Sleep 200
    }
    return false
}

SpotifyIsUsableMainWindow(hwnd) {
    if !(hwnd is Integer) || hwnd <= 0
        return false
    if !DllCall("IsWindowVisible", "Ptr", hwnd)
        return false
    try {
        WinGetPos(, , &w, &h, "ahk_id " hwnd)
        if (w < SPOTIFY_MIN_WINDOW_SIZE || h < SPOTIFY_MIN_WINDOW_SIZE)
            return false
    } catch {
        return false
    }
    return true
}

; Main Spotify UI HWND (visible, sized); prefers title containing "Spotify". Returns 0 if none.
GetSpotifyMainHwnd() {
    bestWithTitle := 0
    bestAny := 0
    for hwnd in WinGetList("ahk_exe Spotify.exe") {
        if !SpotifyIsUsableMainWindow(hwnd)
            continue
        try title := WinGetTitle(hwnd)
        catch
            title := ""
        if InStr(title, "Spotify") {
            if !bestWithTitle
                bestWithTitle := hwnd
        } else if !bestAny
            bestAny := hwnd
    }
    return bestWithTitle ? bestWithTitle : bestAny
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

; Returns Spotify main window HWND (integer) or 0 if not found. Strict sentinel contract.
GetSpotifyHwnd() {
    return GetSpotifyMainHwnd()
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
