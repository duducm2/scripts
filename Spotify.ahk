#Requires AutoHotkey v2.0+
#SingleInstance Force
#UseHook  ; Ensure Volume hotkeys are captured before the OS processes them

; -----------------------------------------------------------------------------
; This script consolidates all Spotify related hotkeys and functions.
; -----------------------------------------------------------------------------

; --- Includes ----------------------------------------------------------------
#include vendor\UIA-v2\Lib\UIA.ahk
#include vendor\UIA-v2\Lib\UIA_Browser.ahk
#include %A_ScriptDir%\env.ahk
#include %A_ScriptDir%\lib\SpotifyWASAPI.ahk
; Thin include: Show/Update/Hide only. Do NOT #include full Utils.ahk (would inject
; unrelated hotkeys / a second global Escape host — see Act.ahk comment).
#include %A_ScriptDir%\Utils\standard_loading_bar.ahk

; standard_loading_bar.ahk references this only on its keys-overlay paths, which Spotify
; never uses. Stub keeps load-time resolution valid without importing the global Escape
; hook from Utils\print_screen_escape.ahk (a second Escape host breaks modals - see Act.ahk 121).
Utils_EnsureGlobalEscapeHotkey() {
}

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
global SPOTIFY_OPEN_TOTAL_BUDGET_MS := 24000
global SPOTIFY_FIRST_ATTEMPT_WAIT_SEC := 10
global SPOTIFY_RETRY_WAIT_SEC := 6
global SPOTIFY_OPEN_GUARD_MAX_MS := 45000
global SPOTIFY_BANNER_DELAY_MS := 400
global SPOTIFY_RESPONSIVE_TIMEOUT_MS := 300
; Optional open diagnosis → .cursor\spotify_open_quality.log (leave false in normal use).
global SPOTIFY_OPEN_DEBUG := false

global g_SpotifyOpenInProgress := false
global g_SpotifyOpenStartTick := 0
global g_SpotifyBannerVisible := false
global g_SpotifyBannerArmText := ""
global g_SpotifyBannerTimerArmed := false
global g_SpotifyOpenLastPhase := ""
global g_SpotifyOpenLastAttempt := 0
global g_SpotifyOpenDeadline := 0

OpenSpotify() {
    global g_SpotifyOpenInProgress, g_SpotifyOpenStartTick, g_SpotifyOpenLastPhase
    global g_SpotifyOpenLastAttempt, g_SpotifyOpenDeadline, g_SpotifyBannerVisible

    ; Self-expiring re-entrancy guard: a crashed run cannot permanently disable #!+s.
    if (g_SpotifyOpenInProgress) {
        if (g_SpotifyOpenStartTick > 0 && (A_TickCount - g_SpotifyOpenStartTick) < SPOTIFY_OPEN_GUARD_MAX_MS) {
            SpotifyBanner_Result("⏳ Already opening Spotify...", BANNER_ACCENT_INTERMEDIATE, 900)
            return
        }
        g_SpotifyOpenInProgress := false
    }

    g_SpotifyOpenInProgress := true
    g_SpotifyOpenStartTick := A_TickCount
    g_SpotifyOpenLastPhase := "start"
    g_SpotifyOpenLastAttempt := 0
    g_SpotifyOpenDeadline := A_TickCount + SPOTIFY_OPEN_TOTAL_BUDGET_MS
    succeeded := false
    try {
        SpotifyBanner_Arm("⏳ Opening Spotify...")
        SpotifyOpenDebugLog("begin")
        loop SPOTIFY_OPEN_RETRIES {
            g_SpotifyOpenLastAttempt := A_Index
            if (A_TickCount >= g_SpotifyOpenDeadline) {
                g_SpotifyOpenLastPhase := "budget_exhausted"
                break
            }
            if SpotifyOpenAttempt(A_Index, g_SpotifyOpenDeadline) {
                succeeded := true
                g_SpotifyOpenLastPhase := "ok"
                SpotifyOpenDebugLog("success", GetSpotifyMainHwnd())
                break
            }
            if (A_Index < SPOTIFY_OPEN_RETRIES && A_TickCount < g_SpotifyOpenDeadline)
                SpotifySleepBudget(400, g_SpotifyOpenDeadline)
        }
        if (succeeded) {
            ; Brief success only when the deferred loading bar actually appeared.
            if (g_SpotifyBannerVisible)
                SpotifyBanner_Result("✅ Spotify is open", BANNER_ACCENT_SUCCESS, 700)
            else
                SpotifyBanner_Cancel()
        } else {
            SpotifyShowOpenFailure()
        }
    } finally {
        ; Never Hide after Result — it already replaced the loading bar with a delayed
        ; Information Only banner. Only clear a still-visible loading bar (e.g. exception).
        try SetTimer(SpotifyBanner_ShowDeferred, 0)
        catch {
        }
        global g_SpotifyBannerTimerArmed, g_SpotifyBannerArmText
        g_SpotifyBannerTimerArmed := false
        g_SpotifyBannerArmText := ""
        if (g_SpotifyBannerVisible) {
            try StandardLoadingBar_Hide(0)
            catch {
            }
            g_SpotifyBannerVisible := false
        }
        g_SpotifyOpenInProgress := false
        g_SpotifyOpenStartTick := 0
    }
}

; One open cycle: launch/restore as needed, then quality-gate on a usable foreground window.
SpotifyOpenAttempt(attempt := 1, deadline := 0) {
    global g_SpotifyOpenLastPhase
    if (!deadline)
        deadline := A_TickCount + SPOTIFY_OPEN_TOTAL_BUDGET_MS

    hwnd := GetSpotifyMainHwnd()
    if hwnd > 0 {
        g_SpotifyOpenLastPhase := "activate_existing"
        SpotifyBanner_Set("🔄 Spotify: verifying...")
        SpotifyOpenDebugLog("activate_existing", hwnd, attempt)
        if SpotifyActivateAndVerify(hwnd, deadline)
            return true
    }

    waitSec := (attempt = 1) ? SPOTIFY_FIRST_ATTEMPT_WAIT_SEC : SPOTIFY_RETRY_WAIT_SEC
    g_SpotifyOpenLastPhase := "launch"
    SpotifyBanner_Set("⏳ Spotify: launching (attempt " attempt "/" SPOTIFY_OPEN_RETRIES ")...")
    SpotifyOpenDebugLog("launch", 0, attempt)
    SpotifyLaunchForAttempt(attempt, deadline)
    if (A_TickCount >= deadline) {
        g_SpotifyOpenLastPhase := "budget_exhausted"
        return false
    }
    g_SpotifyOpenLastPhase := "wait_window"
    SpotifyBanner_Set("🔄 Spotify: waiting for window...")
    return SpotifyWaitActivateAndVerify(waitSec, deadline)
}

; Quality gate: main Spotify window exists, is usable, responsive, and is the active foreground window.
SpotifyIsOpenedAndActive(hwnd := 0) {
    if !hwnd
        hwnd := GetSpotifyMainHwnd()
    if !(hwnd is Integer) || hwnd <= 0 || !WinExist("ahk_id " hwnd)
        return false
    if !SpotifyIsUsableMainWindow(hwnd)
        return false
    if !SpotifyIsWindowResponsive(hwnd)
        return false
    return WinActive("ahk_id " hwnd)
}

SpotifyActivateAndVerify(hwnd, deadline := 0) {
    global g_SpotifyOpenLastPhase
    if !(hwnd is Integer) || hwnd <= 0
        return false
    if (!deadline)
        deadline := A_TickCount + SPOTIFY_OPEN_TOTAL_BUDGET_MS
    g_SpotifyOpenLastPhase := "activate"
    SpotifyBanner_Set("🔄 Spotify: verifying...")
    if !SpotifyActivateWindow(hwnd, 3, 300, deadline)
        return false
    SpotifySleepBudget(SPOTIFY_VERIFY_SETTLE_MS, deadline)
    current := GetSpotifyMainHwnd()
    if current
        hwnd := current
    ; Hung window at final verify fails the attempt (escalate) instead of reporting success.
    if !SpotifyIsWindowResponsive(hwnd)
        return false
    return SpotifyIsOpenedAndActive(hwnd)
}

SpotifyWaitActivateAndVerify(timeoutSec := 12, deadline := 0) {
    if (!deadline)
        deadline := A_TickCount + SPOTIFY_OPEN_TOTAL_BUDGET_MS
    remainingMs := deadline - A_TickCount
    if (remainingMs <= 0)
        return false
    cappedSec := Min(timeoutSec, Ceil(remainingMs / 1000.0))
    hwnd := SpotifyWaitForMainHwnd(cappedSec, deadline)
    if hwnd > 0 && SpotifyActivateAndVerify(hwnd, deadline)
        return true
    ; Window may appear slightly after ProcessWait; one short re-check before failing the attempt.
    SpotifySleepBudget(400, deadline)
    hwnd := GetSpotifyMainHwnd()
    return hwnd > 0 && SpotifyActivateAndVerify(hwnd, deadline)
}

; Escalating launch methods per retry (shortcut/exe path may move between installs).
SpotifyLaunchForAttempt(attempt := 1, deadline := 0) {
    if (!deadline)
        deadline := A_TickCount + SPOTIFY_OPEN_TOTAL_BUDGET_MS
    if (attempt = 1) {
        if SpotifyProcessExists()
            SpotifyLaunchRestore()
        else
            SpotifyLaunchFresh(deadline)
        return
    }
    if (attempt = 2) {
        Run("spotify:")
        if !SpotifyProcessExists()
            SpotifyLaunchFresh(deadline)
        return
    }
    exe := SpotifyResolveExePath()
    if exe != ""
        Run('"' exe '"')
    else if (link := GetSpotifyShortcutPath()) != ""
        Run(link)
    else
        Run("shell:AppsFolder\SpotifyAB.SpotifyMusic_zpdnekdrzrea0!Spotify")
    SpotifyProcessWaitBudget(Min(SPOTIFY_OPEN_WAIT_SEC, 8), deadline)
}

SpotifyProcessExists() {
    return ProcessExist("Spotify.exe") > 0
}

SpotifyLaunchFresh(deadline := 0) {
    if (!deadline)
        deadline := A_TickCount + SPOTIFY_OPEN_TOTAL_BUDGET_MS
    link := GetSpotifyShortcutPath()
    if (link != "")
        Run(link)
    else if (exe := SpotifyResolveExePath()) != ""
        Run('"' exe '"')
    else
        Run("shell:AppsFolder\SpotifyAB.SpotifyMusic_zpdnekdrzrea0!Spotify")
    SpotifyProcessWaitBudget(SPOTIFY_FIRST_ATTEMPT_WAIT_SEC, deadline)
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

SpotifyWaitForMainHwnd(timeoutSec := 12, deadline := 0) {
    if (!deadline)
        deadline := A_TickCount + (timeoutSec * 1000)
    localDeadline := A_TickCount + (timeoutSec * 1000)
    if (localDeadline > deadline)
        localDeadline := deadline
    loop {
        hwnd := GetSpotifyMainHwnd()
        if hwnd > 0 {
            ; Hung window counts as not ready yet during wait (keep waiting to deadline).
            if SpotifyIsWindowResponsive(hwnd)
                return hwnd
        }
        if (A_TickCount >= localDeadline)
            break
        Sleep 200
    }
    return 0
}

SpotifyShowOpenFailure() {
    global g_SpotifyOpenLastPhase, g_SpotifyOpenLastAttempt, g_SpotifyOpenStartTick
    elapsedSec := g_SpotifyOpenStartTick > 0 ? Round((A_TickCount - g_SpotifyOpenStartTick) / 1000) : 0
    attempts := g_SpotifyOpenLastAttempt > 0 ? g_SpotifyOpenLastAttempt : SPOTIFY_OPEN_RETRIES
    phase := g_SpotifyOpenLastPhase != "" ? g_SpotifyOpenLastPhase : "unknown"
    SpotifyOpenDebugLog("failure")
    msg := "❌ Spotify did not open (" attempts " attempts, " elapsedSec "s) - " phase
    SpotifyBanner_Result(msg, BANNER_ACCENT_ERROR, 2500)
}

SpotifyActivateWindow(hwnd, attempts := 3, waitMs := 300, deadline := 0) {
    if !(hwnd is Integer) || hwnd <= 0 || !WinExist("ahk_id " hwnd)
        return false
    if (!deadline)
        deadline := A_TickCount + SPOTIFY_OPEN_TOTAL_BUDGET_MS
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
        if (A_TickCount >= deadline)
            return false
        remainingWaitMs := Min(waitMs, Max(50, deadline - A_TickCount))
        waitSec := remainingWaitMs / 1000
        try {
            if (originalState = -1) {
                WinRestore(hwnd)
                SpotifySleepBudget(100, deadline)
            }
            WinActivate(hwnd)
            if WinActive("ahk_id " hwnd) || WinWaitActive("ahk_id " hwnd, , waitSec)
                return true
        } catch {
        }
        try {
            if (originalState = -1) {
                DllCall("ShowWindow", "Ptr", hwnd, "Int", 9)
                SpotifySleepBudget(100, deadline)
            }
            DllCall("SetForegroundWindow", "Ptr", hwnd)
            if WinActive("ahk_id " hwnd) || WinWaitActive("ahk_id " hwnd, , waitSec)
                return true
        } catch {
        }
        try {
            DllCall("BringWindowToTop", "Ptr", hwnd)
            SpotifySleepBudget(100, deadline)
            WinActivate(hwnd)
            if WinActive("ahk_id " hwnd) || WinWaitActive("ahk_id " hwnd, , waitSec)
                return true
        } catch {
        }
        SpotifySleepBudget(200, deadline)
    }
    return false
}

SpotifyIsUsableMainWindow(hwnd) {
    if !(hwnd is Integer) || hwnd <= 0
        return false
    if !DllCall("IsWindowVisible", "Ptr", hwnd)
        return false
    if SpotifyIsWindowCloaked(hwnd)
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

; DWMWA_CLOAKED = 14: UWP/Electron cloaked helpers are not a usable main UI.
SpotifyIsWindowCloaked(hwnd) {
    if !(hwnd is Integer) || hwnd <= 0
        return false
    cloaked := 0
    try {
        hr := DllCall("dwmapi\DwmGetWindowAttribute", "Ptr", hwnd, "UInt", 14, "Int*", &cloaked, "UInt", 4)
        if (hr != 0)
            return false
    } catch {
        return false
    }
    return cloaked != 0
}

; SendMessageTimeout WM_NULL + IsHungAppWindow; false when unresponsive within timeout.
SpotifyIsWindowResponsive(hwnd, timeoutMs := 0) {
    if !(hwnd is Integer) || hwnd <= 0 || !WinExist("ahk_id " hwnd)
        return false
    if (!timeoutMs)
        timeoutMs := SPOTIFY_RESPONSIVE_TIMEOUT_MS
    try {
        if DllCall("IsHungAppWindow", "Ptr", hwnd)
            return false
    } catch {
        ;
    }
    result := 0
    ; SMTO_ABORTIFHUNG = 0x0002
    ok := DllCall("SendMessageTimeout", "Ptr", hwnd, "UInt", 0, "Ptr", 0, "Ptr", 0, "UInt", 0x0002, "UInt",
        timeoutMs, "Ptr*", &result)
    return ok != 0
}

; Main Spotify UI HWND (visible, sized, not cloaked); prefers title containing "Spotify",
; then largest area among usable candidates. Returns 0 if none.
GetSpotifyMainHwnd() {
    bestWithTitle := 0
    bestWithTitleArea := 0
    bestAny := 0
    bestAnyArea := 0
    for hwnd in WinGetList("ahk_exe Spotify.exe") {
        if !SpotifyIsUsableMainWindow(hwnd)
            continue
        area := 0
        try {
            WinGetPos(, , &w, &h, "ahk_id " hwnd)
            area := w * h
        } catch {
            area := 0
        }
        try title := WinGetTitle(hwnd)
        catch
            title := ""
        if InStr(title, "Spotify") {
            if (area > bestWithTitleArea) {
                bestWithTitleArea := area
                bestWithTitle := hwnd
            }
        } else if (area > bestAnyArea) {
            bestAnyArea := area
            bestAny := hwnd
        }
    }
    return bestWithTitle ? bestWithTitle : bestAny
}

; --- Deferred Loading Indication (silent on the fast path) -------------------

SpotifyBanner_Arm(text) {
    global g_SpotifyBannerArmText, g_SpotifyBannerTimerArmed, g_SpotifyBannerVisible
    g_SpotifyBannerArmText := text
    g_SpotifyBannerVisible := false
    try SetTimer(SpotifyBanner_ShowDeferred, 0)
    catch {
    }
    SetTimer(SpotifyBanner_ShowDeferred, -SPOTIFY_BANNER_DELAY_MS)
    g_SpotifyBannerTimerArmed := true
}

SpotifyBanner_ShowDeferred(*) {
    global g_SpotifyBannerArmText, g_SpotifyBannerTimerArmed, g_SpotifyBannerVisible
    g_SpotifyBannerTimerArmed := false
    if (g_SpotifyBannerVisible)
        return
    text := g_SpotifyBannerArmText != "" ? g_SpotifyBannerArmText : "⏳ Opening Spotify..."
    try {
        StandardLoadingBar_Show(text, BANNER_ACCENT_INTERMEDIATE)
        g_SpotifyBannerVisible := true
    } catch {
        ToolTip(text)
        SetTimer(() => ToolTip(), -2000)
    }
}

SpotifyBanner_Set(text) {
    global g_SpotifyBannerArmText, g_SpotifyBannerVisible
    g_SpotifyBannerArmText := text
    if (!g_SpotifyBannerVisible)
        return
    try StandardLoadingBar_Update(text, BANNER_ACCENT_INTERMEDIATE)
    catch {
    }
}

; Disarm timer, clear visible flag, then show Information Only result (order matters so
; a later Cancel / finally cannot kill the result banner mid-display).
SpotifyBanner_Result(text, accent, ms := 1500) {
    global g_SpotifyBannerTimerArmed, g_SpotifyBannerVisible, g_SpotifyBannerArmText
    try SetTimer(SpotifyBanner_ShowDeferred, 0)
    catch {
    }
    g_SpotifyBannerTimerArmed := false
    g_SpotifyBannerVisible := false
    g_SpotifyBannerArmText := ""
    try {
        StandardLoadingBar_Show(text, accent, { passive: true, centerOnHwnd: 0, textWidth: 500, fontSize: 17,
            passiveBgColor: accent })
        if (ms < 1)
            ms := 1
        StandardLoadingBar_Hide(ms)
        StandardLoadingBar_ArmForceHide()
    } catch {
        ToolTip(text)
        SetTimer(() => ToolTip(), -Min(ms, 3000))
    }
}

SpotifyBanner_Cancel() {
    global g_SpotifyBannerTimerArmed, g_SpotifyBannerVisible, g_SpotifyBannerArmText
    try SetTimer(SpotifyBanner_ShowDeferred, 0)
    catch {
    }
    g_SpotifyBannerTimerArmed := false
    g_SpotifyBannerArmText := ""
    if (g_SpotifyBannerVisible) {
        try StandardLoadingBar_Hide(0)
        catch {
        }
    }
    g_SpotifyBannerVisible := false
}

; Sleep that respects the open deadline (returns early when budget exhausted).
SpotifySleepBudget(ms, deadline) {
    if (ms <= 0)
        return
    remaining := deadline - A_TickCount
    if (remaining <= 0)
        return
    Sleep Min(ms, remaining)
}

SpotifyProcessWaitBudget(timeoutSec, deadline) {
    remainingMs := deadline - A_TickCount
    if (remainingMs <= 0)
        return
    cappedSec := Min(timeoutSec, Max(1, Ceil(remainingMs / 1000.0)))
    try
        ProcessWait("Spotify.exe", cappedSec)
    catch {
        ;
    }
}

SpotifyOpenDebugLog(phase, hwnd := 0, attempt := 0) {
    global SPOTIFY_OPEN_DEBUG, g_SpotifyOpenStartTick, g_SpotifyOpenLastPhase, g_SpotifyOpenLastAttempt
    if (!SPOTIFY_OPEN_DEBUG)
        return
    elapsed := g_SpotifyOpenStartTick > 0 ? (A_TickCount - g_SpotifyOpenStartTick) : 0
    if (!attempt)
        attempt := g_SpotifyOpenLastAttempt
    stamp := FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss")
    line := stamp " +" elapsed "ms phase=" phase " last=" g_SpotifyOpenLastPhase " attempt=" attempt " hwnd=" hwnd
    path := A_ScriptDir "\.cursor\spotify_open_quality.log"
    try {
        DirCreate(A_ScriptDir "\.cursor")
        FileAppend(line "`n", path, "UTF-8")
    } catch {
    }
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
