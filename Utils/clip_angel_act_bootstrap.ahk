; =============================================================================
; Utils/clip_angel_act_bootstrap.ahk
; Thin Clip Angel stack for Act.ahk only — enough to call ClipAngel_EnsureRunning
; (soft open, else Macros [R] hard restart). Do NOT use as a hotkey host.
; Mirrors Spotify.ahk thin-include pattern: stubs avoid full Utils.ahk / Escape host.
; =============================================================================

#include %A_ScriptDir%\vendor\UIA-v2\Lib\UIA.ahk

; activate.ahk RegisterMacro at load — no-op (Act is not a macros host).
RegisterMacro(func, title, char := "") {
}

; favorite.ahk favorite-success paths only; Act ensure never marks favorites.
ScriptSoundPlay(path, wait := false) {
    return false
}

; Same helper as Utils\handy_selector_entry.ahk (avoid pulling Handy selector).
ShowCenteredOverlay_Utils(text, duration := 1500, bgColor := BANNER_ACCENT_INTERMEDIATE) {
    StandardLoadingBar_Show(text, bgColor, { passive: true, centerOnHwnd: 0, textWidth: 500, fontSize: 17,
        passiveBgColor: bgColor })
    if (duration < 1)
        duration := 1
    StandardLoadingBar_Hide(duration)
    StandardLoadingBar_ArmForceHide()
}

; Monitor/layout helpers used by ClipAngel_ApplyLayoutOnMonitor (from peek_pdf_study_03).
MoveWindowToMonitor(hwnd, monitorIndex := 2) {
    if (!hwnd)
        return
    try {
        MonitorGetWorkArea(monitorIndex, &l, &t, &r, &b)
    } catch {
        return
    }
    w := r - l
    h := b - t
    try {
        mm := WinGetMinMax("ahk_id " hwnd)
        if (mm != 0) {
            WinRestore("ahk_id " hwnd)
            Sleep 80
        }
    } catch {
    }
    try WinMove(l, t, w, h, "ahk_id " hwnd)
}

TryMaximizeWindow(hwnd) {
    if (!hwnd)
        return false
    try {
        WinMaximize("ahk_id " hwnd)
        return true
    } catch {
        try {
            PostMessage 0x0112, 0xF030, , , "ahk_id " hwnd
            return true
        } catch {
            return false
        }
    }
}

GetAhkMonitorIndexFromHwnd(hwnd) {
    if (!hwnd)
        return 0
    hMon := 0
    try hMon := DllCall("user32\MonitorFromWindow", "ptr", hwnd, "uint", 2, "ptr")
    catch
        return 0
    if (!hMon)
        return 0
    count := 0
    try count := MonitorGetCount()
    catch
        return 0
    if (count < 1)
        return 0
    loop count {
        i := A_Index
        try MonitorGet i, &l, &t, &r, &b
        catch
            continue
        cx := (l + r) // 2
        cy := (t + b) // 2
        pt64 := ((cy & 0xFFFFFFFF) << 32) | (cx & 0xFFFFFFFF)
        cur := 0
        try cur := DllCall("user32\MonitorFromPoint", "int64", pt64, "uint", 2, "ptr")
        catch
            cur := 0
        if (cur = hMon)
            return i
    }
    return 0
}

; favorite.ahk references ClipAngelDb_* at load (#Warn); include before it (same as Utils.ahk).
#include %A_ScriptDir%\Utils\clip_angel_db.ahk
#include %A_ScriptDir%\Utils\clip_angel_favorite.ahk
#include %A_ScriptDir%\Utils\clip_angel_activate.ahk