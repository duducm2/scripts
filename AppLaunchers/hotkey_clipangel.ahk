; =============================================================================
; AppLaunchers module: hotkey_clipangel.ahk
; #!+. Clip Angel paste and favorite flow; #!+7 tap/hold edit / paste-file
; Extracted verbatim from AppLaunchers.ahk; loaded via #include into the
; AppLaunchers.ahk process, which remains the entry point / source of truth.
; =============================================================================

; Hold threshold matches ZMK hold-tap tapping-term-ms = 200.
CLIPANGEL_WAS7_HOLD_MS := 200

; =============================================================================
; Send specific key combinations
; Hotkey: Win+Alt+Shift+.
; =============================================================================
#!+.::
{
    Sleep(100)
    Send("^c")
    Sleep(200)
    if hwnd := ClipAngel_MainHwnd()
        ClipAngel_ShowWindow(hwnd)
    Sleep(700)
    Send("!q")
    Sleep(200)
    ClipAngel_CloseAndRestoreFocus(0)
}

; =============================================================================
; Win+Alt+Shift+7 — tap: open Clip Angel + Edit (F4); hold 200ms+: Alt+P (non-favorites) then Paste file then hide
; =============================================================================
#!+7::
{
    priorHwnd := 0
    try priorHwnd := WinGetID("A")
    catch {
    }

    pressTime := A_TickCount
    KeyWait "7", "T" . (CLIPANGEL_WAS7_HOLD_MS / 1000)
    isHold := (A_TickCount - pressTime) >= CLIPANGEL_WAS7_HOLD_MS
    if isHold
        try ScriptSoundPlay(A_ScriptDir . "\assets\sounds\copy.wav")

    ClipAngel_WaitChordModifiersReleased()
    ClipAngel_ReleaseChordModifiersForSend()

    if !isHold {
        if !ClipAngel_TryAcquireAutomationLock()
            return
        hideMs := 350
        StandardLoadingBar_Show("⏳ Clip Angel: opening...", BANNER_ACCENT_INTERMEDIATE)
        try {
            if !ActivateClipAngelWithFocusCorrection(true) {
                StandardLoadingBar_Update("❌ Clip Angel is not running.", BANNER_ACCENT_ERROR)
                hideMs := 2000
                return
            }
            hwnd := ClipAngel_MainHwnd()
            if !hwnd {
                StandardLoadingBar_Update("❌ Clip Angel window not found.", BANNER_ACCENT_ERROR)
                hideMs := 2000
                return
            }
            if !ClipAngel_EnsureWindowActive(hwnd) {
                StandardLoadingBar_Update("❌ Clip Angel: failed to activate", BANNER_ACCENT_ERROR)
                hideMs := 2000
                return
            }
            if !ClipAngel_WaitForListReady(CLIPANGEL_FAVORITE_OPEN_READY_MS, true) {
                StandardLoadingBar_Update("❌ Clip Angel: list not ready", BANNER_ACCENT_ERROR)
                hideMs := 2000
                return
            }
            if !WinActive("ahk_id " hwnd) {
                StandardLoadingBar_Update("❌ Clip Angel: lost focus before Edit", BANNER_ACCENT_ERROR)
                hideMs := 2000
                return
            }
            StandardLoadingBar_Update("⏳ Clip Angel: opening editor...", BANNER_ACCENT_INTERMEDIATE)
            SendInput "{F4}"
            StandardLoadingBar_Update("✅ Clip Angel: Edit", BANNER_ACCENT_SUCCESS)
        } finally {
            StandardLoadingBar_Hide(hideMs)
            ClipAngel_ReleaseAutomationLock()
        }
        return
    }

    ; Hold: Alt+P opens non-favorite list (ActivateClipAngelWithFocusCorrection keeps last view, often Alt+B favorites).
    try {
        ClipAngel_ActivateNativeFirstClip(priorHwnd)
        hwnd := ClipAngel_MainHwnd()
        if !hwnd {
            ShowCenteredOverlay_Utils("❌ Clip Angel window not found.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        ClipAngel_EnsureWindowActive(hwnd)
        ClipAngel_WaitForListReady(CLIPANGEL_FAVORITE_OPEN_READY_MS, true)
        if !ClipAngel_InvokePasteEnter(hwnd)
            ShowCenteredOverlay_Utils("❌ Clip Angel Paste file failed", 1500, BANNER_ACCENT_ERROR)
    } finally {
        ClipAngel_CloseAndRestoreFocus(priorHwnd)
    }
}

; =============================================================================
; Initialize Wikipedia scroll position auto-save timer - REMOVED
; =============================================================================
; Auto-save timer removed - now using manual save via Shift keys.ahk shortcut
