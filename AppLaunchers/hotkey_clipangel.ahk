; =============================================================================
; AppLaunchers module: hotkey_clipangel.ahk
; #!+. Clip Angel paste and favorite flow; #!+7 tap/hold edit / paste-file (active folder)
; Extracted verbatim from AppLaunchers.ahk; loaded via #include into the
; AppLaunchers.ahk process, which remains the entry point / source of truth.
; =============================================================================

; Hold threshold matches ZMK hold-tap tapping-term-ms = 200.
CLIPANGEL_WAS7_HOLD_MS := 200

; =============================================================================
; Send specific key combinations
; Hotkey: Win+Alt+Shift+.
; Copy selection then mark newest clip as favorite (shared MarkLastClipAsFavorite path).
; =============================================================================
#!+.::
{
    ClipAngel_WaitChordModifiersReleased()
    ClipAngel_ReleaseChordModifiersForSend()
    prevId := -1
    maxId := ClipAngelDb_MaxId()
    if (maxId != "" && IsInteger(maxId))
        prevId := Integer(maxId)
    Send("^c")
    try ClipWait(0.4)
    catch {
    }
    MarkLastClipAsFavorite("first", true, prevId)
}

; =============================================================================
; Win+Alt+Shift+7 — tap: open Clip Angel + Edit (F4); hold 200ms+: Paste file into
; focused Explorer / Desktop / Cursor·VS Code Files Explorer (Desktop fallback).
; =============================================================================
#!+7::
{
    priorHwnd := 0
    try priorHwnd := WinGetID("A")
    catch {
    }

    holdMs := IsSet(CLIPANGEL_WAS7_HOLD_MS) ? CLIPANGEL_WAS7_HOLD_MS : 200
    pressTime := A_TickCount
    KeyWait "7", "T" . (holdMs / 1000)
    isHold := (A_TickCount - pressTime) >= holdMs

    ClipAngel_WaitChordModifiersReleased()
    ClipAngel_ReleaseChordModifiersForSend()

    if !isHold {
        if !ClipAngel_TryAcquireAutomationLock()
            return
        hideMs := 350
        StandardLoadingBar_Show("⏳ Clip Angel: opening...", BANNER_ACCENT_INTERMEDIATE)
        try {
            if !ClipAngel_OpenNative("p", 0, false) {
                StandardLoadingBar_Update("❌ Clip Angel is not running.", BANNER_ACCENT_ERROR)
                hideMs := 2000
                return
            }
            hwnd := ClipAngel_MainHwnd()
            if (!hwnd || !ClipAngel_EnsureWindowActive(hwnd, 400)) {
                StandardLoadingBar_Update("❌ Clip Angel: failed to activate", BANNER_ACCENT_ERROR)
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
            StandardLoadingBar_Update("✅ Clip Angel: edit", BANNER_ACCENT_SUCCESS)
            hideMs := 350
        } finally {
            StandardLoadingBar_Hide(hideMs)
            ClipAngel_ReleaseAutomationLock()
        }
        return
    }

    ; Hold: paste top clip as file into focused folder target (shared export helper).
    destLabel := "Desktop"
    try destLabel := ClipAngelExport_DescribePasteTarget(priorHwnd)
    catch {
        destLabel := "Desktop"
    }
    try ShowCenteredOverlay_Utils("📎 Paste clip → " destLabel, 1500, BANNER_ACCENT_INFO)
    catch {
    }
    if ClipAngelExport_PasteFirstClipToActiveTarget(priorHwnd) {
        try ScriptSoundPlay(A_ScriptDir . "\assets\sounds\clip-paste-desktop.mp3")
        if (destLabel = "Desktop")
            DesktopCutNewest_CopyPath()
    } else {
        try ShowCenteredOverlay_Utils("❌ Paste file failed", 2000, BANNER_ACCENT_ERROR)
        catch {
        }
    }
}
