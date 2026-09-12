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
; Win+Alt+Shift+7 — tap: open Clip Angel + Edit (F4); hold 200ms+: save clipboard to Desktop (or ClipAngel Paste file fallback)
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
            hwnd := 0
            root := 0
            if !ClipAngel_OpenForAutomation("all", 0, false, &hwnd, &root) {
                StandardLoadingBar_Update("❌ Clip Angel is not running.", BANNER_ACCENT_ERROR)
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
            StandardLoadingBar_Update("✅ Clip Angel: edit", BANNER_ACCENT_SUCCESS)
            hideMs := 350
        } finally {
            StandardLoadingBar_Hide(hideMs)
            ClipAngel_ReleaseAutomationLock()
        }
        return
    }

    ; Hold: paste top clip as file onto Desktop (shared export helper).
    try ShowCenteredOverlay_Utils("📎 Paste clip to Desktop", 1500, BANNER_ACCENT_INFO)
    catch {
    }
    ClipAngelExport_PasteFirstClipToDesktop()
}
