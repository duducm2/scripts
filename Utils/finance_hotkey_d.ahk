; =============================================================================
; Utils module: finance_hotkey_d.ahk
; Win+Alt+Shift+D tap / double-tap / hold (400 ms = AI_QD_DOUBLE_TAP_MS / ZMK tap-dance):
;   1× = Tasks (Utility Shortcuts [T])
;   2× = Finance (Utility Shortcuts [F])
;   hold 700ms+ = Memory Palace (Utility Shortcuts [N])
; Pattern mirrors WindowManagement\audio_bt_menu.ahk #!+9.
; =============================================================================

FINANCE_D_HOLD_MS := 700
global g_FinanceD_DoubleTapArmed := false
global g_FinanceD_LastPressTick := 0
global g_FinanceD_DoubleTapTimer := 0

class FinanceD_DoubleTapTimerObj {
    static OnSingleTapTimeout() {
        global g_FinanceD_DoubleTapArmed, g_FinanceD_DoubleTapTimer
        if (!g_FinanceD_DoubleTapArmed)
            return
        g_FinanceD_DoubleTapArmed := false
        g_FinanceD_DoubleTapTimer := 0
        Task_LaunchApp()
    }
}

FinanceD_DisarmDoubleTap() {
    global g_FinanceD_DoubleTapArmed, g_FinanceD_DoubleTapTimer
    global g_FinanceD_LastPressTick
    g_FinanceD_DoubleTapArmed := false
    g_FinanceD_LastPressTick := 0
    if (g_FinanceD_DoubleTapTimer) {
        SetTimer(g_FinanceD_DoubleTapTimer, 0)
        g_FinanceD_DoubleTapTimer := 0
    }
}

#!+d:: {
    global g_FinanceD_DoubleTapArmed, g_FinanceD_LastPressTick, g_FinanceD_DoubleTapTimer

    ; Hotkey fires on key-down. Drop queued auto-repeat ghosts that run after a hold
    ; released (those start with D already up and would otherwise arm single-tap Tasks).
    if !GetKeyState("d", "P")
        return

    thresholdMs := 400
    try thresholdMs := AI_QD_DOUBLE_TAP_MS
    catch {
        thresholdMs := 400
    }

    ; Detect double-tap on key-down (before KeyWait) so a slow release cannot miss the window.
    pressTime := A_TickCount
    elapsed := (g_FinanceD_LastPressTick > 0) ? (pressTime - g_FinanceD_LastPressTick) : 9999
    isSecondTap := g_FinanceD_DoubleTapArmed && elapsed >= 0 && elapsed < thresholdMs

    KeyWait "d", "T" . (FINANCE_D_HOLD_MS / 1000)
    isHold := (A_TickCount - pressTime) >= FINANCE_D_HOLD_MS

    if (isHold) {
        FinanceD_DisarmDoubleTap()
        Palace_LaunchApp()
        ; Stay in this thread until physical release so a repeat cannot start mid-hold
        ; and arm single-tap after we return.
        KeyWait "d"
        return
    }

    if (isSecondTap) {
        FinanceD_DisarmDoubleTap()
        Finance_LaunchApp()
        return
    }

    ; Arm after quick release: window starts now (matches #!+9).
    if (g_FinanceD_DoubleTapTimer) {
        SetTimer(g_FinanceD_DoubleTapTimer, 0)
        g_FinanceD_DoubleTapTimer := 0
    }
    g_FinanceD_LastPressTick := A_TickCount
    g_FinanceD_DoubleTapArmed := true
    g_FinanceD_DoubleTapTimer := ObjBindMethod(FinanceD_DoubleTapTimerObj, "OnSingleTapTimeout")
    SetTimer(g_FinanceD_DoubleTapTimer, -thresholdMs)
}
