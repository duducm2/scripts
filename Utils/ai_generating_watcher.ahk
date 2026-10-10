; =============================================================================
; Utils module: ai_generating_watcher.ahk
; One Stop-button answer shared by the pack pipeline and dictation-to-cursor.
; Callers keep their own completion state machines and ask this module
; whether the companion is still generating.
; =============================================================================

global g_AiGenWatchKey := ""
global g_AiGenWatchAt := 0
global g_AiGenWatchVal := false

AiGeneratingWatcher_IsGenerating(hwnd, companion, probe) {
    global g_AiGenWatchKey, g_AiGenWatchAt, g_AiGenWatchVal
    if (!hwnd || !WinExist("ahk_id " hwnd) || !IsObject(probe))
        return false
    key := hwnd . "|" . StrLower(Trim(companion))
    if (key = g_AiGenWatchKey && (A_TickCount - g_AiGenWatchAt) < 450)
        return g_AiGenWatchVal
    val := false
    try val := !!probe.Call(hwnd, companion)
    catch {
        val := false
    }
    g_AiGenWatchKey := key
    g_AiGenWatchAt := A_TickCount
    g_AiGenWatchVal := val
    return val
}
