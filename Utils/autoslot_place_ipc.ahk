; =============================================================================
; Utils module: autoslot_place_ipc.ahk
; Cross-process request for AutoSlot free-capacity place (Study Topic QuickLook).
; Writer lives in Utils (Shift keys); consumer is WindowManagement AutoSlot.
; Do not call AutoSlot_TryPlaceBackgroundHwnd here — that symbol exists only in the
; WindowManagement process; referencing it under #Warn treats it as an unassigned local.
; =============================================================================

; Must match AutoSlot.ahk (RegisterWindowMessage string).
AUTOSLOT_PLACE_MSG_NAME := "EDU_AutoSlot_PlaceHwnd"
global g_AutoSlotPlaceRequestFile := A_ScriptDir "\.cursor\autoslot_place_request"
global g_AutoSlotPlaceMsg := 0

; QL BeginShow.PositionWindow undoes external resize unless WPF WindowState=Maximized;
; several deferred passes cover paint / plugin.View races.
QL_AUTOSLOT_PLACE_MS_1 := 400
QL_AUTOSLOT_PLACE_MS_2 := 1200
QL_AUTOSLOT_PLACE_MS_3 := 2500
QL_AUTOSLOT_PLACE_MS_4 := 4000

AutoSlot_PlaceMsgId() {
    global g_AutoSlotPlaceMsg
    if (!g_AutoSlotPlaceMsg)
        g_AutoSlotPlaceMsg := DllCall("RegisterWindowMessage", "str", AUTOSLOT_PLACE_MSG_NAME, "uint")
    return g_AutoSlotPlaceMsg
}

; Hidden AutoHotkey main window for WindowManagement.ahk (same scripts folder).
AutoSlot_FindWindowManagementHwnd() {
    prevDetect := A_DetectHiddenWindows
    prevMatch := A_TitleMatchMode
    DetectHiddenWindows true
    SetTitleMatchMode 2
    hwnd := 0
    try hwnd := WinExist("WindowManagement.ahk ahk_class AutoHotkey")
    catch
        hwnd := 0
    DetectHiddenWindows prevDetect
    SetTitleMatchMode prevMatch
    return hwnd
}

; Queue hwnd for AutoSlot free-capacity place (empty max / free half 50/50).
; Prefer PostMessage (reliable on Google Drive paths); file is fallback only.
; Consumer polls the file at most every AutoSlot_PLACE_REQUEST_POLL_MS (1000 ms).
AutoSlot_RequestPlaceCrossProcess(hwnd) {
    if (!hwnd)
        return false
    hwnd := Integer(hwnd)
    if (!DllCall("IsWindow", "ptr", hwnd))
        return false
    global g_AutoSlotPlaceRequestFile
    msg := AutoSlot_PlaceMsgId()
    wmHwnd := AutoSlot_FindWindowManagementHwnd()
    if (msg && wmHwnd) {
        try {
            PostMessage msg, 0, hwnd, , "ahk_id " wmHwnd
            return true
        } catch {
        }
    }
    try {
        cursorDir := A_ScriptDir "\.cursor"
        if !DirExist(cursorDir)
            DirCreate(cursorDir)
        if FileExist(g_AutoSlotPlaceRequestFile)
            FileDelete(g_AutoSlotPlaceRequestFile)
        FileAppend(hwnd "`n", g_AutoSlotPlaceRequestFile, "UTF-8")
    } catch {
        return false
    }
    return true
}

; After study layout: defer place until QL finishes load/size settle (four passes).
; Re-resolves ahk_exe QuickLook.exe each fire so stale hwnds are not used.
QuickLook_ScheduleAutoSlotPlace(*) {
    SetTimer(QuickLook_DeferredAutoSlotPlacePass1, 0)
    SetTimer(QuickLook_DeferredAutoSlotPlacePass2, 0)
    SetTimer(QuickLook_DeferredAutoSlotPlacePass3, 0)
    SetTimer(QuickLook_DeferredAutoSlotPlacePass4, 0)
    SetTimer(QuickLook_DeferredAutoSlotPlacePass1, -QL_AUTOSLOT_PLACE_MS_1)
    SetTimer(QuickLook_DeferredAutoSlotPlacePass2, -QL_AUTOSLOT_PLACE_MS_2)
    SetTimer(QuickLook_DeferredAutoSlotPlacePass3, -QL_AUTOSLOT_PLACE_MS_3)
    SetTimer(QuickLook_DeferredAutoSlotPlacePass4, -QL_AUTOSLOT_PLACE_MS_4)
}

QuickLook_DeferredAutoSlotPlacePass1(*) {
    QuickLook_FireAutoSlotPlace()
}

QuickLook_DeferredAutoSlotPlacePass2(*) {
    QuickLook_FireAutoSlotPlace()
}

QuickLook_DeferredAutoSlotPlacePass3(*) {
    QuickLook_FireAutoSlotPlace()
}

QuickLook_DeferredAutoSlotPlacePass4(*) {
    QuickLook_FireAutoSlotPlace()
}

QuickLook_FireAutoSlotPlace(*) {
    hwnd := WinExist("ahk_exe QuickLook.exe")
    if (!hwnd)
        return
    AutoSlot_RequestPlaceCrossProcess(hwnd)
}
