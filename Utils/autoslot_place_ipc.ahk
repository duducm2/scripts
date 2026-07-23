; =============================================================================
; Utils module: autoslot_place_ipc.ahk
; Cross-process request for AutoSlot free-capacity place (Study Topic QuickLook).
; Writer lives in Utils (Shift keys); consumer polls in WindowManagement AutoSlot.
; Do not call AutoSlot_TryPlaceBackgroundHwnd here — that symbol exists only in the
; WindowManagement process; referencing it under #Warn treats it as an unassigned local.
; =============================================================================

global g_AutoSlotPlaceRequestFile := A_ScriptDir "\.cursor\autoslot_place_request"

QL_AUTOSLOT_PLACE_MS_1 := 700
QL_AUTOSLOT_PLACE_MS_2 := 1800

; Queue hwnd for AutoSlot free-capacity place (empty max / free half 50/50).
; WindowManagement AutoSlot_CheckPlaceRequest polls and runs TryPlaceBackgroundHwnd.
AutoSlot_RequestPlaceCrossProcess(hwnd) {
    if (!hwnd)
        return false
    hwnd := Integer(hwnd)
    if (!DllCall("IsWindow", "ptr", hwnd))
        return false
    global g_AutoSlotPlaceRequestFile
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

; After study layout: defer place until QL finishes load/size settle (two passes).
; Re-resolves ahk_exe QuickLook.exe each fire so stale hwnds are not used.
QuickLook_ScheduleAutoSlotPlace(*) {
    SetTimer(QuickLook_DeferredAutoSlotPlacePass1, 0)
    SetTimer(QuickLook_DeferredAutoSlotPlacePass2, 0)
    SetTimer(QuickLook_DeferredAutoSlotPlacePass1, -QL_AUTOSLOT_PLACE_MS_1)
    SetTimer(QuickLook_DeferredAutoSlotPlacePass2, -QL_AUTOSLOT_PLACE_MS_2)
}

QuickLook_DeferredAutoSlotPlacePass1(*) {
    QuickLook_FireAutoSlotPlace()
}

QuickLook_DeferredAutoSlotPlacePass2(*) {
    QuickLook_FireAutoSlotPlace()
}

QuickLook_FireAutoSlotPlace(*) {
    hwnd := WinExist("ahk_exe QuickLook.exe")
    if (!hwnd)
        return
    AutoSlot_RequestPlaceCrossProcess(hwnd)
}
