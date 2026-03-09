; -----------------------------------------------------------------------------
; WASAPI per-process volume control. Used by Spotify.ahk for silent volume
; adjustment without window activation. Requires Windows Vista+.
; -----------------------------------------------------------------------------

; IID_IAudioSessionManager2   = {77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F}
; IID_IAudioSessionControl2   = {BFB7FF88-7239-4FC9-8FA2-07C950BE9C6D}
; IID_ISimpleAudioVolume      = {BFB7FF88-7239-4FC9-8FA2-07C950BE9C6D}
; CLSID_MMDeviceEnumerator    = {BCDE0395-E52F-467C-8E3D-C4579291692E}
; IID_IMMDeviceEnumerator     = {A95664D2-9614-4F35-A746-DE8DB63617E6}
; IID_IMMDevice              = {D666063F-1587-4E43-81F1-B948E807363F}

; Adjust volume for the audio session matching pid. deltaPercent is a signed step (e.g. +10 or -10).
; Volume is clamped to [0, 100]. Returns true if session was found and updated, false otherwise.
AdjustProcessVolumeByPid(pid, deltaPercent) {
    if !(pid is Integer) || (pid <= 0)
        return false
    try {
        DllCall("ole32\CoInitializeEx", "Ptr", 0, "UInt", 0)
    } catch {
        ; Already initialized is OK
    }
    try {
        mgr := GetDefaultSessionManager()
        if !mgr
            return false
        enum := ComValue(13, 0)
        ComCall(5, mgr, "Ptr*", &enum := 0)
        if !enum
            return false
        enumObj := ComValue(13, enum)
        count := 0
        ComCall(3, enumObj, "UInt*", &count := 0)
        loop count {
            session := 0
            ComCall(4, enumObj, "Int", A_Index - 1, "Ptr*", &session := 0)
            if !session
                continue
            sessionObj := ComValue(13, session)
            ctrl2 := QueryInterface(sessionObj, "{BFB7FF88-7239-4FC9-8FA2-07C950BE9C6D}")
            if !ctrl2
                continue
            sessionPid := 0
            ComCall(7, ctrl2, "UInt*", &sessionPid := 0)
            if (sessionPid != pid)
                continue
            ; Same object implements ISimpleAudioVolume: GetMasterVolume at 9, SetMasterVolume at 8 (0.0-1.0)
            cur := 0.0
            ComCall(9, ctrl2, "Float*", &cur := 0)
            newScalar := (cur * 100) + deltaPercent
            newScalar := Max(0, Min(100, newScalar)) / 100.0
            ComCall(8, ctrl2, "Float", newScalar, "Ptr", 0)
            return true
        }
        return false
    } catch {
        return false
    }
}

QueryInterface(obj, iidStr) {
    iid := Buffer(16)
    DllCall("ole32\IIDFromString", "Str", iidStr, "Ptr", iid)
    out := 0
    ComCall(0, obj, "Ptr", iid, "Ptr*", &out := 0)
    return out ? ComValue(13, out) : 0
}

GetDefaultSessionManager() {
    clsid := Buffer(16)
    DllCall("ole32\CLSIDFromString", "Str", "{BCDE0395-E52F-467C-8E3D-C4579291692E}", "Ptr", clsid)
    iidEnum := Buffer(16)
    DllCall("ole32\IIDFromString", "Str", "{A95664D2-9614-4F35-A746-DE8DB63617E6}", "Ptr", iidEnum)
    pEnum := 0
    if DllCall("ole32\CoCreateInstance", "Ptr", clsid, "Ptr", 0, "UInt", 23, "Ptr", iidEnum, "Ptr*", &pEnum := 0)
        return 0
    enum := ComValue(13, pEnum)
    pDev := 0
    ComCall(4, enum, "Int", 0, "Int", 0, "Ptr*", &pDev := 0)
    if !pDev
        return 0
    dev := ComValue(13, pDev)
    iidMgr := Buffer(16)
    DllCall("ole32\IIDFromString", "Str", "{77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F}", "Ptr", iidMgr)
    pMgr := 0
    ComCall(3, dev, "Ptr", iidMgr, "UInt", 23, "Ptr", 0, "Ptr*", &pMgr := 0)
    if !pMgr
        return 0
    return ComValue(13, pMgr)
}
