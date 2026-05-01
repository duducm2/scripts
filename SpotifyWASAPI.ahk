; -----------------------------------------------------------------------------
; WASAPI per-process volume control. Used by Spotify.ahk for silent volume
; adjustment without window activation. Requires Windows Vista+.
; -----------------------------------------------------------------------------

; IID_IAudioSessionManager2   = {77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F}
; IID_IAudioSessionControl2   = {BFB7FF88-7239-4FC9-8FA2-07C950BE9C6D}
; IID_ISimpleAudioVolume      = {F8679F50-25A0-4E46-9F29-B28F71F0019C}  ; audioclient.h — not IAudioSessionControl2
; CLSID_MMDeviceEnumerator    = {BCDE0395-E52F-467C-8E3D-C4579291692E}
; IID_IMMDeviceEnumerator     = {A95664D2-9614-4F35-A746-DE8DB63617E6}
; IID_IMMDevice              = {D666063F-1587-4E43-81F1-B948E807363F}

global WASAPI_IID_IAudioSessionControl2 := "{BFB7FF88-7239-4FC9-8FA2-07C950BE9C6D}"
global WASAPI_IID_ISimpleAudioVolume := "{F8679F50-25A0-4E46-9F29-B28F71F0019C}"

; IAudioSessionControl2::GetProcessId is vtable index 14 (IUnknown 0–2, IAudioSessionControl 3–11, IAudioSessionControl2 12–16).
; Use ComObjQuery for QI — manual IUnknown::QueryInterface via ComCall throws E_NOINTERFACE (0x80004002) on failure.
GetAudioSessionProcessId(sessionObj) {
    try {
        ctrl2 := ComObjQuery(sessionObj, WASAPI_IID_IAudioSessionControl2)
        if !ctrl2
            return 0
        sessionPid := 0
        try ComCall(14, ctrl2, "UInt*", &sessionPid := 0)
        catch {
            return 0
        }
        return sessionPid
    } catch {
        return 0
    }
}

; ISimpleAudioVolume: SetMasterVolume = slot 3, GetMasterVolume = slot 4 (after IUnknown).
; Primary: QI session -> ISimpleAudioVolume. Fallback: GetGroupingParam (IAudioSessionControl slot 8) +
; IAudioSessionManager::GetSimpleAudioVolume (slot 4) — QI from enumerator often does not expose ISimpleAudioVolume on Win10/11.
WASAPI_SetSessionScalar(sessionObj, scalar, mgr := 0) {
    scalar := Float(scalar)
    if (scalar < 0.0)
        scalar := 0.0
    if (scalar > 1.0)
        scalar := 1.0
    try {
        vol := ComObjQuery(sessionObj, WASAPI_IID_ISimpleAudioVolume)
        if vol {
            ComCall(3, vol, "Float", scalar)
            return true
        }
    } catch {
    }
    if !mgr
        return false
    try {
        guidBuf := Buffer(16, 0)
        ComCall(8, sessionObj, "Ptr", guidBuf.Ptr)
        pVol := 0
        ComCall(4, mgr, "Ptr", guidBuf.Ptr, "UInt", 0, "Ptr*", &pVol := 0)
        if !pVol
            return false
        vol := ComValue(13, pVol)
        ComCall(3, vol, "Float", scalar)
        return true
    } catch {
        return false
    }
}

WASAPI_GetSessionScalar(sessionObj, &outScalar, mgr := 0) {
    try {
        vol := ComObjQuery(sessionObj, WASAPI_IID_ISimpleAudioVolume)
        if vol {
            ComCall(4, vol, "Float*", &outScalar := 0.0)
            return true
        }
    } catch {
    }
    if !mgr
        return false
    try {
        guidBuf := Buffer(16, 0)
        ComCall(8, sessionObj, "Ptr", guidBuf.Ptr)
        pVol := 0
        ComCall(4, mgr, "Ptr", guidBuf.Ptr, "UInt", 0, "Ptr*", &pVol := 0)
        if !pVol
            return false
        vol := ComValue(13, pVol)
        ComCall(4, vol, "Float*", &outScalar := 0.0)
        return true
    } catch {
        return false
    }
}

; IAudioSessionManager: GetAudioSessionControl=3, Register=4, Unregister=5.
; IAudioSessionManager2: GetSessionEnumerator=6.
; IAudioSessionEnumerator: GetCount=3, GetSession=4.

; Adjust volume for the audio session matching pid. deltaPercent is a signed step in percentage points (e.g. +5 or -5).
; Volume is clamped to [0, 100] scalar. Returns true if session was found and updated, false otherwise.
AdjustProcessVolumeByPid(pid, deltaPercent) {
    if !(pid is Integer) || pid <= 0
        return false
    mgr := GetDefaultSessionManager()
    if !mgr
        return false
    pEnum := 0
    try ComCall(6, mgr, "Ptr*", &pEnum := 0)
    catch {
        return false
    }
    if !pEnum
        return false
    enum := ComValue(13, pEnum)
    count := 0
    try ComCall(3, enum, "Int*", &count := 0)
    catch {
        try ObjRelease(enum)
        return false
    }
    updated := false
    loop count {
        idx := A_Index - 1
        pSess := 0
        try ComCall(4, enum, "Int", idx, "Ptr*", &pSess := 0)
        catch {
            continue
        }
        if !pSess
            continue
        sess := ComValue(13, pSess)
        sessPid := GetAudioSessionProcessId(sess)
        if (sessPid != pid) {
            try ObjRelease(sess)
            continue
        }
        cur := 0.0
        if !WASAPI_GetSessionScalar(sess, &cur, mgr) {
            try ObjRelease(sess)
            continue
        }
        pct := Round(cur * 100.0) + deltaPercent
        if (pct < 0)
            pct := 0
        if (pct > 100)
            pct := 100
        newScalar := pct / 100.0
        if WASAPI_SetSessionScalar(sess, newScalar, mgr)
            updated := true
        try ObjRelease(sess)
        break
    }
    try ObjRelease(enum)
    return updated
}

; Set absolute playback volume (0-100) for the audio session matching pid. Does not change Windows master volume.
; Returns true if a session was found and updated.
SetProcessPlaybackVolumePercent(pid, percent) {
    if !(pid is Integer) || pid <= 0
        return false
    pct := Integer(percent)
    if (pct < 0)
        pct := 0
    if (pct > 100)
        pct := 100
    scalar := pct / 100.0
    mgr := GetDefaultSessionManager()
    if !mgr
        return false
    pEnum := 0
    try ComCall(6, mgr, "Ptr*", &pEnum := 0)
    catch {
        return false
    }
    if !pEnum
        return false
    enum := ComValue(13, pEnum)
    count := 0
    try ComCall(3, enum, "Int*", &count := 0)
    catch {
        try ObjRelease(enum)
        return false
    }
    loop count {
        idx := A_Index - 1
        pSess := 0
        try ComCall(4, enum, "Int", idx, "Ptr*", &pSess := 0)
        catch {
            continue
        }
        if !pSess
            continue
        sess := ComValue(13, pSess)
        if (GetAudioSessionProcessId(sess) != pid) {
            try ObjRelease(sess)
            continue
        }
        ok := WASAPI_SetSessionScalar(sess, scalar, mgr)
        try ObjRelease(sess)
        try ObjRelease(enum)
        return ok ? true : false
    }
    try ObjRelease(enum)
    return false
}

; Set absolute playback volume for every audio session whose process name contains "AutoHotkey" (any casing).
; Does not change the default device master volume. Returns number of sessions updated (0 if none / error).
ApplyAutoHotkeyAudioSessionsVolumePercent(percent) {
    pct := Integer(percent)
    if (pct < 0)
        pct := 0
    if (pct > 100)
        pct := 100
    scalar := pct / 100.0
    mgr := GetDefaultSessionManager()
    if !mgr
        return 0
    pEnum := 0
    try ComCall(6, mgr, "Ptr*", &pEnum := 0)
    catch {
        return 0
    }
    if !pEnum
        return 0
    enum := ComValue(13, pEnum)
    count := 0
    try ComCall(3, enum, "Int*", &count := 0)
    catch {
        try ObjRelease(enum)
        return 0
    }
    n := 0
    loop count {
        idx := A_Index - 1
        pSess := 0
        try ComCall(4, enum, "Int", idx, "Ptr*", &pSess := 0)
        catch {
            continue
        }
        if !pSess
            continue
        sess := ComValue(13, pSess)
        sessPid := GetAudioSessionProcessId(sess)
        if !sessPid {
            try ObjRelease(sess)
            continue
        }
        procName := ""
        try procName := ProcessGetName(sessPid)
        catch {
            try ObjRelease(sess)
            continue
        }
        if !InStr(StrLower(procName), "autohotkey") {
            try ObjRelease(sess)
            continue
        }
        if WASAPI_SetSessionScalar(sess, scalar, mgr)
            n++
        try ObjRelease(sess)
    }
    try ObjRelease(enum)
    return n
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
