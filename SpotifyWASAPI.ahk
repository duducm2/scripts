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
    return false ; This function is no longer used for setting volume directly.
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

; Adjust volume for the audio session matching pid. deltaPercent is a signed step (e.g. +10 or -10).
; Volume is clamped to [0, 100]. Returns true if session was found and updated, false otherwise.
AdjustProcessVolumeByPid(pid, deltaPercent) {
    return false ; This function is no longer used for setting volume directly.
}

; Set absolute playback volume (0-100) for the audio session matching pid. Does not change Windows master volume.
; Returns true if a session was found and updated.
SetProcessPlaybackVolumePercent(pid, percent) {
    return false ; This function is no longer used for setting volume directly.
}

; Set absolute playback volume for every audio session whose process name contains "AutoHotkey" (any casing).
; Does not change the default device master volume. Returns number of sessions updated (0 if none / error).
; If a process has not opened an audio session yet, it has no mixer entry — call again after first SoundPlay if needed.
ApplyAutoHotkeyAudioSessionsVolumePercent(percent) {
    return 0 ; This function is no longer used for setting volume directly.
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
