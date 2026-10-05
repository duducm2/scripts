; =============================================================================
; Utils module: global_sound_audio.ahk
; Global sound toggle and script audio helpers
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; Global Sound Toggle System
; File-backed state management for muting/unmuting sounds across all scripts
; =============================================================================

; Check if sound is enabled (reads from INI file for cross-process persistence)
IsSoundEnabled() {
    settingsFile := A_ScriptDir . "\assets\data\settings.ini"
    ; Default to enabled (1) if file doesn't exist or key is missing
    soundEnabled := IniRead(settingsFile, "Settings", "SoundEnabled", "1")
    return (soundEnabled = "1")
}

; -----------------------------------------------------------------------------
; Central gate for the global sound toggle: all script-triggered audio must use
; these helpers (SoundPlay, WMP, SoundBeep, MessageBeep, system scheme sounds).
; When IsSoundEnabled() is false, these are no-ops.
; -----------------------------------------------------------------------------
ScriptSoundPlay(path, wait := false) {
    if (!IsSoundEnabled())
        return false
    try {
        SoundPlay(path, wait)
        return true
    } catch {
        return false
    }
}

; System scheme sounds, e.g. *64 (asterisk), *16 (exclamation) - see SoundPlay docs.
ScriptSoundPlaySystem(scheme) {
    if (!IsSoundEnabled())
        return false
    try {
        SoundPlay(scheme, false)
        return true
    } catch {
        return false
    }
}

; Study subtopic / article link API save success (Manage Study Link flows).
StudyLink_PlayApiSuccessSound() {
    try ScriptSoundPlay(A_ScriptDir . "\assets\sounds\api-success.mp3")
}

ScriptSoundBeep(freq, duration) {
    if (!IsSoundEnabled())
        return false
    try {
        SoundBeep(freq, duration)
        return true
    } catch {
        return false
    }
}

ScriptMessageBeep(type := 0xFFFFFFFF) {
    if (!IsSoundEnabled())
        return false
    try {
        return DllCall("User32\MessageBeep", "UInt", type)
    } catch {
        return false
    }
}

; Toggle AutoHotkey mute in the Windows volume mixer (the per-app speaker icon).
; Keeps SoundEnabled in step so script chimes follow that same mute.
ToggleSoundState() {
    settingsFile := A_ScriptDir . "\assets\data\settings.ini"
    state := AutoHotkeyAudioSessionsMuteState()
    if (state = "")
        wantMute := IsSoundEnabled()
    else
        wantMute := (state = "unmuted")
    n := SetAutoHotkeyAudioSessionsMute(wantMute)
    if (n = 0 && state != "") {
        ShowCenteredOverlay_Utils("❌ Could not change AutoHotkey mute", 2000, BANNER_ACCENT_ERROR)
        return
    }
    IniWrite(wantMute ? "0" : "1", settingsFile, "Settings", "SoundEnabled")
    if (wantMute)
        ShowCenteredOverlay_Utils("🔇 AutoHotkey muted", 2000, BANNER_ACCENT_INTERMEDIATE)
    else
        ShowCenteredOverlay_Utils("🔊 AutoHotkey unmuted", 2000, BANNER_ACCENT_INTERMEDIATE)
}

; =============================================================================
; Centralized audio levels (AHK playback app volume vs mic capture - not Windows master)
; =============================================================================
global SCRIPT_MASTER_VOLUME_PERCENT := 40
global SCRIPT_MIC_CAPTURE_VOLUME_PERCENT := 100
global SCRIPT_MICROPHONE_INPUT_SLIDER_PERCENT := 100

; Per-process AutoHotkey playback volume via WASAPI (see ApplyAutoHotkeyAudioSessionsVolumePercent).
; Does not call SoundSetVolume - leaves the default device master volume unchanged.
ApplyScriptMasterVolumeTarget() {
    return 0 ; Only SetAutoHotkeyVolume.ps1 should set volume now.
}

; Apply now and again after delays - audio sessions for new AutoHotkey processes often do not exist for hundreds of ms after Start-Process (Quick Update / multi-script startup).
; Use distinct timer callbacks (lambdas): SetTimer with the *same* function reference replaces the previous timer - two SetTimer(ApplyScriptMasterVolumeTarget, ...) would only keep the last delay.
ScheduleApplyScriptMasterVolumeTargetWithRetries() {
    ApplyScriptMasterVolumeTarget()
    SetTimer(() => ApplyScriptMasterVolumeTarget(), -2500)
    SetTimer(() => ApplyScriptMasterVolumeTarget(), -6000)
}

; Run after Quick Update relaunch only: AppLaunchers starts last with /Updated; other scripts need time to spawn audio sessions.
ScheduleApplyScriptMasterVolumeTargetAfterQuickUpdate() {
    ApplyScriptMasterVolumeTarget()
    SetTimer(() => ApplyScriptMasterVolumeTarget(), -2000)
    SetTimer(() => ApplyScriptMasterVolumeTarget(), -5000)
    SetTimer(() => ApplyScriptMasterVolumeTarget(), -10000)
    SetTimer(() => ApplyScriptMasterVolumeTarget(), -15000)
    SetTimer(() => ApplyScriptMasterVolumeTarget(), -20000)
    SetTimer(() => ApplyScriptMasterVolumeTarget(), -25000)
}

RunSetMicVolumeScript() {
    RunMicVolumeScript("")
}

; muteAction: "" sets capture level, "mute" cuts the endpoint off, "unmute" clears mute and restores the level.
RunMicVolumeScript(muteAction := "") {
    global SCRIPT_MIC_CAPTURE_VOLUME_PERCENT
    micVolumeScript := A_ScriptDir "\infra\tools\Set-MicVolume.ps1"
    if (!FileExist(micVolumeScript))
        return
    args := "-Level " SCRIPT_MIC_CAPTURE_VOLUME_PERCENT
    if (muteAction = "mute")
        args := "-Mute"
    else if (muteAction = "unmute")
        args := "-Unmute -Level " SCRIPT_MIC_CAPTURE_VOLUME_PERCENT
    try {
        Run("powershell.exe -ExecutionPolicy Bypass -File `"" micVolumeScript "`" " args, , "Hide")
    } catch {
    }
}

Dictation_MuteCaptureForSwap() {
    global g_DictationSwapMicMuted
    g_DictationSwapMicMuted := true
    RunMicVolumeScript("mute")
}

; No-op unless this swap muted the endpoint. Safe to call from finally after a successful unmute.
Dictation_UnmuteCaptureAfterSwap() {
    global g_DictationSwapMicMuted
    if (!g_DictationSwapMicMuted)
        return
    g_DictationSwapMicMuted := false
    RunMicVolumeScript("unmute")
}

; =============================================================================
; Outlook: classic OUTLOOK.EXE and Microsoft Store "new" Outlook (olk.exe)
; =============================================================================
OutlookGetOlkExePath() {
    candidate :=
        "C:\Program Files\WindowsApps\Microsoft.OutlookForWindows_1.2026.317.100_x64__8wekyb3d8bbwe\olk.exe"
    if FileExist(candidate)
        return candidate
    try {
        loop files "C:\Program Files\WindowsApps\Microsoft.OutlookForWindows_*_x64__8wekyb3d8bbwe\olk.exe", "F" {
            return A_LoopFileFullPath
        }
    } catch {
    }
    return ""
}

OutlookProcessRunning() {
    return ProcessExist("OUTLOOK.EXE") || ProcessExist("olk.exe")
}
