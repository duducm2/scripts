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
    settingsFile := A_ScriptDir . "\data\settings.ini"
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
    try ScriptSoundPlay(A_ScriptDir . "\sounds\api-success.mp3")
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

; Toggle sound state and show visual feedback
ToggleSoundState() {
    settingsFile := A_ScriptDir . "\data\settings.ini"
    currentState := IsSoundEnabled()
    newState := currentState ? "0" : "1"

    ; Update INI file
    IniWrite(newState, settingsFile, "Settings", "SoundEnabled")

    ; Show visual feedback
    if (newState = "1") {
        ShowCenteredOverlay_Utils("🔊 Sound: ON", 2000, BANNER_ACCENT_INTERMEDIATE)
    } else {
        ShowCenteredOverlay_Utils("🔇 Sound: OFF", 2000, BANNER_ACCENT_INTERMEDIATE)
    }
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
    micVolumeScript := A_ScriptDir "\scripts\Set-MicVolume.ps1"
    if (!FileExist(micVolumeScript))
        return
    try {
        Run("powershell.exe -ExecutionPolicy Bypass -File `"" micVolumeScript "`" -Level " SCRIPT_MIC_CAPTURE_VOLUME_PERCENT, ,
            "Hide")
    } catch {
    }
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
