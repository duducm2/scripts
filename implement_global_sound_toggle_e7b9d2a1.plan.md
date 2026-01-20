---
name: Implement Global Sound Toggle
overview: Implement a system-wide sound toggle using a file-backed state to mute/unmute audio feedback across all AutoHotkey scripts, integrated into the Selection Modal.
todos:
  - id: create_sound_state_manager
    content: Implement GetSoundState and ToggleSoundState functions in Utils.ahk using an INI file for cross-process persistence.
    status: pending
    dependencies: []
  - id: register_toggle_macro
    content: Register the "Toggle Sound" macro in InitMacros inside Utils.ahk with a hotkey assignment.
    status: pending
    dependencies: [create_sound_state_manager]
  - id: refactor_utils_sounds
    content: Wrap all sound playback calls in Utils.ahk with the IsSoundEnabled check.
    status: pending
    dependencies: [create_sound_state_manager]
  - id: refactor_shiftkeys_sounds
    content: Wrap all sound playback calls in Shift keys.ahk with the IsSoundEnabled check.
    status: pending
    dependencies: [create_sound_state_manager]
  - id: refactor_external_scripts_sounds
    content: Wrap sound playback calls in Microsoft Teams.ahk, Gemini.ahk, and AppLaunchers.ahk with the IsSoundEnabled check.
    status: pending
    dependencies: [create_sound_state_manager]
---

# Implement Global Sound Toggle

## Analysis / Context
The current codebase contains approximately 14 instances of sound reproduction (beeps, WAV files, system sounds) distributed across 5 different script files (`Utils.ahk`, `AppLaunchers.ahk`, `Gemini.ahk`, `Microsoft Teams.ahk`, `Shift keys.ahk`). These sounds provide feedback for actions like screenshots, dictation state changes, and task completions.

In an office environment, these sounds can be disruptive. The goal is to implement a global "mute" switch that suppresses these sounds without disabling the functionality of the scripts themselves. Since the scripts appear to run as separate processes (evidenced by `QuickUpdateScripts` launching them individually), a simple in-memory global variable will not suffice for cross-script communication. A file-backed state (INI file) accessed via a helper function in the shared `Utils.ahk` library is the most robust solution.

## Proposed Changes

### 1. State Management (Utils.ahk)
*   **Storage:** Use `data\settings.ini` to store the sound state (`[Settings] SoundEnabled=1`).
*   **Access:** Create `IsSoundEnabled()` function that reads this state.
*   **Modification:** Create `ToggleSoundState()` function that flips the state, updates the INI file, and displays a visual overlay (using `ShowCenteredOverlay_Utils`) to confirm the new state (e.g., "🔇 Sound: OFF" or "🔊 Sound: ON").

### 2. Selection Modal Integration (Utils.ahk)
*   Register the `ToggleSoundState` function as a macro in `InitMacros()`.
*   Assign a mnemonic key (e.g., 's' for Sound) for quick access via `Win+Alt+Shift+U`.

### 3. Sound Guard (All Files)
*   Refactor every instance identified in `sound-audit.md` to check `IsSoundEnabled()` before calling `SoundPlay`, `SoundBeep`, or `DllCall("MessageBeep")`.

## Files to Modify

### Utils.ahk
*   **New Functions:** `IsSoundEnabled()`, `ToggleSoundState()`.
*   **InitMacros:** Add `RegisterMacro(ToggleSoundState, "🔊 Toggle Sound (Mute/Unmute)", "s")`.
*   **Sound Instances:**
    *   `SafePlayPrintScreenSound` (~line 3989)
    *   `SafePlayDictationSound` (~line 4367)
    *   `DictationLoopStop` (~line 811)
    *   `#!+7` (Dictation Paste) (~line 4484)
    *   `#!+j` (Dictation Paste Enter) (~line 4507)

### Shift keys.ahk
*   **Sound Instances:**
    *   `PlayCompletionChime_Gemini` (~line 12527)
    *   `PlayCompletionChime_ChatGPT` (~line 13489)

### Microsoft Teams.ahk
*   **Sound Instances:**
    *   `PlayMicrophoneBeep` (Toggle Mute, Camera, Screen Share)

### Gemini.ahk
*   **Sound Instances:**
    *   `PlayCopyCompletedChime`

### AppLaunchers.ahk
*   **Sound Instances:**
    *   `PomodoroChimeCallback`
    *   Pomodoro Start Sound

## Implementation Strategy
1.  **Core Logic:** Implement the state functions in `Utils.ahk` first, as other scripts depend on it.
2.  **UI Integration:** Add the toggle macro to the Selection Modal to enable testing.
3.  **Refactoring:** Systematically go through each file listed in the audit and wrap sound calls with the conditional check.
4.  **Verification:** Use the toggle macro to switch states and verify that sounds are suppressed/enabled across different scripts.
