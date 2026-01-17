# Sound Audit Report

This document catalogs all instances of sound reproduction (beeps, chimes, audio playback) across all script files in the codebase.

---

## Utils.ahk

### 1. PrintScreen Beep
- **File Path:** `Utils.ahk`
- **Line Reference:** ~3989
- **Trigger:** `!PrintScreen` hotkey (Alt+PrintScreen) - triggers `SafePlayPrintScreenSound()`
- **Sound Type:** Beep
- **Code:** `SoundBeep(800, 200)` - 800 Hz frequency, 200ms duration
- **Function:** `SafePlayPrintScreenSound()` (debounced to prevent duplicates within 1000ms)

### 2. Dictation Start Sound
- **File Path:** `Utils.ahk`
- **Line Reference:** ~4367
- **Trigger:** State detection when dictation window appears via `~#!+0` hotkey (Win+Alt+Shift+0) - calls `SafePlayDictationSound(g_DictationStartSound)`
- **Sound Type:** WAV file
- **Sound File:** `sounds\speach-start.wav`
- **Function:** `SafePlayDictationSound()` (throttled to prevent duplicates within 1000ms)

### 3. Dictation Stop Sound
- **File Path:** `Utils.ahk`
- **Line Reference:** ~4300
- **Trigger:** Timer callback `PlayDictationCompletionChime()` scheduled 2.5 seconds after dictation window closes
- **Sound Type:** WAV file
- **Sound File:** `sounds\speach-finished.wav`
- **Function:** `SafePlayDictationSound(g_DictationStopSound)` within `PlayDictationCompletionChime()`

### 4. Dictation Loop Sound
- **File Path:** `Utils.ahk`
- **Line Reference:** ~811
- **Trigger:** `DictationLoopStop()` function called automatically after 60 seconds of dictation loop
- **Sound Type:** WAV file
- **Sound File:** `sounds\retro1.wav`
- **Code:** `SoundPlay(g_DictationLoopSound)`

### 5. Dictation Paste Signal
- **File Path:** `Utils.ahk`
- **Line Reference:** ~4484
- **Trigger:** `#!+7` hotkey (Win+Alt+Shift+7) - Dictation with paste action
- **Sound Type:** WAV file
- **Sound File:** `sounds\retro3.wav`
- **Code:** `SoundPlay(A_ScriptDir . "\sounds\retro3.wav")`

### 6. Dictation Paste Enter Signal
- **File Path:** `Utils.ahk`
- **Line Reference:** ~4507
- **Trigger:** `#!+j` hotkey (Win+Alt+Shift+J) - Dictation with paste and submit action
- **Sound Type:** WAV file
- **Sound File:** `sounds\retro4.wav`
- **Code:** `SoundPlay(A_ScriptDir . "\sounds\retro4.wav")`

---

## AppLaunchers.ahk

### 7. Pomodoro Chime (Multiple Sounds)
- **File Path:** `AppLaunchers.ahk`
- **Line Reference:** ~1304-1323
- **Trigger:** `PlayCompletionChime(30000)` called when Pomodoro timer completes - triggers `PomodoroChimeCallback()` every 1 second for 30 seconds
- **Sound Type:** Multiple (Beep, System Sound, DLL Call)
- **Function:** `PomodoroChimeCallback()` plays three sounds simultaneously:
  1. **Beep:** `SoundBeep(2000, 300)` - 2000 Hz frequency, 300ms duration
  2. **MessageBeep:** `DllCall("User32\MessageBeep", "UInt", 0xFFFFFFFF)` - Windows system beep
  3. **System Sound:** `SoundPlay("*16")` - System asterisk sound
- **Note:** All three sounds play simultaneously for maximum audibility, repeated every 1 second for 30 seconds

---

## Gemini.ahk

### 8. Copy Completion Chime (Fallback Chain)
- **File Path:** `Gemini.ahk`
- **Line Reference:** ~76-111
- **Trigger:** `PlayCopyCompletedChime()` called after copy button click operations complete (multiple locations at lines ~202, ~359, ~705)
- **Sound Type:** Fallback chain (System Sound, Beep)
- **Function:** `PlayCopyCompletedChime()` with fallback pattern:
  1. **Primary:** `DllCall("User32\\MessageBeep", "UInt", 0xFFFFFFFF)` - Windows MessageBeep
  2. **Fallback:** `SoundPlay("*64", false)` - System asterisk sound (64 notification)
  3. **Last Resort:** `SoundBeep(1100, 130)` - 1100 Hz frequency, 130ms duration
- **Note:** Debounced to prevent duplicate sounds within 1500ms

---

## Microsoft Teams.ahk

### 9. Microphone Beep (Toggle Mute)
- **File Path:** `Microsoft Teams.ahk`
- **Line Reference:** ~436
- **Trigger:** `#!+5` hotkey (Win+Alt+Shift+5) - Toggle Mute action, calls `PlayMicrophoneBeep()` on success
- **Sound Type:** Beep
- **Function:** `PlayMicrophoneBeep()` (line ~277-279)
- **Code:** `SoundBeep(800, 150)` - 800 Hz frequency, 150ms duration

### 10. Microphone Beep (Toggle Camera)
- **File Path:** `Microsoft Teams.ahk`
- **Line Reference:** ~481
- **Trigger:** `#!+4` hotkey (Win+Alt+Shift+4) - Toggle Camera action, calls `PlayMicrophoneBeep()` on success
- **Sound Type:** Beep
- **Function:** `PlayMicrophoneBeep()` (line ~277-279)
- **Code:** `SoundBeep(800, 150)` - 800 Hz frequency, 150ms duration

### 11. Microphone Beep (Toggle Screen Share - First Call)
- **File Path:** `Microsoft Teams.ahk`
- **Line Reference:** ~545
- **Trigger:** `#!+t` hotkey (Win+Alt+Shift+T) - Toggle Screen Share action, first `PlayMicrophoneBeep()` call on success
- **Sound Type:** Beep
- **Function:** `PlayMicrophoneBeep()` (line ~277-279)
- **Code:** `SoundBeep(800, 150)` - 800 Hz frequency, 150ms duration

### 12. Microphone Beep (Toggle Screen Share - Second Call)
- **File Path:** `Microsoft Teams.ahk`
- **Line Reference:** ~549
- **Trigger:** `#!+t` hotkey (Win+Alt+Shift+T) - Toggle Screen Share action, second `PlayMicrophoneBeep()` call in fallback path
- **Sound Type:** Beep
- **Function:** `PlayMicrophoneBeep()` (line ~277-279)
- **Code:** `SoundBeep(800, 150)` - 800 Hz frequency, 150ms duration

---

## Shift keys.ahk

### 13. Gemini Completion Chime (Fallback Chain)
- **File Path:** `Shift keys.ahk`
- **Line Reference:** ~12527-12559
- **Trigger:** `PlayCompletionChime_Gemini()` called after Gemini responses complete (line ~12518)
- **Sound Type:** Fallback chain (System Sound, Beep)
- **Function:** `PlayCompletionChime_Gemini()` with fallback pattern:
  1. **Primary:** `DllCall("User32\\MessageBeep", "UInt", 0xFFFFFFFF)` - Windows MessageBeep
  2. **Fallback:** `SoundPlay("*64", false)` - System asterisk sound (64 notification)
  3. **Last Resort:** `SoundBeep(1100, 130)` - 1100 Hz frequency, 130ms duration
- **Note:** Debounced to prevent duplicate sounds within 1500ms

### 14. ChatGPT Completion Chime (Fallback Chain)
- **File Path:** `Shift keys.ahk`
- **Line Reference:** ~13489-13521
- **Trigger:** `PlayCompletionChime_ChatGPT()` called after ChatGPT responses complete (lines ~8349, ~8375, ~13587)
- **Sound Type:** Fallback chain (System Sound, Beep)
- **Function:** `PlayCompletionChime_ChatGPT()` with fallback pattern:
  1. **Primary:** `DllCall("User32\\MessageBeep", "UInt", 0xFFFFFFFF)` - Windows MessageBeep
  2. **Fallback:** `SoundPlay("*64", false)` - System asterisk sound (64 notification)
  3. **Last Resort:** `SoundBeep(1100, 130)` - 1100 Hz frequency, 130ms duration
- **Note:** Debounced to prevent duplicate sounds within 1500ms

---

## Summary

**Total Sound Instances:** 14 unique instances across 5 files

**Sound Files Referenced:**
- `sounds\speach-start.wav` (Dictation start)
- `sounds\speach-finished.wav` (Dictation stop)
- `sounds\retro1.wav` (Dictation loop)
- `sounds\retro3.wav` (Dictation paste signal)
- `sounds\retro4.wav` (Dictation paste enter signal)

**Sound Types:**
- **Beep sounds:** `SoundBeep()` calls with various frequencies and durations
- **System sounds:** `SoundPlay()` with system sound identifiers (`*16`, `*64`)
- **WAV files:** `SoundPlay()` with file paths to `.wav` files
- **Windows API:** `DllCall("User32\MessageBeep")` for system beeps

**Files with Sound Instances:**
1. `Utils.ahk` - 6 instances
2. `AppLaunchers.ahk` - 1 instance (multiple sounds)
3. `Gemini.ahk` - 1 instance (fallback chain)
4. `Microsoft Teams.ahk` - 4 instances (same function, different triggers)
5. `Shift keys.ahk` - 2 instances (fallback chains)

---

*Generated: Sound Audit Report*