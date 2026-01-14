---
name: Dictation Loop Macro
overview: Implement a togglable macro in Utils.ahk that automates a 40-second dictation loop for the "handy" application to prevent transcription timeouts.
todos:
  - id: define_globals
    content: Define global variables for dictation loop state and timers in Utils.ahk.
    status: pending
    dependencies: []
  - id: implement_loop_logic
    content: Create the ToggleDictationLoop function and helper timer functions to manage the Start-Wait-Stop-Repeat cycle.
    status: pending
    dependencies: [define_globals]
  - id: register_macro
    content: Register the new macro in InitMacros() to make it accessible via the Win+Alt+Shift+U modal.
    status: pending
    dependencies: [implement_loop_logic]
---

# Dictation Loop Macro

## Analysis / Context
The `Utils.ahk` script currently manages system-wide utilities and macros, including a modal selector (`Win+Alt+Shift+U`) and a dictation toggle (`Win+Alt+Shift+0`). The "handy" transcription software has a limitation where long dictations may fail or time out. To mitigate this, the user requires a mechanism to automatically chunk dictation into 40-second intervals.

The goal is to create a macro that, when activated, triggers the existing dictation toggle (`#!+0`), waits 40 seconds, triggers it again to stop (processing the text), and then repeats the cycle indefinitely until disabled.

## Proposed Changes

### 1. Global Variables
Define `g_DictationLoopActive` to track the loop state.

### 2. Loop Logic Functions
Implement a non-blocking loop using `SetTimer` to ensure the UI remains responsive.
*   **`ToggleDictationLoop()`**: Entry point. Toggles the state. If enabling, starts the cycle. If disabling, stops timers and resets state.
*   **`DictationLoopStart()`**: Sends `#!+0` to start dictation and schedules the stop event after 40 seconds.
*   **`DictationLoopStop()`**: Sends `#!+0` to stop dictation (triggering transcription) and schedules the next start event after a short buffer (e.g., 2 seconds) to allow for processing.

### 3. Modal Integration
Add `RegisterMacro(ToggleDictationLoop, "🎙️ Dictation Loop (40s)")` to the `InitMacros()` function. The existing logic in `Utils.ahk` will automatically assign the next available character key to this macro.

## Files to Modify

*   `c:\Users\fie7ca\Documents\scripts\Utils.ahk`

## Implementation Strategy

1.  **Define Globals**:
    Add `global g_DictationLoopActive := false` near the top of the Macros System section or with other globals.

2.  **Implement Logic**:
    Add the following functions in the Macros System section:
    *   `ToggleDictationLoop()`:
        *   Check `g_DictationLoopActive`.
        *   If true: Set to false, turn off timers (`DictationLoopStop`, `DictationLoopStart`), show overlay "Dictation Loop Stopped".
        *   If false: Set to true, call `DictationLoopStart()`, show overlay "Dictation Loop Started".
    *   `DictationLoopStart()`:
        *   Check `g_DictationLoopActive` (safety).
        *   Send `#!+0`.
        *   SetTimer `DictationLoopStop`, 40000 (40s).
    *   `DictationLoopStop()`:
        *   Check `g_DictationLoopActive` (safety).
        *   Send `#!+0`.
        *   SetTimer `DictationLoopStart`, 2000 (2s buffer).

3.  **Register Macro**:
    In `InitMacros()`, add:
    ```ahk
    RegisterMacro(ToggleDictationLoop, "🎙️ Dictation Loop (40s)")
    ```
