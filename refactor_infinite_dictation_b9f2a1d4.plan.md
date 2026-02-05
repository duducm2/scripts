---
name: Refactor Infinite Dictation Module
overview: Architect a robust, isolated AutoHotkey module for "Infinite Dictation" that manages Handy.exe lifecycles and ClipAngel merging with strict state management to prevent race conditions.
todos:
  - id: create_module_file
    content: Create Lib/InfiniteDictation.ahk
    status: pending
    dependencies: []
  - id: define_class_structure
    content: Define InfiniteDictation class with static state properties (IsActive, LoopState)
    status: pending
    dependencies: [create_module_file]
  - id: implement_start_method
    content: Implement Start() method (Clear clipboard, Sound, Banner, Set Active)
    status: pending
    dependencies: [define_class_structure]
  - id: implement_loop_cycle
    content: Implement LoopCycle() method (Start Handy, Schedule Stop)
    status: pending
    dependencies: [implement_start_method]
  - id: implement_handy_toggle
    content: Implement helper to trigger Win+Alt+Shift+0
    status: pending
    dependencies: [define_class_structure]
  - id: implement_stop_handy_logic
    content: Implement StopHandyAndRestart() (Trigger Stop, Force Close if needed, Monitor Termination)
    status: pending
    dependencies: [implement_loop_cycle]
  - id: implement_termination_monitor
    content: Implement WaitForHandyExit() with high-frequency polling
    status: pending
    dependencies: [implement_stop_handy_logic]
  - id: implement_stop_method
    content: Implement Stop() method (Reset state, Cancel timers, Trigger Merge)
    status: pending
    dependencies: [define_class_structure]
  - id: implement_merge_logic
    content: Implement MergeClips() using ClipAngel integration
    status: pending
    dependencies: [implement_stop_method]
  - id: update_utils_hotkeys
    content: Update Utils.ahk to import module and map #!+7 and #!+0 to new class methods
    status: pending
    dependencies: [implement_start_method, implement_stop_method]
---

# Refactor Infinite Dictation Module

## Analysis / Context

The current "Infinite Dictation" implementation in `Utils.ahk` relies on scattered global variables (`g_DictationLoopActive`, `g_DictationCheckTimer`) and loosely coupled timers. This architecture is prone to race conditions, specifically where `Handy.exe` might not be fully terminated before the next loop begins, or where state flags become desynchronized from the actual process state. A clean, isolated module is required to enforce strict state management and ensure the "restart every 60s" requirement is met reliably.

## Proposed Changes

Create a new `Lib/InfiniteDictation.ahk` file containing a static class `InfiniteDictation`. This class will encapsulate all state (`IsActive`, `Timer references`) and logic. The `Utils.ahk` file will be simplified to delegate hotkey actions to this new module.

## Files to Modify

- `Lib/InfiniteDictation.ahk` (New)
- `Utils.ahk` (Modify hotkeys `#!+7` and `#!+0`)

## Implementation Strategy

1.  **State Management:** Use a static class `InfiniteDictation` to hold `IsActive` flag and `LoopTimer` references.
2.  **Initialization:** `Start()` method will handle clipboard clearing (via `ClipAngel` or internal logic) and UI feedback.
3.  **The Loop:**
    - `Cycle()` method triggers `#!+0` to start Handy.
    - Uses `SetTimer` (negative) to schedule the stop action after 15s.
4.  **Critical Wait:**
    - The stop action triggers `#!+0` and then enters a "Monitoring" state.
    - A high-frequency timer (e.g., 50ms) checks `ProcessExist("handy.exe")`.
    - _Self-Correction:_ If `#!+0` does not close Handy, the monitor will force `ProcessClose` after a short timeout to ensure the "restart" requirement is met.
    - Once confirmed gone, it triggers the next `Cycle()`.
5.  **Termination:**
    - `Stop()` method cancels all pending timers immediately.
    - Calls `MergeClips()` which automates `ClipAngel`.
    - Resets state to allow clean restart.
