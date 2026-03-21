## Hand Off Warning Cues – Synchronization Audio Reference

This document is the authoritative reference for **Hand Off** synchronization audio cues in the Dictation → Gemini → Cursor (D2C) workflows.

The goal is to:

- Warn you briefly **before** the automation takes over window focus or copies content on your behalf.
- Keep you in control of the original window while Gemini is working.
- Minimize blocking sleeps by applying a **strict vector rule** for when the 2-second cue is allowed.

All cues use the same asset and timing:

- **Sound file**: `pre-movement.wav` (implemented today via `pre-movement.mp3` in `sounds/`).
- **Delay**: exactly **2 seconds** between the cue and the automated action.

---

## Vector-based cue rules

- **Outbound (trigger → target)**  
  When focus or control moves from the **original trigger window** (where you started the flow) to a **target window** for a significant automated action (submit to Gemini, copy last Gemini response), the system:
  - Plays `pre-movement.wav`.
  - Waits **2 seconds**.
  - Performs the automated action.

- **Exception — TTS from selection (`Win+Alt+Shift+7`, `GeminiAsyncTTS` in `Gemini.ahk`)**  
  The **first** Original → Gemini transition (activate, paste “repeat exactly” prompt, submit, then return focus to the original window) does **not** play the Hand Off cue.  
  After Gemini has **started** responding (streaming observed) and **finished** (Stop streaming button gone), the **second** Original → Gemini transition (activate trash tab and run read aloud) **does** play the cue **immediately before** that activation—same asset and **2-second** delay as other outbound moves.

- **Inbound (target → trigger)**  
  Any transition that **returns** focus to the original trigger window happens **immediately**:
  - **No** sound.
  - **No** additional delay.

- **Secondary-to-secondary transitions**  
  Moves between two non-original windows (for example, **Gemini → Cursor** during transfer) do **not** play the Hand Off cue.

- **Background processing**  
  While Gemini is generating a response, the original window stays in user control.  
  For most D2C flows, no cue plays until **after the response is ready** and **immediately before** the outbound automated action (submit, copy, etc.).  
  For **TTS from selection (`#!+7`)**, the **initial** submit to Gemini also runs **without** a cue; the cue applies only **before** the later outbound move to Gemini for **read aloud**, once the response has completed (see integration table).

---

## Integration table

| File Reference | Trigger Event | Next Action | Timing Logic |
| -------------- | ------------- | ----------- | ------------ |
| `Utils.ahk` (`D2C_FlowManager.ExecuteGeminiSubmit`) | **Send to Gemini?** banner: user presses **Y** or 6s timeout selects **Y** by default while original trigger window is active | Activate Gemini tab, paste dictated text into the prompt field, optionally send **Enter** to submit | **Outbound**: play `pre-movement.wav`, wait 2s, then focus/paste into Gemini |
| `Utils.ahk` (`D2C_FlowManager.DoCopyCore`) | **Copy response?** banner: user presses **Y** (copy) or **R** (copy+read), or presses **C** (copy+transfer) in the manager-driven D2C flow | Ensure Gemini is active, copy the last Gemini response to the clipboard; optionally dispatch read-aloud or Cursor transfer | **Outbound** (original → Gemini for copy path): play `pre-movement.wav`, wait 2s, then copy |
| `Utils.ahk` (`D2C_FlowManager.DoCopyCore` via `DoCopyOnTimeout`) | **Copy response?** banner times out in the manager-driven D2C flow; `DoCopyOnTimeout` path runs | Ensure Gemini is active, copy the last Gemini response to the clipboard, then optionally restore focus to the original window | **Outbound**: play `pre-movement.wav`, wait 2s, then copy |
| `Gemini.ahk` (`GeminiDelayedSubmitMonitor.DoCopyCore`) | Legacy **Copy response?** banner: user presses **Y** or **R** when the delayed-submit monitor shows **Copy?** after Ctrl+Alt+Win+L | Activate Gemini if needed, copy last Gemini response, optionally trigger read aloud and/or restore focus | **Outbound**: play `pre-movement.wav`, wait 2s, then copy |
| `Gemini.ahk` (`GeminiDelayedSubmitMonitor.DoCopyCore` via `DoCopyOnTimeout`) | Legacy monitor: **Copy response?** banner times out and `DoCopyOnTimeout` calls `DoCopyCore(false)` | Activate Gemini if needed, copy last Gemini response, then restore focus to the original window | **Outbound**: play `pre-movement.wav`, wait 2s, then copy |
| `Utils.ahk` / `Gemini.ahk` (various return paths after copy) | Flow finished using Gemini’s response; script is returning focus to the original trigger window | Activate original window only | **Inbound**: **no** cue, immediate activation (no delay) |
| `Utils.ahk` (`D2C_FlowManager.PromptForCursorTransfer`) and `Gemini.ahk` (`GeminiDelayedSubmitMonitor.CopyAndTransferToCursor`) | User presses **C** at **Copy response?**, then selects a Cursor window 1–9 | Show window picker, then activate Cursor and paste+Enter into its AI field | **Secondary → secondary**: **no** Hand Off cue; any delays are limited to normal focus and paste waits |
| `Gemini.ahk` (`GeminiAsyncTTS.Start`) | **Win+Alt+Shift+7**: user triggers TTS from selection | First Original → Gemini: paste prompt, submit, restore focus to original | **No** Hand Off cue on this transition |
| `Gemini.ahk` (`GeminiAsyncTTS.CheckCompletion` → `GeminiTriggerReadAloud`) | Same flow: streaming has been observed and has finished | Second Original → Gemini: activate trash tab, open Listen / read aloud | **Outbound**: play `pre-movement.wav`, wait 2s, then focus Gemini and read aloud |

This table can be extended as new synchronization cues are added to other flows.

---

## Implementation guidance for AIB

When adding or modifying synchronization audio cues in this repository:

- **Honor the vector rule**  
  - Only **outbound** transitions from the original trigger window to a target window may play the 2-second Hand Off cue.  
  - Do **not** add cues to inbound (return) transitions or secondary-to-secondary transitions.
  - **TTS from selection (`#!+7`)**: do **not** play the cue on the **first** outbound Original → Gemini (submit); **do** play it **immediately before** the **second** outbound Original → Gemini for read aloud **after** the response has finished (streaming stopped).

- **Reuse the same contract**  
  - Always use `pre-movement.wav` (current implementation: `pre-movement.mp3`) and a **2-second delay**.  
  - Do not introduce additional fixed sleeps around the cue; rely on existing condition-based waits (`WinWaitActive`, UIA checks, clipboard validation) to maintain synchronization.

- **Preserve background control**  
  - Keep the original window active while Gemini is processing.  
  - Only fire the cue when the system is about to take focus or perform an automated copy/paste, and only **after** the response or state is ready.

- **Follow the efficiency canon**  
  - Avoid unbounded waits and extra polling loops.  
  - Keep all waits timeout-bounded and condition-driven.  
  - Do not use synthetic keystrokes to unknown windows; always validate `hwnd` and active window state before sending input.

- **Cursor transfer (paste + Enter)**  
  - Restore focus to the anchored window via `CursorTransfer_ActivateFocusPaste(targetHwnd, restoreFocusHwnd)` after Enter; **no** Hand Off cue on that return. Details: `docs/dictation-to-gemini-cursor-flow.md` (Instructions for AIB).

For a narrative description of how these cues fit into the D2C flow, see `docs/dictation-to-gemini-cursor-flow.md`.

