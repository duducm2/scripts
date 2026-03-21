# Dictation → Gemini → Cursor Flow

1. **Start dictation** with **Win+Alt+Shift+0** (`~#!+0`).
2. **Finish dictating** (second press of `~#!+0` or stop manually).
3. **First banner — Send to Gemini?**  
   **Y** = paste and auto-send (Enter) to Gemini. **S** = paste only (no Enter). **N** = cancel (flow ends).  
   If no action is taken within 6 seconds, **Y** (yes) is selected by default. When the script moves focus from the original window to Gemini to perform this paste, it first shows a **2-second “✋ Hands off!” pre-movement cue** so you can stop typing before the automated transition.
4. **If you chose Y**, after Gemini responds you see **Copy response?**  
   **Y** = copy. **C** = send to Cursor. **R** = same copy as **Y**, then read aloud (Listen); both steps use synchronous IPC so focus is not restored until the read-aloud flow finishes. **N** = cancel (flow ends).  
   If you press **Y**, a 2-second **“✋ Hands off!”** cue plays before the script copies the last Gemini response.  
   If no action is taken within 6 seconds, **N** (no) is selected by default and `DoCopyOnTimeout` will still copy Gemini's response after playing the same 2-second **“✋ Hands off!”** cue, because the script is about to take control of focus to complete the copy.

Pressing **N** at any banner terminates the whole flow.

## High-level flow

```mermaid
flowchart TB
    H0["~#!+0 first press"]
    H0 --> StartDict["Start dictation"]
    StartDict --> H0Stop["~#!+0 second press"]
    H0Stop --> SetPending["Set pending Gemini if active"]
    SetPending --> ToggleMode["ToggleDictationMode, clipboard fills"]
    ToggleMode --> WaitClip["ChimeOrWait clipboard"]
    WaitClip --> PlayChime["PlayDictationCompletionChime"]
    PlayChime --> Branch["Branch by action"]
    Branch --> |"Paste"| PasteOnly["^v, hide indicator"]
    Branch --> |"pendingGemini"| ShowAndWait["Send to Gemini? 6s"]
    ShowAndWait --> User6s{"Y / S / N / timeout"}
    User6s --> |"N"| Stop6s["Flow ends"]
    User6s --> |"timeout"| Finalize["FinalizeSubmit"]
    User6s --> |"Y"| Finalize
    User6s --> |"S"| PasteOnly6s["Paste only, no Enter"]
    Finalize --> PasteGemini["Focus Gemini, paste"]
    PasteGemini --> WaitContent["Wait content max 5s"]
    WaitContent --> SendEnter["Send Enter"]
    SendEnter --> StartMonitor["MonitorStart"]
    StartMonitor --> ResponseDone["Response done"]
    ResponseDone --> BannerCopy["Copy response? 5s"]
    BannerCopy --> UserCopy{"Y / N / R / C / timeout"}
    UserCopy --> |"N"| StopCopy["Flow ends"]
    UserCopy --> |"Y"| DoCopyOnly["DoCopyCore"]
    UserCopy --> |"R"| DoCopyRead["DoCopyCore + read"]
    UserCopy --> |"C"| DoTransfer["Copy + CursorTransfer"]
    UserCopy --> |"timeout"| DoCopyTimeout["DoCopyOnTimeout"]

    DoTransfer --> SelectWin["Pick window 1–9, N or Esc=cancel"]
    SelectWin --> ActivatePaste["ActivateFocusPaste ^v Enter"]
```

## Pre-movement cue behavior

- **When the cue plays**:
  - A 2-second pre-movement cue (sound + centered overlay \"✋ Hands off!\" warning) plays when the flow first moves focus from the original trigger window to Gemini (Original → Gemini) to submit the prompt.
  - The same 2-second cue also plays right before copying Gemini's last response, whether you press **Y** at **Copy response?** or the banner times out and `DoCopyOnTimeout` runs.
- **When the cue does not play**:
  - **Returns** from Gemini or Cursor back to the original trigger window are **immediate** (no sound cue, no added delay).
  - Transitions between non-original windows (e.g., Gemini → Cursor during transfer) also run **without** the pre-movement cue.

These cues are synchronization guard rails that appear only when the script is about to take control of focus for a significant action (sending to Gemini or copying Gemini's response); they do **not** change which windows are ultimately activated or the order in which they are activated.

For a complete list of where Hand Off audio cues are used, see `docs/hand_off_warning_cues.md` (including **Win+Alt+Shift+7** TTS from selection, which skips the cue on the first move to Gemini for submit but plays it before the second move for read aloud after the response completes).

## Where the user can stop the flow

| Step | Banner                                 | Actions                                                                       | Effect                                                                              |
| ---- | -------------------------------------- | ----------------------------------------------------------------------------- | ----------------------------------------------------------------------------------- |
| 1    | **Send to Gemini?** (6s)               | **Y** or timeout = send. **S** = paste only. **N** = cancel.                  | N ends flow; S = paste only; Y/timeout = paste + Enter, then Copy response? banner. |
| 2    | **Copy response?** (5s)                | **Y** = copy. **C** = transfer to Cursor. **R** = read aloud. **N** = cancel. | N ends flow. Y/C/R/timeout perform their action.                                    |
| 3    | **Transfer to Cursor** (window picker) | **N** or **Esc** = cancel. **1–9** = paste to that window.                    | Cancel = no paste to Cursor.                                                        |

## Hotkeys

| Hotkey                  | Role                                                                                                            |
| ----------------------- | --------------------------------------------------------------------------------------------------------------- |
| `~#!+0`                 | Start dictation (first press); stop dictation (second press). If stopped manually, sets “Send to Gemini?” path. |
| `Ctrl+Alt+Win+L`        | Direct paste+Enter to Gemini (no dictation).                                                                    |
| `#!+U` then **L** twice | Same from hotstring selector.                                                                                   |

## Files and entry points

- **Utils.ahk**: `~#!+0`, `DictationGeminiConfirm_*`, `GeminiDelayedSubmitFlow`, `GeminiFinalizeSubmit`, `CursorTransfer_ShowWindowSelector`, `CursorTransfer_ActivateFocusPaste`.
- **Gemini.ahk**: `GeminiDelayedSubmitMonitor` (response done detection), “Copy response?” banner, Copy/Read/Transfer actions.

## Happy path (no cancel)

```mermaid
sequenceDiagram
    participant User
    participant AHK as Utils.ahk
    participant Gemini as Gemini.ahk

    User->>AHK: ~#!+0 start
    AHK->>User: Indicator on
    User->>AHK: ~#!+0 stop
    AHK->>AHK: Chime
    AHK->>User: Send to Gemini? 6s
    User->>AHK: Y
    AHK->>AHK: Paste, Enter
    AHK->>Gemini: Monitor start
    Gemini->>Gemini: Poll until done
    Gemini->>User: Copy response? 5s
    User->>Gemini: C Transfer
    Gemini->>AHK: Copy + CursorTransfer
    AHK->>User: Pick window 1–9
    User->>AHK: 1–9
    AHK->>AHK: ^v Enter
```
