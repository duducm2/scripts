# Dictation → Gemini → Cursor Flow

Flow starting from **Win+Alt+Shift+0** (`~#!+0`) to dictate, optionally send transcription to Gemini, then optionally copy/transfer the response to Cursor. The user can **stop the flow at any moment** at the highlighted cancel points.

## High-level flow

```mermaid
flowchart TB
    H0["~#!+0 first press"]
    H0 --> StartDict["Start dictation, indicator on"]
    StartDict --> H0Stop["~#!+0 second press"]
    H0Stop --> SetPending["Set pending Gemini if active"]
    SetPending --> ToggleMode["ToggleDictationMode, clipboard fills"]
    ToggleMode --> WaitClip["ChimeOrWait clipboard 1.5s"]
    WaitClip --> PlayChime["PlayDictationCompletionChime"]
    PlayChime --> Branch["Branch by action"]
    Branch --> |"Paste"| PasteOnly["^v, hide indicator"]
    Branch --> |"pendingGemini"| ShowAndWait["Send to Gemini? 6s"]
    ShowAndWait --> User6s{"Y / S / N / timeout"}
    User6s --> |"N"| Stop6s["Stop, no submit"]
    User6s --> |"timeout"| DelayedFlow
    User6s --> |"Y"| DelayedFlow["DelayedSubmitFlow"]
    User6s --> |"S"| PasteOnly6s["Paste only, no Enter"]
    DelayedFlow --> Banner4s["Submitting 4s, 3s timer"]
    Banner4s --> User4s{"Y / N / 3s"}
    User4s --> |"N"| Stop4s["Paste only, 4s banner"]
    User4s --> |"Y or 3s"| Finalize["FinalizeSubmit"]
    Finalize --> PasteGemini["Focus Gemini, paste"]
    PasteGemini --> WaitContent["Wait content max 5s"]
    WaitContent --> SendEnter["Send Enter"]
    SendEnter --> StartMonitor["MonitorStart"]
    StartMonitor --> ResponseDone["Response done"]
    ResponseDone --> BannerCopy["Copy response? 5s"]
    BannerCopy --> UserCopy{"Y / N / R / C / timeout"}
    UserCopy --> |"N"| StopCopy["Stop, no copy"]
    UserCopy --> |"Y"| DoCopyOnly["DoCopyCore"]
    UserCopy --> |"R"| DoCopyRead["DoCopyCore + read"]
    UserCopy --> |"C"| DoTransfer["Copy + CursorTransfer"]
    UserCopy --> |"timeout"| DoCopyTimeout["DoCopyOnTimeout"]
    DoTransfer --> SelectWin["Pick window 1–9, N or Esc=cancel"]
    SelectWin --> ActivatePaste["ActivateFocusPaste ^v Enter"]
```

On the 6s banner (step 1), **Y** = send (then 4s countdown, then paste + Enter; you can press N during 4s for paste-only). **S** = paste to Gemini only (no Enter, no 4s). **N** = cancel and show “Gemini submission cancelled”. **timeout** = same as Y. The 4s countdown is part of the send path, not a separate step.

## Where the user can stop the flow

| Step | Banner / moment                        | User action to stop                                                                                                                  | Effect                                                                                          |
| ---- | -------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------ | ----------------------------------------------------------------------------------------------- |
| 1    | **Send transcription to Gemini?** (6s) | **N**: cancel. **S**: paste only (no Enter). **Y** or **timeout**: send (4s countdown then paste + Enter; N during 4s = paste-only). | Flow ends at N; S = paste only; Y/timeout = proceed to paste + Enter (4s is part of this path). |
| 2    | **Copy response?** (5s)                | Press **N**                                                                                                                          | No copy, no read aloud, no transfer to Cursor.                                                  |
| 3    | **Transfer to Cursor** (window picker) | Press **N** or **Esc**                                                                                                               | Transfer cancelled; no paste/Enter to Cursor.                                                   |

## Hotkeys

| Hotkey                  | Role                                                                                                            |
| ----------------------- | --------------------------------------------------------------------------------------------------------------- |
| `~#!+0`                 | Start dictation (first press); stop dictation (second press). If stopped manually, sets “Send to Gemini?” path. |
| `Ctrl+Alt+Win+L`        | Direct delayed-submit flow (4s banner, paste + optional Enter to Gemini).                                       |
| `#!+U` then **L** twice | Same flow from hotstring selector (double-tap L in hotstring context).                                          |

## Files and entry points

- **Utils.ahk**: `~#!+0`, `DictationGeminiConfirm_*`, `GeminiDelayedSubmitFlow`, `GeminiFinalizeSubmit`, `GeminiCancelAutoSubmit`, `CursorTransfer_ShowWindowSelector`, `CursorTransfer_ActivateFocusPaste`.
- **Gemini.ahk**: `GeminiDelayedSubmitMonitor` (completion detection, “Copy response?” banner, Copy/Read/Transfer actions).

## Simplified “happy path” (no cancel)

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
    AHK->>User: Submitting 4s
    User->>AHK: Y or 3s
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
