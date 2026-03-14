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
    Branch --> |"PasteEnter"| PasteEnter["^v Enter, hide indicator"]
    Branch --> |"pendingGemini"| ShowAndWait["Send to Gemini? 6s"]
    ShowAndWait --> User6s{"Y / N / timeout"}
    User6s --> |"N"| Stop6s["Stop, no submit"]
    User6s --> |"timeout"| Stop6s
    User6s --> |"Y"| DelayedFlow["DelayedSubmitFlow"]
    DelayedFlow --> Banner4s["Submitting 4s, 3s timer"]
    Banner4s --> User4s{"Y / N / 3s"}
    User4s --> |"N"| Stop4s["Stop, paste only"]
    User4s --> |"Y or 3s"| Finalize["FinalizeSubmit"]
    Finalize --> PasteGemini["Focus Gemini, paste"]
    PasteGemini --> WaitContent["Wait content max 5s"]
    WaitContent --> SendEnter["Send Enter"]
    SendEnter --> StartMonitor["MonitorStart"]
    StartMonitor --> ResponseDone["Response done"]
    ResponseDone --> BannerCopy["Copy response? 5s"]
    BannerCopy --> UserCopy{"Y / N / R / C / E / timeout"}
    UserCopy --> |"N"| StopCopy["Stop, no copy"]
    UserCopy --> |"E"| StopCopy
    UserCopy --> |"Y"| DoCopyOnly["DoCopyCore"]
    UserCopy --> |"R"| DoCopyRead["DoCopyCore + read"]
    UserCopy --> |"C"| DoTransfer["Copy + CursorTransfer"]
    UserCopy --> |"timeout"| DoCopyTimeout["DoCopyOnTimeout"]
    DoTransfer --> SelectWin["Pick window 1–9, Esc=cancel"]
    SelectWin --> ActivatePaste["ActivateFocusPaste ^v Enter"]
```

On the 6s banner, only **N** shows the “Gemini submission cancelled” overlay; **timeout** ends the flow without that overlay.

## Where the user can stop the flow

| Step | Banner / moment                        | User action to stop             | Effect                                                 |
| ---- | -------------------------------------- | ------------------------------- | ------------------------------------------------------ |
| 1    | **Send transcription to Gemini?** (6s) | Press **N** or no key (timeout) | Flow ends; no paste to Gemini, no Enter, no 4s banner. |
| 2    | **Submitting in 4s...** (4s)           | Press **N**                     | No paste to Gemini, no Enter; no monitor started.      |
| 3    | **Copy response?** (5s)                | Press **N** or **E**            | No copy, no read aloud, no transfer to Cursor.         |
| 4    | **Transfer to Cursor** (window picker) | Press **Esc**                   | Transfer cancelled; no paste/Enter to Cursor.          |

## Hotkeys

| Hotkey                  | Role                                                                                                                           |
| ----------------------- | ------------------------------------------------------------------------------------------------------------------------------ |
| `~#!+0`                 | Start dictation (first press); stop dictation (second press). If stopped manually (not via #!+j), sets “Send to Gemini?” path. |
| `#!+j`                  | Programmatic stop: set PasteEnter, send ~#!+0. Transcription is pasted and submitted in **current app** (no Gemini banner).    |
| `Ctrl+Alt+Win+L`        | Direct delayed-submit flow (4s banner, paste + optional Enter to Gemini).                                                      |
| `#!+U` then **L** twice | Same flow from hotstring selector (double-tap L in hotstring context).                                                         |

## Files and entry points

- **Utils.ahk**: `~#!+0`, `#!+j`, `DictationGeminiConfirm_*`, `GeminiDelayedSubmitFlow`, `GeminiFinalizeSubmit`, `GeminiCancelAutoSubmit`, `CursorTransfer_ShowWindowSelector`, `CursorTransfer_ActivateFocusPaste`.
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
