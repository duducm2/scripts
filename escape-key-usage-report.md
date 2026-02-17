## Escape Key Usage Report

**Scope**: Deliberate Escape key sends in the core scripts:
`WindowManagement.ahk`, `Utils.ahk`, `Spotify.ahk`, `Shift keys.ahk`,
`Gemini.ahk`, and `AppLaunchers.ahk`.

**Patterns included**: `Send "{Esc}"`, `Send "{Escape}"`,
`Send "{Escape 2}"`, `SendInput "{Escape}"`, and equivalents in function-call
syntax such as `Send("{Esc}")`.

> Note: Other scripts in the repo (e.g. `Outlook.ahk`) also send Escape but
> are out of scope for this report, which focuses on the six core files above.

---

### 1. `Shift keys.ahk`

#### Summary

`Shift keys.ahk` uses Escape heavily to:

- Close context menus and reaction pickers.
- Reset focus or clear selections before opening search/command UIs.
- Fallback-cancel flows when UI automation fails (e.g. Power BI, file dialogs).
- Clear Gemini/Cursor UI state before injecting text or switching AI modes/models.

#### Usages

| Line   | Send expression      | Context / purpose |
|--------|----------------------|-------------------|
| 4024   | `Send "{Esc}"`       | After navigating a context menu (AppsKey, arrows, Enter), Escape is sent to close the menu/UI before continuing. |
| 4398   | `Send "{Esc}"`       | Shift+L “Like” reaction: Enter → Enter applies reaction, then Esc closes the reaction picker / UI. |
| 4407   | `Send "{Esc}"`       | Shift+G “Heart” reaction: Enter/Down/Enter selects heart, then Esc closes the picker. |
| 4417   | `Send "{Esc}"`       | Shift+J “Laugh” reaction: Enter/Down/Down/Enter selects laugh, then Esc closes the picker. |
| 6466   | `Send "{Esc}"`       | ChatGPT helper: Esc is sent before pasting the rules prompt to ensure the composer is focused and any popups/overlays are dismissed. |
| 7042   | `Send "{Esc}"`       | Power BI Shift+U “Close and apply”: first Esc to exit current edit/field. |
| 7043   | `Send "{Esc}"`       | Power BI Shift+U “Close and apply”: second Esc as extra safety to fully exit dialogs before Alt-based ribbon commands. |
| 7487   | `Send "{Esc}"`       | Power BI fallback when cancel button cannot be found via UIA: Esc used as universal cancel for context menus/dialogs. |
| 8527   | `Send "{Escape}"`    | Ctrl+E helper: after triggering Editor UI, double Esc is used to close extra overlays/search/UI before focusing on code. |
| 8528   | `Send "{Escape}"`    | Second Esc in the same Ctrl+E flow to ensure the UI is fully cleared. |
| 8536   | `Send "{Escape}"`    | Ctrl+1 “Remove clustering and focus code”: comment “Send ESC two times” — first Esc to clear any clustering/overlay. |
| 8538   | `Send "{Escape}"`    | Second Esc in Ctrl+1 sequence to fully reset Cursor UI before layout hotkeys run. |
| 9099   | `Send "{Escape 2}"`  | `ExecuteAIModelSelection`: Esc twice to clear any active menus/modals before locating the Agent/Ask edit field. |
| 9118   | `Send "{Escape}"`    | `ExecuteAIModelSelection` case 1: after using Ctrl+; and Enter to set model, Esc closes the model selection menu. |
| 10029  | `Send "{Escape 2}"`  | `SwitchAIMode`: double Esc to clear UI before focusing the edit field and opening the Ctrl+. mode picker. |
| 10064  | `Send "{Escape 2}"`  | `SwitchAIModel`: double Esc to clear UI before focusing the edit field for model selection. |
| 10083  | `Send "{Escape}"`    | `SwitchAIModel` case 1 (auto): after Ctrl+; and Enter to set model, Esc closes the model selection menu. |
| 11617  | `Send "{Esc}"`       | Tree-item search helper: Esc clears any current selection/focus before opening Ctrl+F search in the UIA Tree Inspector window. |
| 11629  | `Send "{Esc}"`       | Tree-item search helper: Esc closes the search box after pasting and applying the search. |
| 12216  | `Send "{Escape}"`    | Gemini main menu toggle: fallback path when UIA cannot attach; behaves like previous version by sending Esc to toggle/close Gemini UI. |
| 12249  | `Send "{Escape}"`    | Gemini main menu toggle: if the main menu button cannot be found, fallback is to send Esc to mimic prior behavior. |
| 12253  | `Send "{Escape}"`    | Gemini main menu toggle: if any UIA error occurs, Esc is sent as a graceful fallback to close/toggle the drawer. |
| 13909  | `Send "{Esc}"`       | File dialog helper: universal cancel fallback when no OK/Cancel/Close button can be found via UIA. |

---

### 2. `Utils.ahk`

#### Summary

`Utils.ahk` uses Escape primarily for:

- Cleaning up Hunt-and-Peck overlays/processes.
- Forwarding Escape to the system when dictation is not active (while blocking it during dictation).

#### Usages

| Line   | Send expression                | Context / purpose |
|--------|--------------------------------|-------------------|
| 4983   | `Send "{Esc}"`                 | Hunt-and-Peck cleanup: after safely activating the target window, Esc dismisses any lingering Hunt-and-Peck overlay. |
| 4986   | `SetTimer(() => Send("{Esc}")` | Safety timer: sends another Esc 1s later to dismiss any late/slow overlay that appears after initial cleanup. |
| 5035   | `Send "{Esc}"`                 | Loop-mode Hunt-and-Peck cleanup: Esc clears any residual overlay when the loop is stopped. |
| 5036   | `SetTimer(() => Send("{Esc}")` | Same as above for loop mode: delayed Esc to clear any late overlay. |
| 7079   | `SendInput "{Escape}"`         | Dictation Escape handler: when dictation is not active, the script forwards the user’s Escape keypress to the system via `SendInput` for reliability. (When dictation *is* active, Escape is swallowed.) |

---

### 3. `AppLaunchers.ahk`

#### Summary

`AppLaunchers.ahk` has a single deliberate Escape send used to close a UI after a sequence of keyboard-driven actions.

#### Usages

| Line   | Send expression   | Context / purpose |
|--------|-------------------|-------------------|
| 2036   | `Send("{Esc}")`   | After copying (`^c`) and running Alt-based commands (`!v`, `!q`), Esc is sent to close the invoked menu/dialog and return the UI to a neutral state. |

---

### 4. `WindowManagement.ahk`

No deliberate Escape-key send commands (`Send "{Esc}"`, `Send "{Escape}"`, `SendInput "{Escape}"`, or `Send("{Esc}")`) were found in this file.

---

### 5. `Spotify.ahk`

No deliberate Escape-key send commands (`Send "{Esc}"`, `Send "{Escape}"`, `SendInput "{Escape}"`, or `Send("{Esc}")`) were found in this file.

---

### 6. `Gemini.ahk`

No deliberate Escape-key send commands (`Send "{Esc}"`, `Send "{Escape}"`, `SendInput "{Escape}"`, or `Send("{Esc}")`) were found in this file.

---

### 7. Summary & Observations

- **Context/menu cleanup**: Many Escape sends are used as a generic way to close context menus, dialogs, or overlays after automated UI flows (Power BI, reactions, Hunt-and-Peck, Gemini UI toggles).
- **State reset before actions**: Double-Escape patterns (e.g. `"{Escape 2}"`) are used before focusing text fields or invoking Cursor/Gemini commands to guarantee a clean state.
- **Fallback safety**: In several places, Esc is the last-resort cancel when UI automation fails (e.g. cannot find a button or dialog control).
- **Input routing**: The dictation handler in `Utils.ahk` deliberately blocks or forwards Escape depending on dictation state, ensuring external tools are not accidentally closed.

