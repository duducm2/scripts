## Scripts Toolkit (AutoHotkey v2)

### Philosophy

- Keep hotkeys predictable and memorable
  - Primary set: Shift+[Y U I O P H J K L N M , . W E R T D F G C V B]
  - When the Shift set is full, use Ctrl+Alt+Shift (MEH) in the same order
- Prefer resilient UI Automation (UIA) over pixel/image matching
- Fail safe: if automation fails, fall back to a native key/system action
- Avoid blocking popups; prefer overlays/banners and silent fallbacks

### Folder Inventory

- `Shift keys.ahk` — Global app shortcuts, UIA patterns, Spotify controls, Outlook helpers
- `Spotify.ahk` — App launcher and media helpers (volume control with state-aware focus/return)
- `Microsoft Teams.ahk` — Meeting/chat helpers, robust window activation, mic/camera state verification
- `WindowManagement.ahk` — Window move/maximize, multi‑monitor cycling, cursor centering halo
- `Utils.ahk` — Misc utilities (cursor centering, composite actions, prompts)
- `UIA-v2/` — UI Automation v2 library (`UIA.ahk`, `UIA_Browser.ahk`)

---

## Conventions and Building Blocks

### 1) UIA Anchoring Pattern

When targeting dynamic UIs, first find a robust “anchor” element, then navigate relative to it.

- Try several anchors in order of reliability
  - Example anchors: app title link, a stable button (with localized alternatives)
- Wrap each UIA query in its own try/catch so a failed search doesn’t abort the sequence
- Only if all fail, fall back to a generic element and proceed conservatively

Pseudo-steps:

1. `spot := UIA_Browser("ahk_exe App.exe")`
2. Try anchors A, B, C (each inside its own try/catch)
3. If none found, try first `Button`, then first `Text`, finally the first element
4. Proceed with navigation (tabbing, select, invoke) from the anchor

### 2) Disambiguating Many Identical Elements (Tab Strategy)

Use keyboard tabbing to locate the correct instance of a repeated control (e.g., many “Play” buttons).

- Select anchor → small `Sleep()` → send forward `{Tab}` steps
- After each tab, get `UIA.GetFocusedElement()` and inspect `Name` and `Type`
- Use case‑insensitive name matching and filter by control type (e.g., 50000 = Button)
- Bound the attempts (e.g., 6 tabs) and fall back to a safe system key if not found

This mirrors real keyboard navigation and avoids clicking the wrong instance.

#### Example: Spotify Shift+T (Play/Pause)

Challenge: Many elements named “Play”. The correct target is the transport Play button; playlist tiles also have Play.

Approach:

- Find a reliable anchor (e.g., “Enter Full screen” button or other stable element)
- Focus the anchor and tab forward up to 6 positions
- On each tab, check the focused element:
  - If `Name` equals "Play 01011001" and `Type` is 50000 → press Enter
  - Else if `Name` contains "play" (case‑insensitive) and `Type` is 50000 → press Enter
- If not found after 6 tabs, send `{Media_Play_Pause}` as a fallback

Why it works: Respects Spotify’s focus order and avoids ambiguous direct matches to tile play buttons.

### 3) Robust Window Activation (Teams)

Preserve window state and try multiple activation strategies:

- Read original min/max state; restore only if minimized
- Try `WinActivate` + wait, then `ShowWindow(SW_RESTORE)` + `SetForegroundWindow`, then `BringWindowToTop`
- Regex fallback on titles; last‑resort taskbar navigation
- Use non‑blocking overlays when states are unknown

### 4) State Verification via UIA (Mic/Camera)

- Prefer reading `ToggleState` when available
- Fallback to parsing action text in `Name` with multilingual patterns (EN/PT)
- Retry with small sleeps; keep total wait short

### 5) Visual Feedback without Blocking

- Use centered, timed overlay banners instead of modal `MsgBox`
- Keep operations silent; on failure, fall back gracefully

### 6) Fast Button Clicking Patterns

Choose the right pattern based on your use case: binary state toggles vs option selection.

#### 6a) State-Based Toggle (Binary States)

For functions that toggle between two mutually exclusive states where buttons represent opposite actions (e.g., recording vs sending, on vs off).

**Key principles:**

- Use a global boolean variable to track state between hotkey presses
- Use `WaitForButton()` with regex patterns to find the appropriate button based on state
- Requires if/else branching: searches for the "opposite" button based on tracked state
- Minimize delays: only small sleeps (100ms initial, 300ms after state change) for UI to settle
- Pattern-based matching allows flexible button name matching (localized, dynamic text)

**Why it works:**

- State persists between calls, so the function knows which button to target
- Direct button finding via `WaitForButton()` is faster than navigation or searching
- Minimal sleeps keep the toggle responsive while allowing UI to update

**Example:** WhatsApp voice message toggle (`ToggleVoiceMessage()`)

- Global `isRecording` tracks whether recording is active
- If `isRecording` is true → find and click "Send/Stop recording" button
- If `isRecording` is false → find and click "Voice message" button
- Update state after successful click
- Uses regex patterns like `"i)^(Voice message|Record voice message)$"` for flexible matching

#### 6b) Fast Single-Search Pattern (Option Selection)

For clicking one of multiple options where the same option can be selected multiple times (e.g., switching between modes, selecting settings).

**Key principles:**

- Use combined regex pattern to find ANY valid button in one search (e.g., `"i)^(Fast|Thinking)$"`)
- No if/else branching - click whichever button is found
- Update state AFTER clicking based on button name retrieved from the clicked element
- Prefer `Invoke()` when available, fallback to `Click()` for maximum compatibility
- Ultra-minimal delays: 100ms initial, 150ms after click

**Why it's faster:**

- Single `WaitForButton()` call instead of branching searches
- Shorter timeout (1500ms vs 2000-3000ms) since we're not waiting for specific state
- Eliminates state prediction errors - state syncs to reality after each click
- Handles cases where same option is clicked repeatedly without confusion

**Example:** Gemini model toggle (`ToggleGeminiModel()`)

- Global `isGeminiFastModel` tracks current model selection
- Combined pattern `"i)^(Fast|Thinking)$"` finds whichever button is visible
- Click the found button (could be same as last time)
- Update `isGeminiFastModel` based on button name after click
- 3x faster than position-based approaches, more robust than state-branching

**When to use Single-Search vs State-Based:**

- Same option can be clicked again → Single-Search (6b)
- Buttons are truly opposite actions → State-Based (6a)
- Unsure? → Try Single-Search first (faster and more forgiving)

### Choosing the Right Pattern - Decision Guide

```
Need to click a button in a web UI?
│
├─ Multiple identical elements with same name?
│  └─ Use Tab Strategy (Section 2)
│     Navigate from anchor, inspect focused element
│
├─ Binary toggle (on/off, recording/sending)?
│  └─ Use State-Based Toggle (Section 6a)
│     Track state, search for opposite button
│
├─ Selecting from options (same can be clicked again)?
│  └─ Use Fast Single-Search (Section 6b)
│     Combined pattern, click whatever is found
│
└─ Complex position-based requirement?
   └─ Try to avoid - use WaitForButton with patterns instead
      Position-based searches are slower and more fragile
```

**Quick reference:**

- **Tab Strategy**: Multiple "Play" buttons, need the right one
- **State-Based**: Recording ↔ Sending, Mute ↔ Unmute
- **Single-Search**: Fast/Thinking/Custom mode selection, settings options

---

## Patterns Library (Copy/Paste)

### Try‑Find with Fallbacks (per attempt try/catch)

```ahk
try elem := root.FindElement({ Name: "X", Type: "Button" })
catch elem := ""
if !elem {
    try elem := root.FindElement({ Name: "X", Type: 50000 })
    catch elem := ""
}
```

### Tab Navigation with Bounded Attempts

```ahk
anchor.Select()
Sleep 300
maxTabs := 6
found := false
loop maxTabs {
    try focused := UIA.GetFocusedElement()
    if (focused) {
        name := focused.Name
        type := focused.Type
        if (name = "Play 01011001" && type = 50000) {
            found := true
            break
        }
        if ((InStr(name, "play", false) || InStr(name, "tocar", false)) && type = 50000) {
            found := true
            break
        }
    }
    Send "{Tab}"
    Sleep 20
}
if found
    Send "{Enter}"
else
    Send "{Media_Play_Pause}"
```

### Center Cursor Halo (WindowManagement)

Use `#!+q` to centre the cursor on the active window and show a temporary halo for spatial orientation.

### Fast Single-Search Button Pattern (Option Selection)

```ahk
global currentOption := "Fast"  ; tracks last clicked option

ClickAnyOption() {
    global currentOption
    try {
        uia := UIA_Browser()
        Sleep 100
        
        ; Combined pattern finds ANY valid option in one search
        optionPattern := "i)^(Fast|Thinking|Custom)$"
        FindBtn(p) => WaitForButton(uia, p, 1500)
        
        if (btn := FindBtn(optionPattern)) {
            btnName := ""
            try btnName := btn.Name
            
            supportsInvoke := false
            try {
                supportsInvoke := btn.GetPropertyValue(UIA.Property.IsInvokePatternAvailable)
            } catch {
                supportsInvoke := false
            }
            
            clicked := false
            if (supportsInvoke) {
                try {
                    btn.Invoke()
                    clicked := true
                } catch {
                }
            }
            if (!clicked) {
                try {
                    btn.Click()
                    clicked := true
                } catch {
                }
            }
            
            if (clicked) {
                ; Update state based on what was actually clicked
                currentOption := btnName
                Sleep 150
            }
        }
    } catch Error as err {
        ; Silently fail if anything goes wrong
    }
}
```

**Performance benefits:**

- Single search: finds any valid button in one `WaitForButton()` call
- No branching: clicks whichever button is found immediately
- State syncs to reality: updates after successful click, not before
- Faster timeouts: 1500ms vs 2000-3000ms for state-based approaches
- Handles repeated selection: same option can be clicked multiple times

### State-Based Toggle with Quick Button Finding (Binary States)

Use this pattern for true binary state toggles (recording/sending, on/off). For option selection where the same choice can be clicked again, see "Fast Single-Search Button Pattern" above (Section 6b).

```ahk
global isRecording := false          ; persists between hotkey presses

ToggleVoiceMessage() {
    global isRecording

    try {
        chrome := UIA_Browser()
        Sleep 100                    ; minimal delay for UI to settle

        ; Pattern-based button finding for binary states
        voicePattern := "i)^(Voice message|Record voice message)$"
        sendPattern := "i)^(Send|Stop recording)$"
        FindBtn(p) => WaitForButton(chrome, p, 3000)

        if (isRecording) {           ; stop & send
            if (btn := FindBtn(sendPattern)) {
                btn.Click()
                isRecording := false
                Sleep 300            ; allow UI to restore after sending
            }
        } else {                     ; start recording
            if (btn := FindBtn(voicePattern)) {
                btn.Click()
                isRecording := true
            }
        }
    } catch Error as err {
        ; handle error
    }
}
```

**Performance benefits:**

- Fast execution: direct button finding, minimal sleeps
- State-aware: knows which button to target without checking UI state
- Resilient: pattern matching handles localized/dynamic button names

---

## Tips

- Keep sleeps small but sufficient (20–150ms for most operations, up to 300ms only for UI state transitions)
- Always bound loops (tabs, waits). Provide a safe fallback when exceeded
- Localize patterns (EN/PT) for names like "Share", "Microphone", etc.
- Avoid modal `MsgBox` in automations; prefer overlays or silent fallbacks
- Prefer `WaitForButton()` with patterns over position-based or anchor+tab approaches for speed

---

## Updating This Guide

When you add a new automation:

- Document the anchor(s) and why they are reliable
- Document the resolution strategy if multiple identical elements exist
- Note the fallback behavior and limits (e.g., 6 tabs → media key)
