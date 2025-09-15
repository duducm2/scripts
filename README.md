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

---

## Tips

- Keep sleeps small but sufficient (20–100 ms) to avoid racing UI focus
- Always bound loops (tabs, waits). Provide a safe fallback when exceeded
- Localize patterns (EN/PT) for names like "Share", "Microphone", etc.
- Avoid modal `MsgBox` in automations; prefer overlays or silent fallbacks

---

## Updating This Guide

When you add a new automation:

- Document the anchor(s) and why they are reliable
- Document the resolution strategy if multiple identical elements exist
- Note the fallback behavior and limits (e.g., 6 tabs → media key)
