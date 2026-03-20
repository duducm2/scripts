---
name: Cursor AI Feed Scroll Bottom
overview: Map Alt+U to scroll to the bottom of the Cursor AI feed using UIA.
todos:
  - id: update_cheat_sheet
    content: Add the Alt+U shortcut to the Cursor cheat sheet in `Shift keys.ahk`.
    status: completed
  - id: implement_alt_u_hotkey
    content: Add a `!u::` hotkey inside the Cursor `#HotIf` block in `Shift keys.ahk` to identify `composer-messages-container` and scroll it to the bottom.
    status: completed
    dependencies: [update_cheat_sheet]
---

# Cursor AI Feed Scroll Bottom

## Analysis / Context:
The user requires an AutoHotkey-based shortcut (`Alt+U`) to immediately scroll to the bottom of the dynamically growing AI feed (Chat/Composer) in Cursor. By analyzing the UIA tree, the target feed container has the `ClassName: "composer-messages-container"`, and individual chat bubbles have `ClassName: "composer-rendered-message"`. Standard navigation keys like `Ctrl + End` are unsupported or undesired in this specific case, meaning we must rely on UIA manipulation.

## Proposed Changes:
1. Update the `cheatSheets["Cursor.exe"]` map to document the `[U]` shortcut explicitly as AutoHotkey-based.
2. Add the `!u::` hotkey inside the `#HotIf IsEditorActive() && WinGetClass("A") != "#32770"` section of `Shift keys.ahk`.
3. Implement UIA logic that attempts to find the `composer-messages-container` and uses `ScrollPattern` or `.ScrollIntoView()` on its deepest child to force the view to the bottom.

## Files to Modify:
- `c:\Users\fie7ca\Documents\scripts\Shift keys.ahk`

## Implementation Strategy:

### Step 1: Update the Cheat Sheet
In `Shift keys.ahk`, locate the `cheatSheets["Cursor.exe"]` declaration.
Under the `--- ALT Shortcuts (ahk = AutoHotkey) ---` section, add:
`⬇️ [U] Scroll AI feed to bottom (ahk-based)`

### Step 2: Implement the Hotkey Logic
In `Shift keys.ahk`, locate the Cursor hotkeys section (`#HotIf IsEditorActive() && WinGetClass("A") != "#32770"`). Add the following AHK code snippet:

```ahk
; Alt + U : Scroll AI feed to bottom (AutoHotkey-based shortcut)
!u::
{
    try {
        hwnd := WinExist("A")
        if (!hwnd) return
        
        root := UIA.ElementFromHandle(hwnd)
        chatContainer := root.FindFirst({ClassName: "composer-messages-container"})
        
        if (chatContainer) {
            ; Method 1: Try using ScrollPattern to scroll vertically to 100%
            try {
                if (chatContainer.GetPropertyValue(UIA.Property.IsScrollPatternAvailable)) {
                    chatContainer.ScrollPattern.SetScrollPercent(-1, 100)
                    return
                }
            } catch {}
            
            ; Method 2 (Fallback): Find the last message and scroll it into view
            messages := chatContainer.FindAll({ClassName: "composer-rendered-message", matchmode: "Substring"})
            if (messages && messages.Length > 0) {
                messages[messages.Length].ScrollIntoView()
            }
        }
    } catch {}
}
```