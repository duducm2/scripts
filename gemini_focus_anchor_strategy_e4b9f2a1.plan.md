---
name: Refactor Gemini Focus Logic Using Anchor Element
overview: Update the AutoHotkey script to use the "Open upload file menu" button as a navigation anchor. Instead of targeting the text field directly, focus the anchor button and shift-tab back to the prompt to avoid selecting the "More options" menu.
todos:
  - id: define_anchor_selector
    content: Define the UIA element selector for the "Open upload file menu" button using Name and LocalizedType.
    status: pending
    dependencies: []
  - id: implement_focus_logic
    content: Update the main script to wait for the anchor element and apply the Focus() method (do not Click).
    status: pending
    dependencies: ["define_anchor_selector"]
  - id: implement_backtrack_navigation
    content: Send the Shift+Tab keystroke immediately after focusing the anchor to move the cursor into the text field.
    status: pending
    dependencies: ["implement_focus_logic"]
  - id: verify_execution
    content: Verify the sequence ensures the prompt is active and avoids the "More options" button.
    status: pending
    dependencies: ["implement_backtrack_navigation"]
---

# Refactor Gemini Focus Logic Using Anchor Element

## Analysis / Context
The current implementation for the `Windows+Alt+Shift+I` hotkey attempts to focus the Gemini prompt text field directly. This has proven unreliable, often accidentally selecting the "More options" button in the Google Chrome UI instead of the input field.

To resolve this, we will adopt an "Anchor & Backtrack" strategy. The "Open upload file menu" button is a fixed element immediately following the text field in the tab order. It has distinct UIA properties that make it a reliable target.

## Proposed Changes
1.  **Targeting Strategy:** Stop searching for the text area directly.
2.  **Anchor Identification:** Locate the "Open upload file menu" button using provided UIA properties.
3.  **Navigation Action:** * `Focus()` the anchor button (strictly focus, no click).
    * Send `Shift+Tab` to reverse navigate exactly one step into the text field.

## Technical Details & UIA Tree
**Anchor Element Properties:**
* **Name:** "Open upload file menu"
* **Control Type:** 50000 (Button)
* **Localized Type:** "button"
* **Path Hints:** `{T:30}, {T:26}, {T:0, i:26}`

## Implementation Strategy
1.  **Get UIA Root:** Access the browser window via UIAutomation.
2.  **Find Anchor:** Use `FindFirst` with a condition matching `Name="Open upload file menu"` AND `ControlType=Button`.
3.  **Execute Sequence:**
    * If element found:
        * Call `Element.Focus()`.
        * `Send, +{Tab}` (Shift + Tab).
    * Else:
        * Handle error (optional logging).