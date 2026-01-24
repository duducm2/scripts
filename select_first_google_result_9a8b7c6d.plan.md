---
name: Select First Google Search Result
overview: Implement automation to select the first organic search result on Google Search using UIA selectors derived from the provided UI tree.
todos:
  - id: define_selector_logic
    content: Define the UIA traversal logic to locate the first result title (LC20lb) within the main results container (center_col) and target its parent link.
    status: pending
    dependencies: []
  - id: implement_hotkey
    content: Add the Shift+U hotkey to the Google Search context in Shift keys.ahk to execute the selection logic.
    status: pending
    dependencies: ["define_selector_logic"]
---

# Select First Google Search Result

## Analysis / Context

The provided UI tree (`google-search-tree.txt`) reveals the structure of the Google Search results page.
*   **Main Container:** The results are contained within a Group element with `AutomationId: "center_col"` (Line 272).
*   **Result Items:** Organic search results are distinct from Ads (`AutomationId: "tads"`).
*   **Target Element:** The clickable link for a result is a `Link` element (Type 50005).
*   **Identifier:** The text inside the link consistently uses the class name `LC20lb` (e.g., `ClassName: "LC20lb MBeuO DKV0Md"` at Line 296). This is a robust selector for the result title.
*   **Structure:** `Link` (Parent) -> `Text` (Child with Class `LC20lb`).

The goal is to find the first occurrence of this pattern inside the `center_col` container to ensure we target the first organic result and avoid ads or other navigation elements.

## Proposed Changes

Modify `Shift keys.ahk` to add a new hotkey (`Shift+U`) specifically for the Google Search context.

**Logic:**
1.  Get the UIA root element for the Chrome window.
2.  Find the `center_col` group to narrow the search scope.
3.  Inside `center_col`, find the first Text element where `ClassName` contains "LC20lb".
4.  Get the `Parent` of this Text element, which corresponds to the clickable Link.
5.  Invoke or Click the Link.

## Files to Modify

*   `c:\Users\eduev\Meu Drive\12 - Scripts\Shift keys.ahk`
    *   Locate the `#HotIf WinActive("ahk_exe chrome.exe") && InStr(WinGetTitle("A"), "Google")` block (around line 1575).
    *   Insert the new `+u::` hotkey handler.

## Implementation Strategy

1.  **Context Check:** Ensure the hotkey only fires when the active window title contains "Google".
2.  **UIA Initialization:** Use `UIA_Browser()` to attach to the window.
3.  **Traversal:**
    *   `root.FindFirst({AutomationId: "center_col"})`
    *   `centerCol.FindFirst({ClassName: "LC20lb", matchmode: "Substring"})`
    *   `titleText.WalkTree("p")` (Get Parent)
4.  **Action:** Call `Invoke()` on the parent link. If `Invoke` is not supported, fall back to `Click()`.
5.  **Fallback:** If `center_col` is not found (e.g., different layout), try searching from the root, though this might hit Ads first.
