# AutoHotkey Mouse Navigation Tool: Architectural Plan

## 1. Introduction

This document outlines a structured, step-by-step architectural plan for developing an AutoHotkey (AHK) script that provides keyboard-driven mouse navigation, similar to "Mousemaster" or "Vimium." This tool is designed to enhance workflow during UI evaluations by enabling rapid, keyboard-only interaction with on-screen elements, while adhering to corporate endpoint security policies by avoiding compiled third-party executables.

## 2. Tool Overview

The objective is to create a native AutoHotkey solution that allows users to trigger an overlay displaying character hints over clickable UI elements. Users can then type these hints to interact with the corresponding elements without using the physical mouse.

## 3. Technical Requirements

*   **Activation Hotkey:** The script will be activated by the hotkey `Ctrl+Alt+Win+C` (`^!#c`).
*   **UI Scanning:** Utilize the AutoHotkey UIAutomation (UIA) library to scan the currently active window and identify bounding boxes of clickable elements (buttons, links, text inputs).
*   **Visual Overlay:** Create a transparent, click-through GUI that spans the active window. This overlay will draw text labels (character hints like "A", "B", "AA") anchored to the coordinates of the detected UIA elements.
*   **Input Interception:** The script must trap keyboard inputs while the overlay is active, preventing keystrokes from bleeding into the underlying application.
*   **Action and Cleanup:** Implement a function that matches typed characters to the corresponding element, moves the mouse, performs a click, and instantly destroys the GUI to return keyboard control to normal operation.
*   **Performance Focus:** The UIA tree walking and GUI rendering must be fast enough to support rapid workflow during e-commerce UI evaluations.

## 4. Architectural Plan (Step-by-Step)

### Phase 1: Activation and Initialization

1.  **Define Hotkey:** Map the `^!#c` hotkey to a primary function that initiates the mouse navigation mode.
2.  **Check Active Window:** Upon hotkey activation, verify that there is an active window to operate on. If not, gracefully exit or provide user feedback.
3.  **Global Flags:** Set a global flag (e.g., `MousemasterActive := true`) to indicate the tool's active state, controlling input interception and other behaviors.
4.  **Store Window ID:** Capture the unique ID of the currently active window (`WinGet, ActiveWinID, ID, A`) for UIA operations and overlay positioning.

### Phase 2: UI Scanning and Element Detection

1.  **UIA Library Integration:** Leverage AutoHotkey's built-in UIAutomation capabilities or a well-regarded UIA wrapper for AHK (if available/necessary, while respecting the no-compiled-executables rule).
2.  **Target Active Window:** Initialize UIA to focus specifically on the `ActiveWinID` identified in Phase 1.
3.  **Recursive Tree Walking:** Implement a recursive function to traverse the UIA element tree of the active window.
4.  **Element Identification:** Within the recursive function, identify elements that are considered "clickable." This will likely involve checking `ControlType` properties (e.g., `UIA_ButtonControlTypeId`, `UIA_HyperlinkControlTypeId`, `UIA_EditControlTypeId`) and `IsEnabled` status.
5.  **Bounding Box Extraction:** For each identified clickable element, extract its `BoundingRectangle` properties (x, y, width, height) to determine its screen position and size.
6.  **Filter Irrelevant Elements:** Develop logic to filter out hidden elements, zero-sized elements, or elements that are visually not clickable (e.g., disabled buttons, decorative text that UIA identifies as a link but isn't interactive). Store a list of `[Hint, X_Coord, Y_Coord, ElementObject]` for each valid element.

### Phase 3: Visual Overlay Generation

1.  **GUI Creation:** Create a new, transparent, always-on-top, click-through AHK GUI window. This GUI must be exactly sized and positioned to cover the `ActiveWinID` window (`Gui, +ToolWindow +AlwaysOnTop -Caption +E0x20 +LastFound`).
2.  **Background Transparency:** Ensure the GUI itself is transparent (`Gui, Color, <Color>, <TransparencyValue>`). The `+E0x20` extended style makes it click-through.
3.  **Hint Generation Logic:**
    *   Iterate through the list of identified clickable elements.
    *   Assign unique, short character hints (e.g., "A", "B", "C", then "AA", "AB", "AC") to each element. Prioritize single-character hints for elements appearing earlier in the UIA tree or more prominent positions.
    *   The hint generation should be deterministic and efficient.
4.  **Drawing Hints:** For each element:
    *   Calculate the optimal display position for the hint text (e.g., centered within the bounding box, or at a consistent corner).
    *   Add `Gui, Add, Text` controls to the overlay GUI at the calculated coordinates, displaying the generated hints. Use a clear, contrasting font and background for readability.
5.  **Show GUI:** Display the GUI (`Gui, Show, % "NoActivate x" ActiveWinX " y" ActiveWinY " w" ActiveWinW " h" ActiveWinH, MousemasterOverlay`) ensuring it doesn't steal focus.

### Phase 4: Input Interception and Hint Matching

1.  **Hotkey/Input Hook:** While `MousemasterActive` is true and the overlay GUI is active, use `Input` or a `Hotkey` function with a `~` prefix to capture all subsequent keyboard input, preventing it from reaching the underlying application.
2.  **Buffer Input:** Store typed characters in a temporary buffer.
3.  **Match Logic:**
    *   As each character is typed, attempt to match the current buffer content against the generated hints.
    *   If a unique, full hint match is found (e.g., "A" matches hint "A"), proceed to Phase 5.
    *   If a partial match exists (e.g., "A" could lead to "AA", "AB"), continue buffering.
    *   If no match or partial match exists, clear the buffer and provide feedback (optional: visual or auditory cue).
4.  **Escape/Cancel:** Implement an escape hotkey (e.g., `Esc`) to cancel the overlay mode, clear the buffer, and return to normal operation.

### Phase 5: Action Execution and Cleanup

1.  **Target Resolution:** Once a hint is fully matched, retrieve the stored `X_Coord`, `Y_Coord`, and potentially the UIA `ElementObject` for the corresponding clickable element.
2.  **Mouse Movement:** Use `MouseMove, X_Coord, Y_Coord, 0` to instantly move the mouse cursor to the target coordinates.
3.  **Click Action:** Perform a click using `Click` or `ControlClick` on the target element. If the UIA `ElementObject` is available, consider using its `Invoke()` method if the AHK UIA wrapper supports it, as this can be more robust than a simple mouse click.
4.  **GUI Destruction:** Immediately destroy the overlay GUI (`Gui, Destroy`).
5.  **Reset Global Flag:** Set `MousemasterActive := false` to indicate the tool is no longer active.
6.  **Restore Keyboard:** Release the keyboard input hook/hotkeys to return normal keyboard operation to the underlying application.

### Phase 6: Performance Optimization Considerations

1.  **Efficient UIA Tree Walking:**
    *   Minimize the depth of UIA traversal if possible, focusing only on relevant branches.
    *   Cache UIA elements where appropriate to avoid repeated lookups.
    *   Avoid unnecessary property retrievals; only fetch `BoundingRectangle`, `ControlType`, and `IsEnabled`.
2.  **Optimized GUI Rendering:**
    *   Avoid recreating the GUI window repeatedly; instead, if possible, hide and show it, or update its contents efficiently if the active window changes. (For initial version, destroy and recreate is acceptable).
    *   Use `Gui, Font` and `Gui, Color` sparingly. Define styles once.
    *   Pre-calculate hint positions to reduce overhead during drawing.
    *   Consider using a single `Gui, Add, Text` control with a multi-line string if performance with many individual text controls is an issue (less flexible for positioning though).
3.  **Fast Input Handling:**
    *   The `Input` command in AHK is generally efficient for character capture.
    *   Optimize the hint matching algorithm (e.g., use a hash map or trie-like structure if hints become complex or numerous).
4.  **Error Handling:** Implement robust error handling for UIA failures, window changes, or unexpected element structures to prevent script crashes and maintain responsiveness.
