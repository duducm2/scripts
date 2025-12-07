# Google Gemini Microphone Button Keep-Alive

## Issue

The microphone button in Google Gemini vanishes after a few seconds when no audio is detected, even though the user wants to keep it active.

## Solution

Implemented a keep-alive timer that uses non-intrusive methods (SetFocus, Select, or mouse hover) to prevent the microphone button from vanishing without clicking it repeatedly. This avoids interrupting the user's speech with multiple clicks.

## Implementation Details

### Global Variable

- `g_GeminiMicKeepAliveTimer`: Tracks whether the keep-alive timer is active

### Timer Function: `KeepGeminiMicrophoneAlive()`

- Checks if still in Gemini window before interacting
- Finds the microphone button using the same method as the main shortcut (via "Fast" button's parent/siblings)
- **State checking**: Checks `IsOffscreen` and `IsEnabled` properties before interacting
- **Non-intrusive methods** (in order of preference):
  1. `SetFocus()` - Sets keyboard focus without clicking
  2. `Select()` - Selects the element without clicking
  3. Mouse hover - Moves mouse to button center without clicking (fallback)
- Only interacts when button state indicates it might vanish (offscreen or disabled)
- Automatically stops if user leaves Gemini window

### Main Shortcut: Shift+Y

- Clicks microphone button immediately (to activate it)
- Starts keep-alive timer (2.5 second interval)
- Timer continues until user leaves Gemini window or presses Shift+Y again (which restarts timer)

## Key Points

- Timer interval: 2.5 seconds (2500ms) - less frequent since we're not clicking
- Uses non-intrusive UIA methods to avoid interrupting speech
- State checking prevents unnecessary interactions
- Timer automatically stops when leaving Gemini window
- All errors are silently caught to prevent interruption
- Uses `UIA.TreeWalkerTrue.TryGetParentElement()` to find parent (not `GetParent()` which doesn't exist)
- Global variable must be declared in timer function scope
- Mouse position is saved and restored when using hover fallback
