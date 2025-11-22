# Square Selector Implementation - Summary

## Overview

Implemented a mouse jump system that displays 15 red squares with letters (Q, W, E, R, T, A, S, D, F, G, Z, X, C, V, B) when `Win + Alt + Shift + Arrow Key` is pressed. The user can then press a letter to jump the mouse to the corresponding square.

**Current Status**: ✅ Working well - stable implementation with session ID management and robust cleanup logic.

## Hotkeys

- **Win + Alt + Shift + Right**: Shows squares to the right of mouse cursor
- **Win + Alt + Shift + Left**: Shows squares to the left of mouse cursor
- **Win + Alt + Shift + Down**: Shows squares below mouse cursor
- **Win + Alt + Shift + Up**: Shows squares above mouse cursor

## Key Features

### 1. Square Display

- **Square size**: 40x40 pixels
- **Spacing**: 35 pixels between squares (center-to-center distance: 75 pixels)
- **Number of squares**: 15 (one for each letter)
- **Letter sequence**: Q, W, E, R, T, A, S, D, F, G, Z, X, C, V, B
- **Color**: Red squares with white bold letters
- **Transparency**: 255/255 (fully opaque) - no transparency for better visibility
- **Positioning**: First square (Q) starts AFTER the mouse position, not centered on it
  - Initial offset: 55 pixels from mouse position
  - First square left edge starts 35 pixels after mouse cursor

### 2. Square Positioning

- **Horizontal lines** (Right/Left): Squares arranged horizontally
- **Vertical lines** (Up/Down): Squares arranged vertically
- **Origin**: Based on current mouse cursor position
- **Calculation**: Uses precise center positions, then queries actual window positions using `GetWindowRect` for pixel-perfect accuracy
- **Rendering**: All squares positioned while hidden, then shown simultaneously for instant appearance

### 3. Mouse Jump

- **Trigger**: Press the letter key (Q, W, E, etc.) corresponding to the desired square
- **Action**: Mouse cursor jumps to the exact center of the selected square
- **Accuracy**: Uses actual window rectangle positions (not calculated) for pixel-perfect centering
- **Works with DPI awareness**: Uses physical pixel coordinates across different DPI settings
- **Index-based selection**: Uses factory function to create handlers with properly captured indices

### 4. Cleanup

- **Auto-timeout**: Squares disappear after 2 seconds if no letter is pressed
- **On selection**: Squares disappear immediately after mouse jump
- **Letter hotkeys**: Disabled when squares are not shown (prevents conflicts)
- **Session ID system**: Prevents old timers from cleaning up new squares when direction changes rapidly
- **Robust cleanup**: Immediate flag disabling prevents letter hotkeys from using stale positions

## Technical Implementation

### Core Functions

1. **`HandleDirectionHotkey(direction)`** (Lines 652-693)

   - Comprehensive handler for arrow key hotkeys
   - Immediately disables active flag and clears positions to prevent stale data
   - Increments session ID to invalidate old timers
   - Cancels any existing timer before cleanup
   - Performs complete cleanup of old selector (GUIs, hotkeys, arrays)
   - Resets lock and waits for cleanup to complete
   - Sets new active direction
   - Calls `ShowSquareSelector()` with proper session management

2. **`ShowSquareSelector(direction)`** (Lines 428-585)

   - Creates 15 GUI windows (one per square)
   - Calculates positions based on mouse location
   - Positions all GUIs while hidden (no rendering delay)
   - Shows all GUIs simultaneously using `Show()` method
   - Waits 20ms for full rendering
   - Queries actual window positions using `GetWindowRect` for pixel-perfect accuracy
   - Stores actual center positions (not calculated) for mouse jump
   - Sets up letter key hotkeys using factory function
   - Starts 2-second timeout timer with session ID validation

3. **`SelectSquareByIndex(index)`** (Lines 617-650)

   - Called when a letter key is pressed
   - Validates selector is still active and positions array is valid
   - Moves mouse to the exact center of the corresponding square using stored position
   - Cleans up selector and releases lock
   - Clears active direction

4. **`CleanupSquareSelector()`** (Lines 382-426)

   - Immediately disables active flag
   - Disables all letter hotkeys (both uppercase and lowercase)
   - Destroys all GUI windows
   - Clears arrays (GUIs and positions)
   - Cancels timer if active
   - Handles errors gracefully with try-catch blocks

5. **`SquareSelectorTimerHandler(sessionID)`** (Lines 352-375)

   - Validates session ID matches current session (prevents stale timer cleanup)
   - Only cleans up if selector is still active and session matches
   - Releases lock and clears active direction on timeout

6. **`CreateSquareSelectorHandler(index)`** (Lines 587-592)

   - Factory function that creates handlers with properly captured indices
   - Ensures each handler gets its own copy of the index value
   - Prevents closure issues with index matching

7. **`SetupLetterKeyListener()`** (Lines 594-615)
   - Creates individual hotkey handlers for each letter
   - Uses factory function to ensure proper index capture
   - Enables both uppercase and lowercase versions of each letter
   - Stores handler references for cleanup

### Position Calculation

```
Initial offset = (squareSize / 2) + spacing
                = (40 / 2) + 35
                = 20 + 35
                = 55 pixels

For square i:
  offset = initialOffset + (i - 1) * (squareSize + spacing)
  For i=1: 55px
  For i=2: 55 + 75 = 130px
  For i=3: 55 + 150 = 205px
  etc.
```

### Window Position Query

After all windows are shown, the code queries the actual window rectangle using `GetWindowRect` to get the exact physical pixel coordinates. The center is calculated as:

```
centerX = winLeft + (winRight - winLeft) / 2
centerY = winTop + (winBottom - winTop) / 2
```

This ensures pixel-perfect accuracy, accounting for any DPI scaling or window positioning adjustments.

### Session ID System

The implementation uses a session ID system to prevent timer conflicts when the user rapidly changes directions:

1. Each time `HandleDirectionHotkey` is called, it increments `g_SquareSelectorSessionID`
2. The timer handler is created with a closure that captures the current session ID
3. When the timer fires, it checks if the session ID matches the current session
4. If the session ID doesn't match, the timer is stale and is ignored
5. This prevents old timers from cleaning up new squares when direction changes quickly

### Rendering Process

The square rendering follows a carefully orchestrated process:

1. **Calculate positions**: All 15 square center positions are calculated first
2. **Create GUIs**: All GUI windows are created with letter text controls
3. **Position while hidden**: All GUIs are positioned using `Show("Hide")` to avoid rendering delays
4. **Set transparency**: All GUIs set to fully opaque (255) for maximum visibility
5. **Batch show**: All GUIs are shown simultaneously using `Show("NA")`
6. **Wait for rendering**: 20ms delay ensures all windows are fully rendered
7. **Query positions**: `GetWindowRect` is called for each window to get actual physical pixel coordinates
8. **Store centers**: Actual center positions are calculated and stored for mouse jump accuracy

### Letter Hotkey System

Letter hotkeys are dynamically enabled/disabled:

- **Enabled**: Only when `g_SquareSelectorActive` is true
- **Factory pattern**: Each letter gets a handler created by `CreateSquareSelectorHandler(index)` to ensure proper index capture
- **Case insensitive**: Both uppercase and lowercase versions are enabled
- **Cleanup**: All hotkeys are disabled immediately when selector is cleaned up
- **Handler storage**: Handler references are stored for potential future cleanup (currently not needed as Hotkey("Off") works)

## Solved Issues

### 1. Variable Name Conflict

**Problem**: Global variable `gui` conflicted with local variable `gui`  
**Solution**: Renamed local variable to `squareGuiObj` to avoid conflict

### 2. Hotkeys Not Triggering

**Problem**: Complex critical section logic was blocking hotkeys  
**Solution**: Simplified to basic cleanup and function calls

### 3. Squares Not Appearing

**Problem**: `SetWindowPos` wasn't reliably showing windows  
**Solution**: Switched to `Show()` method after positioning windows with `Show("Hide")`

### 4. Mouse Centering Accuracy

**Problem**: Calculated positions didn't always match actual window positions  
**Solution**: Query actual window rectangles using `GetWindowRect` after windows are shown, with 20ms delay for full rendering

### 5. Timer Conflicts

**Problem**: Old timers from previous direction changes were cleaning up new squares  
**Solution**: Implemented session ID system - each new selector gets a unique session ID, timers validate session ID before cleanup

### 6. Stale Position Data

**Problem**: Letter hotkeys could use old positions when direction changed rapidly  
**Solution**: Immediately disable active flag and clear positions array at start of `HandleDirectionHotkey`

### 7. Index Closure Issues

**Problem**: Letter hotkey handlers had incorrect index values due to closure issues  
**Solution**: Created factory function `CreateSquareSelectorHandler()` that properly captures index at creation time

## Code Structure

- **Global variables**: Lines 333-349

  - `g_SquareSelectorActive`: Boolean flag for active state
  - `g_SquareSelectorGuis`: Array of GUI objects
  - `g_SquareSelectorPositions`: Array of {x, y} positions
  - `g_SquareSelectorLetters`: Array of letter strings
  - `g_SquareSelectorTimer`: Timer function reference
  - `g_SquareSelectorLetterMap`: Map for letter-to-index (currently unused)
  - `g_SquareSelectorSessionID`: Unique session identifier
  - `g_SquareSelectorHotkeyHandlers`: Array of hotkey handler references
  - `g_SquareSelectorLock`: Lock flag (currently used for state management)
  - `g_ActiveDirection`: Current active direction string

- **Timer management**: Lines 352-380

  - `SquareSelectorTimerHandler(sessionID)`: Validates session and cleans up on timeout
  - `CreateTimerHandler(sessionID)`: Factory function for session-bound timer handlers

- **Cleanup function**: Lines 382-426

  - `CleanupSquareSelector()`: Comprehensive cleanup of all resources

- **Main display function**: Lines 428-585

  - `ShowSquareSelector(direction)`: Creates, positions, and shows all squares

- **Selection handlers**: Lines 587-650

  - `CreateSquareSelectorHandler(index)`: Factory for index-bound handlers
  - `SetupLetterKeyListener()`: Sets up all letter hotkeys
  - `SelectSquareByIndex(index)`: Handles letter key press and mouse jump

- **Direction handler**: Lines 652-693

  - `HandleDirectionHotkey(direction)`: Main entry point for arrow key hotkeys

- **Hotkey definitions**: Lines 695-719
  - `#!+Right::`, `#!+Left::`, `#!+Down::`, `#!+Up::`: Arrow key hotkeys

## Current Working Snapshot (Documented)

This documentation reflects the current stable implementation as of the latest update. The system is working well with:

- ✅ Robust session ID management preventing timer conflicts
- ✅ Pixel-perfect mouse positioning using actual window rectangles
- ✅ Proper cleanup preventing resource leaks
- ✅ Immediate flag disabling preventing stale data usage
- ✅ Factory function pattern ensuring correct index capture
- ✅ Simultaneous GUI rendering for instant appearance
- ✅ Full opacity (255) for maximum visibility
- ✅ 2-second timeout for user selection
- ✅ Support for all four directions (Up, Down, Left, Right)

### Key Implementation Strengths

1. **Session ID System**: Prevents race conditions when rapidly changing directions
2. **Actual Position Query**: Uses `GetWindowRect` after rendering for pixel-perfect accuracy
3. **Immediate State Management**: Active flag and positions cleared immediately on direction change
4. **Proper Closure Handling**: Factory function ensures each handler has correct index
5. **Graceful Error Handling**: Try-catch blocks prevent crashes from cleanup errors
6. **DPI Awareness**: Uses physical pixel coordinates across mixed DPI settings

## Click Mode Feature Development

### Overview

Added a click mode feature that allows users to toggle between normal mode (mouse jump + loop) and click mode (mouse jump + click + exit). When click mode is active, squares turn blue to indicate the mode change.

### Implementation History

#### Initial Attempts (Failed)

1. **First Approach**: Tried to detect CTRL key being held while pressing letter/number

   - Problem: CTRL+letter combinations conflicted with system shortcuts
   - Result: Unreliable, often didn't work

2. **Second Approach**: Added CTRL toggle to activate click mode
   - Square colors change from red to blue when CTRL is pressed
   - Click mode flag (`g_SquareSelectorClickMode`) tracks state
   - Problem: Click didn't register - mouse moved but no click happened
   - Attempted fixes:
     - Various delay timings (50ms, 100ms, 200ms, 300ms)
     - Different click methods (`Click()`, `Click(x y)`, `DllCall mouse_event`)
     - Window activation before clicking
     - Updating `g_LastMouseClickTick` to prevent interference

#### Root Cause Discovery

**Problem**: The square selector GUIs are created with `+AlwaysOnTop -Caption +ToolWindow`, making them topmost windows. When trying to click at the square position, the click was hitting the square GUI itself, not the window underneath it.

**Evidence**: Mouse would move to position correctly, but clicks never registered. User reported mouse was centering on previously active window (suggesting MonitorActiveWindow was interfering, but that wasn't the root issue).

#### Solution (Current Implementation)

**Fix**: Destroy/hide the squares BEFORE clicking, so the click hits the actual window underneath.

**New Flow**:

1. User presses letter/number in click mode (blue squares visible)
2. Store target position
3. **Immediately hide/destroy ALL squares** (removes AlwaysOnTop blockers)
4. Wait 50ms for GUI cleanup to complete
5. Find window at target position (now that squares are gone)
6. Activate that window
7. Move mouse to target position
8. Update `g_LastMouseClickTick` to prevent MonitorActiveWindow interference
9. Perform click (now hits actual window)
10. Complete cleanup

**Key Changes**:

- Moved square cleanup to BEFORE click attempt
- Added explicit `Hide()` before `Destroy()` for instant visual removal
- Window detection happens AFTER squares are removed (more reliable)
- Cleaner sequence: remove blocker → find target → click

### Current Status

✅ Click mode toggle works (CTRL toggles, squares change color)
✅ Squares are properly destroyed before clicking
✅ Click now hits the actual window underneath
✅ No interference from AlwaysOnTop GUI windows

### Technical Details

- **Click mode flag**: `g_SquareSelectorClickMode` - tracks whether click mode is active
- **Visual indicator**: Squares turn blue (`0000FF`) when click mode active, red (`FF0000`) when normal
- **CTRL hotkey**: Enabled only when squares are visible, toggles click mode
- **Cleanup order**: Squares → Hotkeys → Click → Final cleanup

## Future Improvements (Optional)

- Add visual feedback when hovering over squares
- Adjustable timeout duration (currently hardcoded to 5000ms)
- Customizable square size and spacing (currently 40px and 35px)
- Support for more than 20 squares
- Visual indicator showing which square will be selected on key press
