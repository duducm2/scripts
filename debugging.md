# Square Selector Implementation - Summary

## Overview

Implemented a mouse jump system that displays 15 red squares with letters (Q, W, E, R, T, A, S, D, F, G, Z, X, C, V, B) when `Win + Alt + Shift + Arrow Key` is pressed. The user can then press a letter to jump the mouse to the corresponding square.

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
- **Transparency**: 230/255 (semi-transparent)
- **Positioning**: First square (Q) starts AFTER the mouse position, not centered on it
  - Initial offset: 55 pixels from mouse position
  - First square left edge starts 35 pixels after mouse cursor

### 2. Square Positioning

- **Horizontal lines** (Right/Left): Squares arranged horizontally
- **Vertical lines** (Up/Down): Squares arranged vertically
- **Origin**: Based on current mouse cursor position
- **Calculation**: Uses precise center positions, then queries actual window positions using `GetWindowRect` for pixel-perfect accuracy

### 3. Mouse Jump

- **Trigger**: Press the letter key (Q, W, E, etc.) corresponding to the desired square
- **Action**: Mouse cursor jumps to the exact center of the selected square
- **Accuracy**: Uses actual window rectangle positions (not calculated) for pixel-perfect centering
- **Works with DPI awareness**: Uses physical pixel coordinates across different DPI settings

### 4. Cleanup

- **Auto-timeout**: Squares disappear after 1 second if no letter is pressed
- **On selection**: Squares disappear immediately after mouse jump
- **Letter hotkeys**: Disabled when squares are not shown (prevents conflicts)

## Technical Implementation

### Core Functions

1. **`HandleDirectionHotkey(direction)`**

   - Simplified handler for arrow key hotkeys
   - Cleans up any existing selector
   - Sets active direction
   - Calls `ShowSquareSelector()`

2. **`ShowSquareSelector(direction)`**

   - Creates 15 GUI windows (one per square)
   - Calculates positions based on mouse location
   - Positions all GUIs while hidden
   - Shows all GUIs simultaneously
   - Queries actual window positions using `GetWindowRect`
   - Sets up letter key hotkeys
   - Starts 1-second timeout timer

3. **`SelectSquareByIndex(index)`**

   - Called when a letter key is pressed
   - Moves mouse to the center of the corresponding square
   - Cleans up selector

4. **`CleanupSquareSelector()`**
   - Destroys all GUI windows
   - Disables all letter hotkeys
   - Clears arrays
   - Cancels timer

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
**Solution**: Query actual window rectangles using `GetWindowRect` after windows are shown

## Code Structure

- **Hotkey definitions**: Lines 649-671
- **`HandleDirectionHotkey`**: Lines 631-646
- **`ShowSquareSelector`**: Lines 405-560
- **`SelectSquareByIndex`**: Lines 592-629
- **`CleanupSquareSelector`**: Lines 358-402
- **Global variables**: Lines 335-348

## Future Improvements (Optional)

- Add visual feedback when hovering over squares
- Adjustable timeout duration
- Customizable square size and spacing
- Support for more than 15 squares
