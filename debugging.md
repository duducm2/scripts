# Square Selector Debugging History

## Issue: Mouse not jumping to center of squares + Need more spacing

### Problem Analysis

**Date**: Current
**Symptoms**:

- Mouse cursor not landing in the middle of selected squares
- Squares need more spacing between them
- Positioning appears incorrect/random

### Current Implementation

**Configuration**:

- Square size: 40x40 pixels
- Spacing: 10 pixels (was 3)
- Number of squares: 15

**Position Calculation Logic**:

1. Get current mouse position (startX, startY)
2. First square (i=1): offset = 0, center at (startX, startY)
3. Subsequent squares: offset = (i-1) \* (squareSize + spacing) in direction
4. GUI positioned at: (squareX - squareSize/2, squareY - squareSize/2)
5. Stored center position: (squareX, squareY)

### Math Verification

For square i:

- Offset from start: `(i-1) * (squareSize + spacing) * direction`
- Square center X: `startX + offset` (horizontal) or `startX` (vertical)
- Square center Y: `startY` (horizontal) or `startY + offset` (vertical)
- GUI top-left X: `squareX - squareSize/2`
- GUI top-left Y: `squareY - squareSize/2`
- Actual GUI center should be: `(guiX + squareSize/2, guiY + squareSize/2) = (squareX, squareY)` ✓

### Issues Identified

1. **Spacing too small**: 10 pixels still not enough visual separation
2. **Possible DPI/scaling issue**: GUI positions might be affected by DPI awareness
3. **Integer division**: Using `//` for division might cause rounding issues with odd numbers

### Fixes Applied

**Date**: Current
**Changes Made**:

1. **Increased spacing**: Changed from 10 to 20 pixels between squares

   - Formula: `offset = (i-1) * (squareSize + spacing)`
   - With squareSize=40 and spacing=20: squares are 60 pixels apart (center-to-center)

2. **Fixed math precision**:

   - Changed from integer division (`//`) to floating-point division (`/ 2.0`)
   - Added `Round()` for final position calculations
   - Ensures exact center positioning

3. **Improved variable naming**:

   - Changed `squareX/squareY` to `squareCenterX/squareCenterY` for clarity
   - Makes it explicit that these are center coordinates

4. **Math verification**:
   - Square center: `squareCenterX/Y = startX/Y + (i-1) * (squareSize + spacing) * direction`
   - GUI top-left: `guiX/Y = squareCenterX/Y - squareSize/2`
   - GUI center: `(guiX + squareSize/2, guiY + squareSize/2) = (squareCenterX, squareCenterY)` ✓

### Testing Notes

After reload:

- Verify squares have visible spacing (20px gap)
- Test mouse jumping to square centers
- Check that positions are accurate for all 15 squares
- Test in all 4 directions (up, down, left, right)

---

## Issue: Mouse Not Jumping to Center - Coordinate System Verification

**Date**: Current
**Symptoms**:

- Mouse still not jumping to exact center of squares
- Getting closer but not perfectly centered
- Need to verify coordinate units are the same for mouse movement and GUI sizes

### Root Cause Analysis

**Problem**: Potential coordinate system mismatch or unit difference

The user suspects that the units for:

- Mouse positioning (`SetCursorPos`, `GetCursorPos`)
- GUI window positioning (`Gui.Show()`, `WinGetPos`)
- GUI window sizes (`squareSize = 40`)

Might be different, causing a mismatch in positioning.

### Coordinate System Investigation

**Current Setup**:

1. **DPI Awareness**: `PER_MONITOR_AWARE_V2` is set (line 13)

   - This ensures physical pixels across mixed scaling
   - All coordinates should be in physical pixels

2. **Mouse Position**: `GetCursorPos` via DllCall (line 173)

   - Returns physical pixels (with DPI awareness)

3. **Mouse Movement**: `SetCursorPos` via DllCall (line 569)

   - Expects physical pixels (with DPI awareness)

4. **GUI Positioning**: `Gui.Show("x... y... w... h...")`

   - Uses physical pixels (with DPI awareness)

5. **Window Position Query**: Currently using calculated positions directly
   - We control exact GUI placement and size
   - Calculated center = actual center

**Verification**:

- ✅ `GetCursorPos` → physical pixels (with DPI awareness)
- ✅ `SetCursorPos` → physical pixels (with DPI awareness)
- ✅ `Gui.Show()` → physical pixels (with DPI awareness)
- ✅ All using same DPI awareness context (`PER_MONITOR_AWARE_V2`)
- ✅ **Units are the same**: All coordinates are in physical pixels

### Solution: Use Calculated Positions Directly

**Key Insight**: Since we control exactly where we place the GUI and its size, we can trust our calculated positions directly.

**Math Verification**:

```
Square size: 40 pixels
Spacing: 35 pixels
Distance between centers: 75 pixels

Position calculation:
- Square center: centerX = startX + (i-1) * 75 * direction
- GUI top-left: guiX = centerX - 20 (for 40x40 square)
- GUI center: guiX + 20 = (centerX - 20) + 20 = centerX ✓
```

**Implementation**:

```autohotkey
// STEP 1: Calculate all center positions
calculatedPositions := []
loop numSquares {
    centerX := startX + (i-1) * 75 * direction
    calculatedPositions.Push({ x: centerX, y: centerY })
}

// STEP 2-4: Create and show all GUIs

// STEP 5: Store calculated positions directly
for i, guiInfo in guiArray {
    g_SquareSelectorPositions.Push(guiInfo.calculatedCenter)
}
```

**Why this works**:

- We position GUI so its center is exactly at calculated position
- GUI size is exactly 40x40 pixels
- Math: `guiX = centerX - 20`, so `center = guiX + 20 = centerX` ✓
- No need to query actual positions - our math is correct!

### Expected Results

After implementing:

- Mouse should jump to exact calculated center
- All coordinate systems are verified to use same units (physical pixels)
- Simpler code - no position queries needed
- Should be pixel-perfect alignment

### If Still Not Working

**Debug steps**:

1. **Add debug output**: Log calculated positions vs actual mouse position after jump
2. **Test with fixed position**: Create square at known position (e.g., 100, 100), jump to it, verify
3. **Compare with working code**: Other code uses `GetWindowRect` for window center - can verify with that if needed

---

## Issue: Mouse Still Not Centering on Letters - Actual Window Position Query

**Date**: Current
**Symptoms**:

- Mouse still not centering on letters in squares
- Calculated positions might not match actual GUI positions
- Need to query actual window positions after GUI is shown

### Root Cause Analysis

**Problem**: Even though we calculate exact positions, the actual GUI window positions might differ slightly due to:

- Window borders or padding
- DPI scaling adjustments
- Rounding differences between calculated and actual positions
- Windows positioning the window slightly differently than requested

### Solution: Query Actual Window Positions

**Key Insight**: Instead of trusting calculated positions, query the actual window rectangle using `GetWindowRect` after all windows are shown, then calculate the actual center.

**Implementation**:

```autohotkey
// STEP 5: Wait for all windows to render
Sleep 20  // Increased delay to ensure windows are fully rendered

// STEP 6: Query actual GUI positions using GetWindowRect
for i, guiInfo in guiArray {
    rect := Buffer(16, 0)  // RECT structure
    if (DllCall("GetWindowRect", "ptr", gui.Hwnd, "ptr", rect)) {
        winLeft := NumGet(rect, 0, "int")
        winTop := NumGet(rect, 4, "int")
        winRight := NumGet(rect, 8, "int")
        winBottom := NumGet(rect, 12, "int")

        // Calculate actual center from window rectangle
        actualCenterX := winLeft + (winRight - winLeft) / 2
        actualCenterY := winTop + (winBottom - winTop) / 2

        // Store actual center position (rounded to nearest pixel)
        g_SquareSelectorPositions.Push({ x: Round(actualCenterX), y: Round(actualCenterY) })
    } else {
        // Fallback to calculated position if GetWindowRect fails
        g_SquareSelectorPositions.Push(guiInfo.calculatedCenter)
    }
}
```

**Why this works**:

- `GetWindowRect` returns the actual physical pixel coordinates of the window
- Works with DPI awareness (returns physical pixels, not logical)
- Accounts for any window borders, padding, or positioning adjustments
- Calculates center from actual window bounds, not from calculated positions
- Provides pixel-perfect accuracy

**Benefits**:

- ✅ Uses actual window positions, not calculated ones
- ✅ Accounts for DPI scaling and window positioning differences
- ✅ Provides exact center coordinates for mouse positioning
- ✅ Works across different DPI settings and monitors

### Expected Results

After implementing:

- Mouse should jump to exact center of each square (based on actual window positions)
- Should work correctly across different DPI settings
- Should account for any window positioning differences
- Should provide pixel-perfect alignment with letters

---

## Issue: First Square Should Start After Mouse, Not Centered on It

**Date**: Current
**Symptoms**:

- Letter Q (first square) is currently centered on mouse position
- User wants Q to START AFTER the mouse position, not centered on it

### Fix Applied

**Date**: Current
**Changes Made**:

1. **Added initial offset**:

   ```autohotkey
   initialOffset := (squareSize / 2.0) + spacing
   // 20 (half square) + 35 (spacing) = 55 pixels
   ```

2. **Updated position calculation**:

   ```autohotkey
   // OLD: First square centered on mouse (offset = 0)
   offset := (i - 1) * (squareSize + spacing) * direction

   // NEW: First square starts after mouse (offset = initialOffset)
   offset := (initialOffset + (i - 1) * (squareSize + spacing)) * direction
   ```

3. **Result**:

   - Square 1 (Q): Center at `startX + 55px` (for right direction)
   - Square 2 (W): Center at `startX + 130px` (55 + 75)
   - Square 3 (E): Center at `startX + 205px` (55 + 150)
   - etc.

4. **Why this works**:
   - Initial offset moves first square center 55px past mouse position
   - First square left edge starts at: `(centerX - 20) = (startX + 55 - 20) = startX + 35px`
   - So the square starts 35px (one spacing) after the mouse cursor
   - Subsequent squares continue from there with consistent spacing

### Expected Results

After reload:

- First square (Q) starts AFTER mouse position, not centered on it
- All squares maintain consistent spacing (35px between squares)
- Mouse jumps correctly to each square center
- Works in all 4 directions
