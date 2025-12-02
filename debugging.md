# Outlook Appointment Controls Debugging

## Issue
Reminder and All-day were working, but Privacy, Status, and Category were not working.

## Root Cause
The Alt key sequences were being sent incorrectly. The code was using:
```
Send "{Alt Down}{Alt Up}"
Sleep 150
Send "7"
```

This doesn't properly send Alt+7 as a key combination. AutoHotkey needs the `!` prefix for Alt key combinations.

## Fixes Applied

### 1. ApplyPrivacy (Alt+7)
**Before:**
```ahk
Send "{Alt Down}{Alt Up}"
Sleep 150
Send "7"
Sleep 200
```

**After:**
```ahk
Send "!7"  ; ! prefix = Alt key
Sleep 200
```

### 2. ApplyStatus (Alt+5)
**Before:**
```ahk
Send "{Alt Down}{Alt Up}"
Sleep 150
Send "5"
Sleep 200
```

**After:**
```ahk
Send "!5"
Sleep 300  ; Increased wait time for menu to open
```

Also increased delays between arrow key presses and before Enter to ensure menu navigation works properly.

### 3. ApplyCategory (Alt+6)
**Before:**
```ahk
Send "{Alt Down}{Alt Up}"
Sleep 150
Send "6"
Sleep 200
```

**After:**
```ahk
Send "!6"
Sleep 300  ; Increased wait time for menu to open
```

Also increased delays between Down arrow presses (200ms instead of 150ms) for better reliability.

### 4. ApplyReminder (Alt+8)
**Before:**
```ahk
Send "{Alt Down}{Alt Up}"
Sleep 150
Send "8"
Sleep 200
```

**After:**
```ahk
Send "!8"
Sleep 200
```

## Key Changes Summary
1. Changed all Alt key sequences from `{Alt Down}{Alt Up}` + number to `!number` format
2. Increased delays after opening menus (300ms for Status and Category)
3. Increased delays between arrow key presses for Category (200ms)
4. Added delay before Enter key to ensure selection is ready

## Testing
- Privacy: Alt+7 should now toggle Private On
- Status: Alt+5 opens menu, arrows navigate, Enter confirms
- Category: Alt+6 opens menu, arrows navigate, Enter confirms
- Reminder: Already working (no change needed)
- All-day: Already working (uses UIA directly, no change needed)

