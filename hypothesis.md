```python?code_reference&code_event_index=2
with open("Shift keys.ahk", "r", encoding="utf-8", errors="ignore") as f:
    content = f.read()
print(content[:2000])
print("---")
print(content[-2000:])


```
```text?code_stdout&code_event_index=2
/* ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** **
    * Win + Alt + Shift symbol layer shortcuts (AHK v2)
    * â€¢ Provides system - wide symbol shortcuts
    ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** /
    /********************************************************************
     *   AVAILABLE WIN+ALT+SHIFT COMBINATIONS
     *   The following combinations are not currently in use:
     *
     *   Letters: P, U
     *   Numbers: (all numbers 0-9 are used)
     *   Symbols: ; ' [ ] \ | ` ~ @ # $ % ^ & * ( ) - _ = + { } : " < > ? /
     *
     *   Note: Some combinations use Ctrl+Alt+Shift+Arrow keys for extended mouse movement
********************************************************************/
#Requires AutoHotkey v2.0+

#SingleInstance Force

SetTitleMatchMode 2

#include %A_ScriptDir%\env.ahk
#include UIA-v2\Lib\UIA.ahk
#include UIA-v2\Lib\UIA_Browser.ahk
#include %A_ScriptDir%\Utils.ahk
#include %A_ScriptDir%\aux\ShiftKeysIPC.ahk

; --- Global Variables ---
global DEBUG_LOG_PATH := A_ScriptDir "\.cursor\debug.log"
; Phase 5: Gate debug I/O; set to true only when diagnosing (avoids file I/O in hot paths).
global DEBUG_SHIFTKEYS := false

; Debug mode agent logging (runtime evidence for this session only)
; (disabled) agent log debug-31b036

; Helper function for safe debug logging with retry on file lock
; Handles file locking gracefully by retrying with exponential backoff
; No-op when DEBUG_SHIFTKEYS is false (production).
SafeDebugLog(text) {
    if (!DEBUG_SHIFTKEYS)
        return true
    maxRetries := 3
    retryDelay := 10
    loop maxRetries {
        try {
            FileAppend text, DEBUG_LOG_PATH
            return true
        } catch Error as err {
            ; Check if error has Number property before accessing it
            ; File lock error is typically error code 32
            hasNumber := false
            errNumber := 0
     
---
= UIA_Browser("ahk_id " chatGPTHwnd)
    catch {
        return ; Failed to get UIA context
    }

    start := A_TickCount
    btn := ""

    ; Wait for the target button to appear and monitor it until it disappears
    deadline := (timeout > 0) ? (start + timeout) : 0
    while (timeout <= 0 || (A_TickCount < deadline)) {
        btn := ""
        for n in buttonNames {
            try btn := cUIA.FindElement({ Name: n, Type: "Button" })
            catch {
                btn := ""
            }
            if !btn {
                ; Fallback: substring match without strict type (handles UI variations)
                try btn := cUIA.FindElement({ Name: n, matchmode: "Substring" })
                catch {
                    btn := ""
                }
            }
            if btn
                break
        }
        if btn {
            ShowSmallLoadingIndicator_ChatGPT(stateText)
            while btn && (timeout <= 0 || (A_TickCount < deadline)) {
                Sleep 250
                btn := ""
                for n in buttonNames {
                    try btn := cUIA.FindElement({ Name: n, Type: "Button" })
                    catch {
                        btn := ""
                    }
                    if !btn {
                        ; Fallback: substring match without strict type
                        try btn := cUIA.FindElement({ Name: n, matchmode: "Substring" })
                        catch {
                            btn := ""
                        }
                    }
                    if btn
                        break
                }
            }
            break
        }
        Sleep 250
    }

    ; Chime only for real AI answering events (not transcription)
    try {
        if (InStr(StrLower(stateText), "transcrib") = 0)
            PlayCompletionChime_ChatGPT()
    } catch {
    }
    ; Always hide the indicator at the end (debounced safety)
    try HideSmallLoadingIndicator_ChatGPT()
    catch {
    }
}


```
```python?code_reference&code_event_index=6
import re

with open("Shift keys.ahk", "r", encoding="utf-8", errors="ignore") as f:
    content = f.read()

# Find the hotkey definition
match = re.search(r'(?i)(#!?[+^\w\s]*j::[\s\S]*?(?:return|}))', content)
if match:
    print(match.group(1)[:3000])
else:
    # try another way
    match = re.search(r'(?i)(fast.*?copy[\s\S]{0,2000})', content)
    if match:
        print(match.group(1))
    else:
        print("Not found")

```
```text?code_stdout&code_event_index=6
#!+j:: {
    global gFastCopyModeActive, FAST_COPY_HOLD_REPEAT_MS
    if (gFastCopyModeActive) {
        FastCopyMode_Finish()
        return

```
```python?code_reference&code_event_index=9
import re

with open("Shift keys.ahk", "r", encoding="utf-8", errors="ignore") as f:
    content = f.read()

# Find the entire FastCopyMode block
match = re.search(r'(?i)(gFastCopyModeActive.*?)(?=/\*\*\*|\n\n;|\Z)', content, re.DOTALL)
if match:
    # Just print all occurrences of "FastCopy" to get a sense of where it is
    lines = content.split('\n')
    for i, line in enumerate(lines):
        if 'FastCopy' in line or '#!+j' in line:
            print(f"{i}: {line.strip()}")


```
```text?code_stdout&code_event_index=9
1561: global gFastCopyModeActive := false
1562: global gFastCopyCount := 0
1563: global gFastCopyPasteTargetHwnd := 0
1564: global gFastCopyLastSuccessfulCount := 0
1565: global gFastCopyScreenshotQueue := []
1566: global gFastCopyLastScreenshotQueue := []
1568: FastCopyMode_DebugLog(hypothesisId, location, message, data := "") {
1579: try parts.Push('"' FastCopyMode_JsonEscape(k) '":"' FastCopyMode_JsonEscape(v) '"')
1583: dataJson := '{"value":"' FastCopyMode_JsonEscape(data) '"}'
1589: . '"hypothesisId":"' FastCopyMode_JsonEscape(hypothesisId) '",'
1590: . '"location":"' FastCopyMode_JsonEscape(location) '",'
1591: . '"message":"' FastCopyMode_JsonEscape(message) '",'
1601: FastCopyMode_JsonEscape(s) {
1618: FastCopyMode_ClipboardHasImage() {
1631: FastCopyMode_ReleaseHotkeyModifiers() {
1638: FastCopyMode_CaptureScreenshotToQueue() {
1639: global gFastCopyScreenshotQueue
1641: FastCopyMode_DebugLog("H1", "Shift keys.ahk:FastCopyMode_CaptureScreenshotToQueue", "capture_enter", Map(
1642: "queueLenBefore", gFastCopyScreenshotQueue.Length,
1643: "hasImageBefore", FastCopyMode_ClipboardHasImage() ? "1" : "0"
1653: FastCopyMode_DebugLog("H2", "Shift keys.ahk:FastCopyMode_CaptureScreenshotToQueue", "clipwait_throw", Map(
1654: "hasImageAfter", FastCopyMode_ClipboardHasImage() ? "1" : "0"
1661: gFastCopyScreenshotQueue.Push(snap)
1664: FastCopyMode_DebugLog("H1", "Shift keys.ahk:FastCopyMode_CaptureScreenshotToQueue", "capture_pushed", Map(
1665: "queueLenAfter", gFastCopyScreenshotQueue.Length,
1666: "hasImageAfter", FastCopyMode_ClipboardHasImage() ? "1" : "0",
1672: FastCopyMode_DebugLog("H2", "Shift keys.ahk:FastCopyMode_CaptureScreenshotToQueue", "clipboardall_failed", Map(
1673: "queueLenAfter", gFastCopyScreenshotQueue.Length,
1674: "hasImageAfter", FastCopyMode_ClipboardHasImage() ? "1" : "0"
1679: FastCopyMode_PasteScreenshotQueue(queue) {
1687: FastCopyMode_DebugLog("H3", "Shift keys.ahk:FastCopyMode_PasteScreenshotQueue", "paste_enter", Map(
1689: "hasImageAtEnter", FastCopyMode_ClipboardHasImage() ? "1" : "0"
1697: FastCopyMode_DebugLog("H3", "Shift keys.ahk:FastCopyMode_PasteScreenshotQueue", "paste_iter_ready", Map(
1701: "hasImageNow", FastCopyMode_ClipboardHasImage() ? "1" : "0"
1704: FastCopyMode_DebugLog("H4", "Shift keys.ahk:FastCopyMode_PasteScreenshotQueue", "paste_iter_sent", Map(
1710: FastCopyMode_DebugLog("H4", "Shift keys.ahk:FastCopyMode_PasteScreenshotQueue", "paste_iter_failed", Map(
1718: FastCopyMode_DebugLog("H3", "Shift keys.ahk:FastCopyMode_PasteScreenshotQueue", "paste_exit", Map(
1740: FastCopyMode_IsActive() {
1741: global gFastCopyModeActive
1742: return gFastCopyModeActive
1745: FastCopyMode_OnCopy() {
1746: global gFastCopyCount
1747: gFastCopyCount += 1
1748: FastCopyModeBanner_Update(gFastCopyCount)
1751: FastCopyMode_Start() {
1752: global gFastCopyModeActive, gFastCopyCount, gFastCopyPasteTargetHwnd, gFastCopyScreenshotQueue
1754: gFastCopyPasteTargetHwnd := WinGetID("A")
1756: gFastCopyPasteTargetHwnd := 0
1758: gFastCopyCount := 0
1759: gFastCopyScreenshotQueue := []
1760: gFastCopyModeActive := true
1762: FastCopyModeBanner_Show()
1764: gFastCopyModeActive := false
1769: FastCopyMode_Finish() {
1770: global gFastCopyModeActive, gFastCopyCount, gFastCopyPasteTargetHwnd
1771: global gFastCopyScreenshotQueue, gFastCopyLastScreenshotQueue
1772: count := gFastCopyCount
1773: shotCount := IsObject(gFastCopyScreenshotQueue) ? gFastCopyScreenshotQueue.Length : 0
1775: FastCopyModeBanner_Hide()
1777: gFastCopyModeActive := false
1778: gFastCopyCount := 0
1782: FastCopyMode_ReleaseHotkeyModifiers()
1784: FastCopyMode_DebugLog("H5", "Shift keys.ahk:FastCopyMode_Finish", "finish_paste_start", Map(
1789: FastCopyMode_PasteScreenshotQueue(gFastCopyScreenshotQueue)
1791: gFastCopyLastScreenshotQueue := gFastCopyScreenshotQueue.Clone()
1793: gFastCopyLastScreenshotQueue := []
1797: FastCopyMode_DebugLog("H5", "Shift keys.ahk:FastCopyMode_Finish", "finish_remaining", Map(
1805: global gFastCopyLastSuccessfulCount
1806: gFastCopyLastSuccessfulCount := count
1814: FastCopyMode_RepeatLastPaste() {
1815: global gFastCopyLastSuccessfulCount
1816: global gFastCopyLastScreenshotQueue
1817: if (gFastCopyLastSuccessfulCount < 1) {
1823: FastCopyMode_ReleaseHotkeyModifiers()
1824: if (IsObject(gFastCopyLastScreenshotQueue) && gFastCopyLastScreenshotQueue.Length > 0) {
1825: FastCopyMode_PasteScreenshotQueue(gFastCopyLastScreenshotQueue)
1826: remaining := gFastCopyLastSuccessfulCount - gFastCopyLastScreenshotQueue.Length
1830: ExecuteSequentialPaste(gFastCopyLastSuccessfulCount)
1837: #!+j:: {
1838: global gFastCopyModeActive, FAST_COPY_HOLD_REPEAT_MS
1839: if (gFastCopyModeActive) {
1840: FastCopyMode_Finish()
1847: FastCopyMode_RepeatLastPaste()
1849: FastCopyMode_Start()
1852: #HotIf FastCopyMode_IsActive()
1853: ~^c:: FastCopyMode_OnCopy()
1854: ~PrintScreen:: FastCopyMode_OnCopy()
1856: FastCopyMode_OnCopy()
1857: FastCopyMode_CaptureScreenshotToQueue()

```
```python?code_reference&code_event_index=12
with open("Shift keys.ahk", "r", encoding="utf-8", errors="ignore") as f:
    lines = f.read().split('\n')
print("\n".join(lines[1560:1865]))

```
```text?code_stdout&code_event_index=12
global FAST_COPY_HOLD_REPEAT_MS := 700
global gFastCopyModeActive := false
global gFastCopyCount := 0
global gFastCopyPasteTargetHwnd := 0
global gFastCopyLastSuccessfulCount := 0
global gFastCopyScreenshotQueue := []
global gFastCopyLastScreenshotQueue := []

FastCopyMode_DebugLog(hypothesisId, location, message, data := "") {
    ; #region agent log
    ; Writes NDJSON to debug-1bed80.log (Debug session: 1bed80)
    try {
        runId := "pre-fix"
        logPath := "C:\Users\fie7ca\Documents\scripts\debug-1bed80.log"
        ; Keep data small and non-sensitive; accept either a string or a Map-like object.
        dataJson := "{}"
        if (IsObject(data)) {
            parts := []
            for k, v in data {
                try parts.Push('"' FastCopyMode_JsonEscape(k) '":"' FastCopyMode_JsonEscape(v) '"')
            }
            dataJson := "{" (parts.Length ? parts.Join(",") : "") "}"
        } else if (data != "") {
            dataJson := '{"value":"' FastCopyMode_JsonEscape(data) '"}'
        }
        line := '{'
            . '"sessionId":"1bed80",'
            . '"timestamp":' A_TickCount + 0 ','
            . '"runId":"' runId '",'
            . '"hypothesisId":"' FastCopyMode_JsonEscape(hypothesisId) '",'
            . '"location":"' FastCopyMode_JsonEscape(location) '",'
            . '"message":"' FastCopyMode_JsonEscape(message) '",'
            . '"data":' dataJson
            . '}'
        FileAppend(line "`n", logPath, "UTF-8")
    } catch {
        ; never break user flow
    }
    ; #endregion agent log
}

FastCopyMode_JsonEscape(s) {
    ; #region agent log
    try {
        if (s = "")
            return ""
    } catch {
        return ""
    }
    s := "" s
    s := StrReplace(s, "\", "\\")
    s := StrReplace(s, '"', '\"')
    s := StrReplace(s, "`r", "\r")
    s := StrReplace(s, "`n", "\n")
    return s
    ; #endregion agent log
}

FastCopyMode_ClipboardHasImage() {
    ; #region agent log
    ; CF_DIB=8, CF_DIBV5=17, CF_BITMAP=2
    try {
        return !!(DllCall("IsClipboardFormatAvailable", "UInt", 8, "Int")
            || DllCall("IsClipboardFormatAvailable", "UInt", 17, "Int")
            || DllCall("IsClipboardFormatAvailable", "UInt", 2, "Int"))
    } catch {
        return false
    }
    ; #endregion agent log
}

FastCopyMode_ReleaseHotkeyModifiers() {
    ; Ensure the Win+Alt+Shift hotkey modifiers can't leak into paste keys.
    ; Releasing modifiers does not activate or focus any other window.
    Send "{LWin up}{RWin up}{Alt up}{Shift up}{Ctrl up}"
    Sleep 30
}

FastCopyMode_CaptureScreenshotToQueue() {
    global gFastCopyScreenshotQueue

    FastCopyMode_DebugLog("H1", "Shift keys.ahk:FastCopyMode_CaptureScreenshotToQueue", "capture_enter", Map(
        "queueLenBefore", gFastCopyScreenshotQueue.Length,
        "hasImageBefore", FastCopyMode_ClipboardHasImage() ? "1" : "0"
    ))

    ; Give the OS a moment to push the screenshot into the clipboard.
    ; Alt+PrintScreen updates the clipboard with an image; rapid captures can overwrite each other
    ; unless we snapshot the clipboard right away.
    Sleep 40
    try {
        ClipWait 0.6, 1
    } catch {
        FastCopyMode_DebugLog("H2", "Shift keys.ahk:FastCopyMode_CaptureScreenshotToQueue", "clipwait_throw", Map(
            "hasImageAfter", FastCopyMode_ClipboardHasImage() ? "1" : "0"
        ))
        return
    }

    try {
        snap := ClipboardAll()
        gFastCopyScreenshotQueue.Push(snap)
        snapSize := ""
        try snapSize := snap.Size
        FastCopyMode_DebugLog("H1", "Shift keys.ahk:FastCopyMode_CaptureScreenshotToQueue", "capture_pushed", Map(
            "queueLenAfter", gFastCopyScreenshotQueue.Length,
            "hasImageAfter", FastCopyMode_ClipboardHasImage() ? "1" : "0",
            "snapType", Type(snap),
            "snapSize", snapSize
        ))
    } catch {
        ; If clipboard snapshot fails, just skip (count still increments).
        FastCopyMode_DebugLog("H2", "Shift keys.ahk:FastCopyMode_CaptureScreenshotToQueue", "clipboardall_failed", Map(
            "queueLenAfter", gFastCopyScreenshotQueue.Length,
            "hasImageAfter", FastCopyMode_ClipboardHasImage() ? "1" : "0"
        ))
    }
}

FastCopyMode_PasteScreenshotQueue(queue) {
    if (!IsObject(queue) || queue.Length < 1)
        return

    clipSave := ""
    try clipSave := ClipboardAll()

    try {
        FastCopyMode_DebugLog("H3", "Shift keys.ahk:FastCopyMode_PasteScreenshotQueue", "paste_enter", Map(
            "queueLen", queue.Length,
            "hasImageAtEnter", FastCopyMode_ClipboardHasImage() ? "1" : "0"
        ))
        for idx, snap in queue {
            try {
                A_Clipboard := snap
                ClipWait 0.6, 1
                snapSize := ""
                try snapSize := snap.Size
                FastCopyMode_DebugLog("H3", "Shift keys.ahk:FastCopyMode_PasteScreenshotQueue", "paste_iter_ready", Map(
                    "idx", idx,
                    "snapType", Type(snap),
                    "snapSize", snapSize,
                    "hasImageNow", FastCopyMode_ClipboardHasImage() ? "1" : "0"
                ))
                Send "^v"
                FastCopyMode_DebugLog("H4", "Shift keys.ahk:FastCopyMode_PasteScreenshotQueue", "paste_iter_sent", Map(
                    "idx", idx
                ))
                Sleep 140
            } catch {
                ; continue to next screenshot
                FastCopyMode_DebugLog("H4", "Shift keys.ahk:FastCopyMode_PasteScreenshotQueue", "paste_iter_failed", Map(
                    "idx", idx
                ))
            }
        }
    } finally {
        if (clipSave != "")
            try A_Clipboard := clipSave
        FastCopyMode_DebugLog("H3", "Shift keys.ahk:FastCopyMode_PasteScreenshotQueue", "paste_exit", Map(
            "restoredClipboard", clipSave != "" ? "1" : "0"
        ))
    }
}

ExecuteSequentialPaste(actionCount) {
    if (!IsInteger(actionCount))
        return
    n := Integer(actionCount)
    if (n < 1)
        return
    Send "!v"
    Sleep 50
    Send "^!b"
    remaining := n - 1
    loop remaining {
        Sleep 300
        Send "^!b"
    }
}

FastCopyMode_IsActive() {
    global gFastCopyModeActive
    return gFastCopyModeActive
}

FastCopyMode_OnCopy() {
    global gFastCopyCount
    gFastCopyCount += 1
    FastCopyModeBanner_Update(gFastCopyCount)
}

FastCopyMode_Start() {
    global gFastCopyModeActive, gFastCopyCount, gFastCopyPasteTargetHwnd, gFastCopyScreenshotQueue
    try {
        gFastCopyPasteTargetHwnd := WinGetID("A")
    } catch {
        gFastCopyPasteTargetHwnd := 0
    }
    gFastCopyCount := 0
    gFastCopyScreenshotQueue := []
    gFastCopyModeActive := true
    try {
        FastCopyModeBanner_Show()
    } catch Error {
        gFastCopyModeActive := false
        ShowCenteredOverlay_Utils("❌ Could not start Fast Copy Mode", 2000, BANNER_ACCENT_ERROR)
    }
}

FastCopyMode_Finish() {
    global gFastCopyModeActive, gFastCopyCount, gFastCopyPasteTargetHwnd
    global gFastCopyScreenshotQueue, gFastCopyLastScreenshotQueue
    count := gFastCopyCount
    shotCount := IsObject(gFastCopyScreenshotQueue) ? gFastCopyScreenshotQueue.Length : 0
    try {
        FastCopyModeBanner_Hide()
    } finally {
        gFastCopyModeActive := false
        gFastCopyCount := 0
    }
    try {
        ; Paste exclusively into the *currently active* window without activating anything else.
        FastCopyMode_ReleaseHotkeyModifiers()
        if (count > 0) {
            FastCopyMode_DebugLog("H5", "Shift keys.ahk:FastCopyMode_Finish", "finish_paste_start", Map(
                "count", count,
                "shotCount", shotCount
            ))
            if (shotCount > 0) {
                FastCopyMode_PasteScreenshotQueue(gFastCopyScreenshotQueue)
                ; Save for hold-to-repeat behavior.
                gFastCopyLastScreenshotQueue := gFastCopyScreenshotQueue.Clone()
            } else {
                gFastCopyLastScreenshotQueue := []
            }

            remaining := count - shotCount
            FastCopyMode_DebugLog("H5", "Shift keys.ahk:FastCopyMode_Finish", "finish_remaining", Map(
                "remaining", remaining
            ))
            if (remaining > 0) {
                ; For non-screenshot copies, fall back to Clip Angel sequential paste.
                ExecuteSequentialPaste(remaining)
            }

            global gFastCopyLastSuccessfulCount
            gFastCopyLastSuccessfulCount := count
        } else
            ShowCenteredOverlay_Utils("⚠ No copies recorded — nothing to paste", 2500, BANNER_ACCENT_INTERMEDIATE)
    } catch Error as e {
        ShowCenteredOverlay_Utils("❌ Fast Copy Mode: " SubStr(e.Message, 1, 80), 2500, BANNER_ACCENT_ERROR)
    }
}

FastCopyMode_RepeatLastPaste() {
    global gFastCopyLastSuccessfulCount
    global gFastCopyLastScreenshotQueue
    if (gFastCopyLastSuccessfulCount < 1) {
        ShowCenteredOverlay_Utils("⚠ No previous Fast Copy paste to repeat", 2500, BANNER_ACCENT_INTERMEDIATE)
        return
    }
    try {
        ; Repeat paste into the *currently active* window without activating anything else.
        FastCopyMode_ReleaseHotkeyModifiers()
        if (IsObject(gFastCopyLastScreenshotQueue) && gFastCopyLastScreenshotQueue.Length > 0) {
            FastCopyMode_PasteScreenshotQueue(gFastCopyLastScreenshotQueue)
            remaining := gFastCopyLastSuccessfulCount - gFastCopyLastScreenshotQueue.Length
            if (remaining > 0)
                ExecuteSequentialPaste(remaining)
        } else {
            ExecuteSequentialPaste(gFastCopyLastSuccessfulCount)
        }
    } catch Error as e {
        ShowCenteredOverlay_Utils("❌ Repeat paste: " SubStr(e.Message, 1, 80), 2500, BANNER_ACCENT_ERROR)
    }
}

#!+j:: {
    global gFastCopyModeActive, FAST_COPY_HOLD_REPEAT_MS
    if (gFastCopyModeActive) {
        FastCopyMode_Finish()
        return
    }
    pressTime := A_TickCount
    KeyWait "j", "T1"
    holdTime := A_TickCount - pressTime
    if (holdTime >= FAST_COPY_HOLD_REPEAT_MS)
        FastCopyMode_RepeatLastPaste()
    else
        FastCopyMode_Start()
}

#HotIf FastCopyMode_IsActive()
~^c:: FastCopyMode_OnCopy()
~PrintScreen:: FastCopyMode_OnCopy()
~!PrintScreen:: {
    FastCopyMode_OnCopy()
    FastCopyMode_CaptureScreenshotToQueue()
}
#HotIf

;-------------------------------------------------------------------
; Environment paths (unchanged)
;-------------------------------------------------------------------
global WORK_SCRIPTS_PATH := "C:\Users\fie7ca\Documents\01 - Scripts"

```
**Section 1: Shortcut system overview**
The AutoHotkey v2 script establishes a dedicated automation layer using the Win+Alt+Shift modifier combination. This prevents conflicts with standard Windows shortcuts. The script manages a global state for the Fast Copy Mode, acting as a bridge between native Windows clipboard events, the user manual inputs, and third-party tools like Clip Angel. It uses a custom debugging module to log events to a JSON file, which is crucial for tracing race conditions.

Risky patterns identified:
1. **Global mutable state**: The script relies on multiple global variables (queue arrays, counters) without atomic locks, which can be disrupted by rapid physical keystrokes.
2. **Timing dependencies**: Hardcoded Sleep commands (40ms, 140ms) govern clipboard access. This is brittle because clipboard I/O speed varies by image size and system load.
3. **Native API contention**: AutoHotkey, the Windows OS, the target application, and Clip Angel all compete for the User32 clipboard API concurrently.
4. **Assumed target window focus**: The script removed the window activation logic to solve the focus jump issue. It now relies entirely on the user not clicking away from the target window. This safely prevents focus stealing but requires the user to maintain focus during the entire flush operation.

**Section 2: Fast Copy Mode flow (state machine)**
1. **Initialization**: Pressing Win+Alt+Shift+J sets the active state, resets global counters and queues, and displays an overlay. It captures the current active window handle but no longer forces activation later.
2. **Capture Phase**: The script hooks standard copy shortcuts. Simple text copies increment a counter. The Alt+PrintScreen combination triggers a custom function that waits for clipboard data, extracts the binary content using the ClipboardAll function, and pushes it to an array.
3. **Finish Phase**: A second Win+Alt+Shift+J press stops the capture state. The script releases modifier keys logically, iterates through the stored screenshot array, injects each binary object back into the clipboard, and sends a paste command. Any remaining non-image items are handed off to Clip Angel via a sequential paste function.

**Section 3: Hypotheses list (with criticality + tests)**

**Hypothesis 1: Missing clipboard clearance before capture.**
* **Criticality**: P0 (Data loss)
* **Mechanism**: AutoHotkey ClipWait function returns instantly if the clipboard already contains any data. Because the script does not empty the clipboard before triggering the screenshot, ClipWait fires prematurely and captures the previous clipboard contents instead of the new screenshot.
* **How to prove**: Check the `capture_pushed` log event.
* **Expected evidence**: The `snapSize` values for rapid captures will be identical or atypically small, and `hasImageBefore` will be true.
* **Suggested fix direction**: Assign an empty string to `A_Clipboard` immediately before the OS capture command.

**Hypothesis 2: Asynchronous paste overwrite (Target application race).**
* **Criticality**: P0 (Data loss)
* **Mechanism**: Modern applications handle paste commands asynchronously. A 140ms delay is often insufficient for an application to open the clipboard, read a large bitmap, and close it. The script overwrites the clipboard with the next image before the target application has finished reading the first one, causing the application to read only the final image multiple times.
* **How to prove**: Increase the paste delay significantly (for example, 2000ms).
* **Expected evidence**: If a 2000ms delay results in all three images pasting correctly, the application read speed is the bottleneck.
* **Suggested fix direction**: Implement a dynamic wait or increase the static sleep time specifically for image payloads.

**Hypothesis 3: Clip Angel clipboard lock collision.**
* **Criticality**: P1 (Frequent workflow break)
* **Mechanism**: Clip Angel constantly monitors the clipboard for changes. When the script pushes a binary object to the clipboard during the paste loop, Clip Angel instantly locks the clipboard to inspect it. The target application attempts to read the clipboard simultaneously and receives an access denied exception.
* **How to prove**: Suspend Clip Angel completely and run the rapid screenshot test.
* **Expected evidence**: The target application successfully pastes all images when Clip Angel is disabled.
* **Suggested fix direction**: Pause Clip Angel monitoring via its API or a dedicated shortcut before initiating the paste loop, and resume it afterward.

**Hypothesis 4: Target application debouncing of rapid inputs.**
* **Criticality**: P1 (Frequent workflow break)
* **Mechanism**: Applications like Microsoft Teams or Slack contain UI thread protections that debounce rapid consecutive paste commands to prevent freezing. The script sends three paste commands within 500ms, which the application interprets as a single user action.
* **How to prove**: Attempt the rapid paste in a simple application like Microsoft Paint or WordPad.
* **Expected evidence**: WordPad pastes all three images, whereas Teams drops the first two.
* **Suggested fix direction**: Adjust the paste pacing based on the specific target application.

**Hypothesis 5: Incomplete OS capture due to rapid polling.**
* **Criticality**: P2 (Edge case)
* **Mechanism**: The Windows Alt+PrintScreen function requires time to render the active window. Pressing it three times rapidly interrupts the internal OS capture thread, resulting in only the final capture successfully writing a bitmap to the clipboard.
* **How to prove**: Review the `capture_enter` and `capture_pushed` timestamps in the debug logs.
* **Expected evidence**: The timestamps are less than 50ms apart, and `hasImageAfter` is false for the intermediate captures.
* **Suggested fix direction**: Add a KeyWait to prevent consecutive captures faster than the OS can render them.

**Hypothesis 6: AutoHotkey memory allocation failure for rapid captures.**
* **Criticality**: P2 (Edge case)
* **Mechanism**: The ClipboardAll function allocates memory for the entire clipboard binary. Rapid consecutive calls might overlap in the AutoHotkey runtime, causing silent allocation failures or corrupted object references in the array.
* **How to prove**: Output the variable type and size of each queue item in the debug log.
* **Expected evidence**: Intermediate queue items register as empty strings or have a size of zero.
* **Suggested fix direction**: Clone the clipboard object explicitly or use Windows API direct memory copying.

**Hypothesis 7: Modifier key state leak during paste.**
* **Criticality**: P2 (Edge case)
* **Mechanism**: The user is holding Alt and Shift while taking screenshots. During the paste loop, if the physical keys are still pressed or bouncing, the paste command might be intercepted by the OS as Ctrl+Shift+V or Ctrl+Alt+V.
* **How to prove**: Use the AutoHotkey Key History tool during the paste operation.
* **Expected evidence**: The log shows modifier key down events interleaving with the V key down events.
* **Suggested fix direction**: Enforce a strict KeyWait for all modifiers before executing the finish function.

**Hypothesis 8: ClipWait fails to block during paste loop.**
* **Criticality**: P3 (Nice to have)
* **Mechanism**: In the paste loop, setting the clipboard is followed immediately by ClipWait. Since AutoHotkey just populated the clipboard, ClipWait returns instantly. It provides no actual confirmation that the OS has finished registering the clipboard format, rendering the wait useless.
* **How to prove**: Measure the execution time of ClipWait during the paste loop.
* **Expected evidence**: ClipWait consistently takes 0ms to execute.
* **Suggested fix direction**: Use DllCall to verify the specific clipboard format availability instead of relying on the generic ClipWait.

**Section 4: Recommended next experiments (top 3)**
1. **The Clipboard Clear Test**: Modify the capture function to include `A_Clipboard := ""` directly above the Sleep 40 line. Attempt three rapid screenshots and observe if the queue captures distinct sizes.
2. **The Artificial Delay Test**: In the paste function, change the Sleep 140 to Sleep 1500. Perform the rapid screenshot workflow and paste into the primary target application. If all three images appear, the root cause is confirmed as a timing race condition with the application.
3. **The Clip Angel Isolation Test**: Completely close Clip Angel. Perform the standard workflow with the original 140ms sleep. If the issue disappears, Clip Angel clipboard locking is the primary interference factor.

**Section 5: Likely fix candidates (ranked)**
1. **Clear clipboard before capture**: This is the most critical programmatic flaw in the current script. ClipWait cannot function as intended without an initially empty clipboard.
2. **Dynamic paste verification**: Instead of a static sleep, implement a check to ensure the target application has processed the paste. Since true verification is difficult, increasing the sleep time for image formats while keeping text paste fast is a pragmatic compromise.
3. **Clip Angel pause integration**: Send Clip Angel suspend monitoring shortcut at the beginning of the paste loop and resume it in the finally block to eliminate lock contention.

**Testing Plan**

**Multiple monitors focus stability**
Step 1: Open a target application on Monitor 1.
Step 2: Open a reference document on Monitor 2.
Step 3: Click inside the target application to ensure it is the active window.
Step 4: Press Win+Alt+Shift+J.
Step 5: Move the mouse to Monitor 2 and press Alt+PrintScreen. Do not click Monitor 2.
Step 6: Press Win+Alt+Shift+J again.
Success Criteria: The image pastes directly into the application on Monitor 1. The focus does not jump to Monitor 2.

**Rapid screenshot capture rates**
Step 1: Start Fast Copy Mode.
Step 2: Perform three Alt+PrintScreen captures with a two second pause between each.
Step 3: Finish and verify paste.
Step 4: Start Fast Copy Mode.
Step 5: Perform three captures as rapidly as physically possible.
Step 6: Finish and verify paste.
Success Criteria: Both the slow and fayst tests result in three distinct images pasted into the target window.

**Different paste targets**
Step 1: Open Microsoft Word, Slack, and an integrated development environment.
Step 2: Capture three screenshots.
Step 3: Focus Microsoft Word and finish Fast Copy Mode. Verify results.
Step 4: Repeat the process for Slack and the environment.
Success Criteria: All applications correctly receive three distinct images without dropping any payloads.

**Regression checks**
Step 1: Start Fast Copy Mode.
Step 2: Highlight text and press Ctrl+C three times on different sentences.
Step 3: Finish Fast Copy Mode in a text editor.
Success Criteria: Clip Angel successfully processes the sequential text pastes without overlapping or losing data.