Fix 3: Enforce clipOk Failure Path (Addresses P5) In Utils.ahk (DoCopyCore), if clipOk evaluates to false, halt the flow, show an error banner (ShowCenteredOverlay_Utils), and explicitly prevent the downstream Cursor transfer from triggering on an empty clipboard.

Fix 4: Align the Bridge (Addresses P3) Update GeminiToCursorBridge.ahk to replace PostMessage and the Sleep 150 loop with the exact same SendMessage(..., 20000) pattern.
copyFromBridge(wParam, lParam, msg, hwnd) {
    geminiHwnd := lParam
    ; ...
    r := CopyLastGeminiMessageToClipboard({ restoreWindow: false, playChimeAndNotify: false, alreadyActive: true }, geminiHwnd)
    ; ...
}

Fix 2: Honor Pre-Activated State in Receiver In Gemini.ahk, update copyFromBridge to accept the HWND and instruct the copy function to skip redundant window activation:


; SendMessage(Msg, wParam, lParam, Control, WinTitle, WinText, ExcludeTitle, ExcludeText, Timeout)
SendMessage(WM_COPY_LAST_GEMINI, 0, this.GeminiHwnd, , "ahk_id " targetHwnd, , , , 20000)

1. Ordered Hypothesis List
Priority 1: SendMessage Default Timeout Race (H-B)

Why it's highly likely: AutoHotkey v2's SendMessage has a hardcoded default timeout of 5000ms. The copyFromBridge execution involves UIA initialization, DOM scanning, ^{End} scrolling, and a 2-second ClipWait. If Chrome is sluggish or the DOM is large, this easily exceeds 5 seconds. When the timeout hits, SendMessage throws an exception, Utils.ahk catches it, and immediately restores focus to the original window. Meanwhile, Gemini.ahk continues running the copy routine in the background. This completely undermines the synchronous block and re-introduces the exact focus-stealing race condition we intended to fix.
Priority 2: Disconnected Window Target / Redundant Activation (H-C / H-D)

Why it's likely: Utils.ahk correctly guarantees focus on this.GeminiHwnd before sending the IPC message. However, copyFromBridge in Gemini.ahk calls CopyLastGeminiMessageToClipboard without passing the specific geminiHwnd and without alreadyActive: true. This forces Gemini.ahk to call GetGeminiWindowHwnd() (guessing the window again) and execute a redundant WinActivate/WinWaitActive sequence. If multiple Chrome windows exist, it might target the wrong one or incur unnecessary delays.
Priority 3: Legacy PostMessage Remains in Alternative Paths (H-G)

Why it's likely: GeminiToCursorBridge.ahk (which handles direct transfers) still uses the asynchronous PostMessage pattern combined with a brittle polling loop (Bridge_CopyGeminiLastMessageToClipboard). If the user is triggering the transfer via a hotkey linked to this bridge rather than D2C_FlowManager, they are running the old, unfixed race condition.
Priority 4: UIA Element Timing Failure (H-C)

Why it's likely: The ^{End} keystroke takes time to physically scroll the DOM. The fixed 350ms sleep (GEMINI_SCROLL_SETTLE_MS) might not be long enough for the "Copy" button to render its bounding rectangle on-screen. If GetLastGeminiCopyButton fails to find the UIA element, it silently fails (r=0).
Priority 5: clipOk Validation is Ignored (H-F)

Why it's likely: Even if the copy fails (empty clipboard), DoCopyCore computes clipOk = false but takes no action. It still plays the success chime and restores focus, creating a confusing user experience where the automation appears to succeed but the clipboard is empty.
2. Investigation Plan
If the user provided the .cursor\d2c_copy_debug.log, parse it for the following evidence. If not, ask them to reproduce the issue using Win+Alt+Shift+0 and provide the log.

Step 1: Verify Timeout (Priority 1)

Read: d2c_copy_debug.log
Look for: DoCopyCore SendMessage THREW: TimeoutError (or similar).
Look for: copyFromBridge EXIT r=1 ms=.... If ms is near or above 5000, the hypothesis is confirmed. SendMessage gave up before Gemini.ahk finished.
Step 2: Verify Window Resolution (Priority 2 & 4)

Read: d2c_copy_debug.log
Look for: copyFromBridge EXIT r=0. This means CopyLastGeminiMessageToClipboard failed entirely.
Check: gemini_copy_result.txt. If it contains 0, UIA failed to find the button or ClipWait failed.
Step 3: Verify Entry Path (Priority 3)

Ask the user: "Are you triggering the failure exclusively from the dictation completion banner (Y / C), or are you pressing a direct shortcut like Win+Alt+Shift+C?" If the latter, GeminiToCursorBridge.ahk is the culprit.
3. Recommended Fix Direction
Do not speculate or rewrite the UIA logic yet. Apply these deterministic boundary fixes to guarantee IPC synchronization:

Fix 1: Override Default Timeout & Pass HWND Context (Addresses P1 & P2) In Utils.ahk (DoCopyCore), pass this.GeminiHwnd via lParam and increase the SendMessage timeout to 20 seconds (20000ms):

---

## Implementation status (debug iteration)

Applied in repo:

- **P1 + P2**: `Utils.ahk` `DoCopyCore` — `SendMessage(..., this.GeminiHwnd, ..., 20000)`; `Gemini.ahk` `copyFromBridge` uses `lParam` as browser hwnd and sets `alreadyActive: true` when that window is active.
- **P3**: `GeminiToCursorBridge.ahk` — replaced `PostMessage` + sleep polling with the same `SendMessage(..., geminiBrowserHwnd, ..., 20000)` pattern.
- **P5**: `DoCopyCore` — `clipOk` / `sendOk` gate overlays and copy chime (no success chime on failure).

**P4** (scroll/UIA settle) not changed pending evidence from `d2c_copy_debug.log` (`copyFromBridge EXIT r=0` with `sendOk=1`).

**Follow-up (log 2026-03-19):** `SendMessage THREW: Target window not found` while `WinGetTitle` on the same `ahk_id` succeeded — **AutoHotkey’s main window is hidden**, so **`SendMessage` required `DetectHiddenWindows true`**. Fixed in `Utils.ahk` (`DoCopyCore`, `#!+L`, read-aloud `PostMessage`).