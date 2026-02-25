# Graceful Error Handling – Proposed Locations

This document lists specific locations in the codebase where graceful error handling can be implemented, following the pattern used for the Outlook Reminder hotkey: check for window existence (or validity) before activation, and show a user-facing message (e.g. "Error: Target window not found.") instead of letting AHK throw or show a call-stack dialog.

**Reference pattern:** Guard `WinActivate` with `WinExist` (or equivalent); if not found, show a centered overlay or notification and return. Reusable helper: `TryActivateWindow_WM(winSpec, errorMessage)` in [WindowManagement.ahk](WindowManagement.ahk) (already implemented).

**Note:** The file `outlook-tree.ahk` was listed in the context but not found under the scripts directory in this analysis; it is omitted below. If it exists elsewhere, it can be added in a future pass.

---

## 1. Outlook.ahk

**Status:** Reminder hotkey already fixed (lines 91–95). Other candidates:

| Location | Current behavior | Proposal |
|----------|------------------|----------|
| **ActivateOutlookMailbox** (lines 21–24) | `WinActivate(hwnd)` after finding a matching window in a loop. The `hwnd` is from `WinGetList` so it should exist; if the window closes between iteration and activate, AHK could throw. | Wrap in `try`/`catch`; on failure show `ShowCenteredOverlay_Utils("Error: Target window not found.", 2000)` and return or rethrow as needed. |
| **ActivateOutlookCalendar** (lines 34–35) | `WinExist("Calendar - Eduardo")` then `WinActivate "Calendar - Eduardo"`. Window could close between check and activate. | Optional: wrap `WinActivate` in try/catch and show same overlay on failure. Low priority since caller already shows "Calendar and Mailbox are not open" when both fail. |

---

## 2. Shift keys.ahk

| Location | Current behavior | Proposal |
|----------|------------------|----------|
| **ActivateReminder()** (lines 2565–2568) | `WinActivate("ahk_exe OUTLOOK.EXE")` with no check. If Outlook is not running, this can throw. | Before `WinActivate`, check `if !WinExist("ahk_exe OUTLOOK.EXE")` and show a graceful message (e.g. via `ShowCenteredOverlay_Utils` or equivalent), then return. |
| **gCommitPushTargetWin** (lines 9145–9147) | `WinActivate gCommitPushTargetWin` then `WinWaitActive("ahk_id " gCommitPushTargetWin, , 2)`. The stored HWND can be stale (window closed). | Before activate: `if !WinExist("ahk_id " gCommitPushTargetWin)` show "Target window not found" (or "Cursor window closed") and skip activate; else proceed. |
| **gEmojiTargetWin** (lines 9254–9256) | `WinActivate gEmojiTargetWin` when inserting emoji. Same stale-HWND risk. | Same as above: check `WinExist("ahk_id " gEmojiTargetWin)`; if not found, show short message and skip activate (optionally still send emoji to active window or skip). |
| **WinActivate("ahk_id " normalMeetingHwnd)** (lines 2774, 2781) | Activation of a previously found meeting window. HWND could become invalid. | Wrap in try/catch or check `WinExist("ahk_id " normalMeetingHwnd)` before activate; show graceful message on failure. |
| **WinActivate("ahk_id " targetHwnd)** in appointment/meeting flows (e.g. 5818, 5895) | Target HWND from earlier in the function; can be stale. | Guard with `WinExist("ahk_id " targetHwnd)`; if not found, show "Target window not found" and return. |
| **WinActivate("ahk_id " chatGPTHwnd)** (line 6348) | HWND captured earlier; window may have closed. | Check `WinExist("ahk_id " chatGPTHwnd)` before activate; show message if not found. |
| **WinActivate("ahk_exe Spotify.exe")** (line 10648) | No guard; Spotify might not be running. | If `!WinExist("ahk_exe Spotify.exe")`, show "Target window not found" (or "Spotify not running") and return. |
| **WinActivate("ahk_id " browserHwnd)** (line 12760) | Browser HWND from earlier; can be stale. | Guard with `WinExist("ahk_id " browserHwnd)` and graceful message. |
| **WinActivate("ahk_id " geminiHwnd)** / **WinActivate("ahk_exe chrome.exe")** (lines 12944, 12950) | Fallback activation; either can fail. | After failure of first, if second also fails or window missing, show "Target window not found." |
| **WinActivate("ahk_id " hwnd)** in file/save dialog flows (e.g. 8695, 8758, 8857) | Dialog HWND can disappear (user closed dialog). | Already wrapped in `try` in some places; ensure all such activations show a short user message in catch (e.g. "Dialog closed."). |

---

## 3. AppLaunchers.ahk

| Location | Current behavior | Proposal |
|----------|------------------|----------|
| **Open Desktop** hotkey (lines 179–186) | After a loop, `targetHwnd` may still be 0 if Explorer never appeared. Code does `if (targetHwnd)` then `WinActivate("ahk_id " targetHwnd)`. If the loop exits without setting `targetHwnd` (e.g. timeout), the block is skipped—so 0 is not passed. But if `targetHwnd` is set and the window closes before activate, AHK could throw. | Optional: inside the `if (targetHwnd)` block, before `WinActivate`, check `WinExist("ahk_id " targetHwnd)`; if false, show "Desktop window not found" and return. |
| **Open Chrome** (lines 210–214) | `Run "chrome.exe"`, `WinWait(..., 10)`, then `WinActivate("ahk_exe chrome.exe")`. If WinWait times out, no window exists and WinActivate can throw. | If `!WinExist("ahk_exe chrome.exe")` after WinWait, show "Chrome did not start in time" (or "Target window not found") and return; else activate. |
| **WhatsApp** (lines 246–247) | Already guarded: `if WinExist("WhatsApp")` then `WinActivate("WhatsApp")`. | Optional: wrap WinActivate in try/catch and show message on failure (window could close between check and activate). |
| **Wikipedia** (lines 364–372, 756–758, 1300–1304) | Various Wikipedia activations; some guarded by WinExist, some by WinWait. | After any `WinWait("Wikipedia", ...)` that times out, do not call WinActivate; show "Target window not found" or "Wikipedia window did not open in time." For 372, 379: optional try/catch around WinActivate with message. |
| **Cursor** (lines 305–306, 311–312) | `if WinExist("ahk_exe Cursor.exe")` then `WinActivate`; else Run and `WinWaitActive`. If Run fails or window never appears, WinWaitActive can timeout without a friendly message. | After Run, if `!WinExist("ahk_exe Cursor.exe")` within timeout, show "Cursor did not start" or "Target window not found." |
| **WinActivate(targetWindow)** / **WinActivate(fallbackWindow)** (lines 86, 94) | `targetWindow`/`fallbackWindow` are window specs from earlier logic; they could be invalid if the window closed. | Before each activate, check with `WinExist(targetWindow)` / `WinExist(fallbackWindow)`; if false, show "Target window not found" and skip (or show fallback panel for Cursor). |

---

## 4. Microsoft Teams.ahk

| Location | Current behavior | Proposal |
|----------|------------------|----------|
| **ActivateWindowWithRetry** (lines 35, 57, 68) | `WinActivate(hwnd)` with hwnd from caller. If caller passes an invalid or closed window, AHK can throw. | At start of function: `if !WinExist("ahk_id " hwnd)` return false (and optionally show "Target window not found"). Wrap each WinActivate in try/catch; on failure, show overlay and continue retry or return false. |
| **ActivateTeamsMeetingWindow** (lines 104–106, 164–165) | After `WinExist("RegEx)...")` or from loop, `WinActivate(hwnd)`. Window could close before activate. | Optional: try/catch around WinActivate; on failure show same style overlay as existing "MEETING WINDOW FOUND BUT COULD NOT ACTIVATE". |
| **ActivateTeamsChatWindow** (lines 159, 165) | `WinActivate(hwnd)` after getting hwnd from list or WinExist. | Optional: guard with `WinExist("ahk_id " hwnd)` or try/catch; show "Target window not found" if activation fails. |
| **WinActivate(teamsWindow)** (line 670) | `teamsWindow` from `WinGetList`; could be invalid if no Teams window. Preceding check is `if !WinExist("ahk_exe ms-teams.exe") && !WinExist("ahk_exe Teams.exe")` then return—so 670 is only reached when a window exists, but it could close before 670. | Optional: before 670, confirm `WinExist("ahk_id " teamsWindow)`; if not, show "Target window not found" and return. |

---

## 5. Gemini.ahk

| Location | Current behavior | Proposal |
|----------|------------------|----------|
| **GeminiTriggerReadAloud** (lines 253–256) | `if hwnd := GetGeminiWindowHwnd()` then `WinActivate("ahk_id " hwnd)`. If GetGeminiWindowHwnd returns 0, the block is skipped. If it returns a valid hwnd that becomes invalid before activate, AHK could throw. | Optional: wrap `WinActivate` in try/catch; on failure show "Gemini window not found" or "Target window not found" and return. |
| **#!+i Open Gemini** (lines 763–765) | Same pattern: hwnd from GetGeminiWindowHwnd, then WinActivate. | Same as above. |
| **Restore focus to origHwnd** (lines 909–915, 1120–1126) | `WinExist("ahk_id " origHwnd)` then `WinActivate`; else unconditional `WinActivate`. The unconditional path can throw if window is gone. | In the else branch, only call WinActivate if `WinExist("ahk_id " origHwnd)`; otherwise skip or show short message. |
| Other **WinActivate("ahk_id " ...)** (e.g. 254, 561, 724, 872, 988, 1007, 1087) | Various; most are after a successful GetGeminiWindowHwnd or WinExist("A"). | Prefer consistent pattern: check window exists before activate; on failure show "Target window not found" or context-specific message. |

---

## 6. GeminiToCursorBridge.ahk

| Location | Current behavior | Proposal |
|----------|------------------|----------|
| **Bridge_ActivateCursorProject** (lines 463–466, 474, 495, 537, 550) | Multiple `WinActivate("ahk_id " targetHwnd)` inside try/catch; returns 0 on failure. Callers (e.g. WindowManagement.ahk) already show "Failed to open project or focus AI field" for `cursor_activate_failed`. | No change required for user-facing message. Optional: in catch blocks, log or set a more specific reason so callers can show "Target window not found" vs "Activation failed" if desired. |
| **Activate Gemini window** (lines 264–274) | try/catch around WinActivate; returns `gemini_activate_failed`; caller shows "Could not activate Gemini window" or similar. | Already graceful. Optional: ensure caller message is "Target window not found" when the failure is specifically "window no longer exists." |

---

## 7. Spotify.ahk

| Location | Current behavior | Proposal |
|----------|------------------|----------|
| **ActivateSpotify()** (lines 136–141) | `hwnd := WinExist(winTitle)` then `WinActivate(winTitle)`. If Spotify is not running, `hwnd` is 0 but `WinActivate("ahk_exe Spotify.exe")` is still called and can throw. | If `!hwnd`, show a short message (e.g. "Spotify not running" or "Target window not found") and return before calling WinActivate. |
| **FocusYouTube()** (lines 155–156) | Iterates windows and calls `WinActivate(win)` then `WinWaitActive`. The `win` is from WinGetList so it exists at call time; low risk. | Optional: wrap in try/catch and return false with message if activation fails. |

---

## 8. Utils.ahk

| Location | Current behavior | Proposal |
|----------|------------------|----------|
| **ClipAngel – Merge Clips** (lines 969–970) | After `if !WinExist("ClipAngel")` show MsgBox and return; else `WinActivate("ClipAngel")`. Window could close between check and activate. | Wrap WinActivate in try/catch; on failure show "ClipAngel window not found" (or same overlay style as other Utils messages) and return. |
| **ActivateClipAngelWithFocusCorrection** (lines 1164–1176) | If `WinExist("ClipAngel")` then activate; else Send !v and `WinWait("ClipAngel", , 10)`. If WinWait fails, return; else `WinActivate("ClipAngel")`. Race possible. | Wrap both WinActivate calls in try/catch; on failure show "ClipAngel window not found" and return. |
| **Handy / Recording** (lines 802–803, 816–817) | `WinExist("Handy ahk_class Tauri Window")` then `WinActivate`; later `WinWaitActive(..., 2)`. | If WinWaitActive fails, show "Target window not found" or "Handy window did not activate" instead of continuing. |
| **Peek** (lines 5124–5127, 5351–5357) | `hwnd := WinExist("Peek")` or WinExist( exe ); if !hwnd show message; else WinShow and WinActivate. | Already guarded. Optional: try/catch around WinActivate and show message on failure. |
| **Outlook activation** (lines 2132–2134) | `WinActivate("ahk_exe OUTLOOK.EXE")` and WinWaitActive. Outlook might not be running. | If `!WinExist("ahk_exe OUTLOOK.EXE")` before activate, show "Outlook not running" or "Target window not found" and return/skip. |
| **Focus restore (g_HotstringGeminiRestoreHwnd)** (lines 6092–6094) | If `g_HotstringGeminiRestoreHwnd && WinExist("ahk_id " ...)` then WinActivate. Good. | No change. |
| **Other WinActivate("ahk_id " . targetHwnd)** (e.g. 3112, 4615, 5271, 5357, 5682, 5729, 5779, 5810, 5818, 5895, 5954, 5958, 6228, 6232) | Various; some already in try/catch. | Where not guarded: add `WinExist("ahk_id " targetHwnd)` check or try/catch and show "Target window not found" on failure. |

---

## 9. WindowManagement.ahk

| Location | Current behavior | Proposal |
|----------|------------------|----------|
| **TryActivateWindow_WM** | Already implemented; use for new code. | N/A. |
| **CycleWindowsOnMonitor** (lines 538–541) | `try WinActivate "ahk_id " target.hwnd` then `catch { return }`. User gets no message. | In catch, call `ShowNotification_WM("Error: Target window not found.")` (or "Window no longer available") then return. |
| **MinimizeWindowOnMonitor** / **CloseWindowOnMonitor** (lines 668–675, 701–708) | Already have try/catch and `ShowNotification_WM("Failed to minimize/close window...")`. | Consider unifying message to "Target window not found" when the failure is specifically window-not-found (e.g. check in catch or use TryActivateWindow_WM for the activate step). |
| Other **WinActivate("ahk_id " ...)** (e.g. in Cursor/project/preview flows) | Many use hwnd from WinGetList or FindAndActivateCursorWindow; some have try/catch. | Where only `catch { return }` exists with no user message, add `ShowNotification_WM("Error: Target window not found.")` in the catch block. |

---

## 10. Summary by priority

- **High (user-facing errors today):** Shift keys.ahk ActivateReminder (2566); Shift keys.ahk gCommitPushTargetWin / gEmojiTargetWin; Spotify.ahk ActivateSpotify when not running; AppLaunchers.ahk Chrome WinWait timeout then WinActivate.
- **Medium (stale HWND or race):** Shift keys.ahk other stored-HWND activations; Utils.ahk ClipAngel and Outlook; Microsoft Teams.ahk ActivateWindowWithRetry and ActivateTeamsChatWindow; AppLaunchers.ahk Cursor/Desktop/Wikipedia.
- **Low (already guarded or rare):** Outlook.ahk Calendar/Mailbox; Gemini.ahk and GeminiToCursorBridge.ahk (already return false or show message in caller); WindowManagement.ahk CycleWindowsOnMonitor (add message in catch).

Implementations can use the existing `TryActivateWindow_WM(winSpec, errorMessage)` in WindowManagement.ahk where scripts include it, or the same pattern (WinExist check + overlay/notification) in scripts that use Utils (ShowCenteredOverlay_Utils) or their own banner.
