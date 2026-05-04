# AIB Allow Button Debug Follow-up

Date: 2026-05-04
Owner: Copilot + user

## Current symptom
- Window juggling is fixed.
- The watcher detects the chat confirmation state, but the Allow Once button is still not being accepted in the target VS Code window.

## Confirmed UI evidence
- Accessible tree shows the target button exists:
  - Type 50000 Button
  - Name Allow Once (Ctrl+Enter)
  - Class monaco-button small monaco-text-button
- Button is inside container class chat-confirmation-widget-container.
- Confirmation text includes accept hint Control+Enter to accept.

## What has been implemented so far
1. Watcher architecture and triggers
- Added timer-based watcher with global timeout, per-window state, and decision flow.
- Wired trigger from VSCode_SubmitChat send path.
- Added Start implementation probe path.

2. Decision UX flow
- First banner 2 seconds.
- Second banner 3 seconds for Y or N.
- No input on second banner is treated as Yes.
- Global timeout set to 8 minutes.

3. Cross-window behavior and scope
- Added scope filtering for ide, vscode, and cursor.
- Added pinned hwnd support so one watcher session can be locked to one window.
- Rapid-fire hotkey now pins to active VS Code window.

4. Click logic hardening
- Target preference:
  - Find button in chat-confirmation-widget-container first.
  - Then fallback to best scored allow button candidate.
- Added verify-after-action checks to confirm prompt is gone.
- Added keyboard fallback Ctrl+Enter.
- Added extra ControlSend routing to Chromium render controls.
- Added non-activating coordinate ControlClick fallback using element bounds.

5. Focus safety
- Removed forced WinActivate calls from allow click flow.
- Removed forced activation from keyboard fallback path.

## Why it may still fail (ranked hypotheses)
1. Background input restriction in Chromium surface
- VS Code may ignore background ControlSend and non-foreground synthetic clicks for this widget in this specific state.

2. Coordinate mismatch during render transitions
- The UIA element bounds may be stale or transformed at click time, causing background coordinate click to miss.

3. Control target mismatch
- Sending to top-level or guessed render host hwnd may not map to the active chat webview input pipeline.

4. Timing race
- The button appears, but click attempt may occur before widget is fully interactive.

5. Silent exception path not visible
- Action branch may be failing without enough runtime trace detail.

## Most likely root cause right now
- The button can be found by UIA, but action delivery to that specific Chromium-hosted button is blocked or missed when done fully in background mode.

## Proposed next debug pass
1. Add high-signal trace line per attempt
- Log selected hwnd, selected button name and class, bounds, offscreen flag.
- Log which action branch fired:
  - invoke
  - click
  - coord-click
  - ctrl-enter-window
  - ctrl-enter-render-host
- Log post-check result (allow still present or cleared).

2. Add one bounded foreground fallback mode for diagnosis only
- If all background branches fail, temporarily activate pinned hwnd, then Send Ctrl+Enter once, then restore previous focus.
- Keep this fallback behind a debug flag so normal behavior remains non-juggling.

3. Add short re-try burst per detection
- Two or three attempts over 400 to 700 ms while prompt is visible.
- Re-read button each attempt to avoid stale bounds.

4. Capture one successful and one failed trace sample
- Compare branch behavior and confirm whether foreground fallback is the only reliable path.

## Minimal instrumentation fields
- tick
- hwnd
- window title
- found_confirmation_container true false
- found_allow_button true false
- button_name
- button_class
- button_bounds
- button_offscreen
- action_branch
- action_ok true false
- verify_allow_still_present true false

## Suggested implementation toggle names
- g_AIB_AllowWatcherDebugTraceEnabled
- g_AIB_AllowWatcherForegroundFallbackEnabled
- g_AIB_AllowWatcherRetryBurstCount

## Exit criteria
- In a pinned VS Code window, Allow Once is accepted reliably in at least 10 consecutive prompts without cross-window focus juggling.
