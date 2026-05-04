# AIB Enter Trigger Debug Guide

Date: 2026-05-04

## Goal
Validate that pressing plain Enter in the AI textfield starts the Allow watcher flow.

## Instrumentation Added
- AppLaunchers now intercepts plain Enter only when active app is VS Code or Cursor.
- For VS Code, when chat input is focused, Enter routes through VSCode_SubmitChat.
- For Cursor, when composer is focused, Enter is sent and watcher is armed in cursor scope.

## Runtime Log Signals
Read from docs/aib-allow-runtime-debug.log.

### Enter detected
- `enter-trigger detected proc=code hwnd=...`
- `enter-trigger detected proc=cursor hwnd=...`

### Watcher started
- `session-start source=chat_submit scope=vscode pinned=...`
- `session-start source=chat_submit_cursor scope=cursor pinned=...`

### Banner emitted
- `trigger-banner source=chat_submit text=...`
- `trigger-banner source=chat_submit_cursor text=...`

## Manual Test Procedure
1. Open VS Code chat input and type a small prompt.
2. Press Enter once.
3. Confirm start banner appears immediately.
4. Check log for `enter-trigger detected proc=code` and `session-start source=chat_submit`.
5. Repeat in Cursor composer and confirm `chat_submit_cursor` source.

## Failure Clues
- If Enter sends but no watcher session-start appears:
  - likely chat focus predicate failed.
- If no `enter-trigger detected` line appears:
  - Enter was not in Code/Cursor foreground scope.
- If `enter-trigger detected` appears but no banner:
  - banner debounce may have suppressed a duplicate event in < 900ms window.
