# AIB Allow Working Baseline

Date: 2026-05-04
Status: Working baseline confirmed

## What is working now
- Allow watcher handles the confirmation flow in the target VS Code window.
- Cross-window focus juggling is eliminated.
- Detection is constrained to chat confirmation framing.
- Runtime debug logging is available for follow-up diagnostics.

## Trigger coverage (implemented)
1. Chat submit trigger
- Implemented in Utils submit path.
- Watcher starts after send action or Enter fallback send.

2. Start implementation trigger
- Implemented via probe that watches Start implementation transition.
- Watcher starts pinned to the detected window.

## Validation checklist
1. Chat submit / Enter path
- Open VS Code chat.
- Send a prompt using your normal Enter-send behavior.
- Expect watcher to start and handle Allow prompt.

2. Start implementation path
- In the AI flow, click Start implementation.
- Expect watcher to start and handle Allow prompt.

3. Expected pass criteria
- No window hopping.
- Allow prompt accepted.
- Session continues or exits according to watcher state.

## If something regresses
- Check runtime trace file: docs/aib-allow-runtime-debug.log
- Review previous deep-dive notes: docs/aib-allow-button-debug-followup.md
