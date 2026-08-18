# Click Sequence vocabulary

Use these terms in the manager UI, the HTML map, and future agent work. Do not invent synonyms.

## Shortcut / Macro

Trigger container for one hotkey. Example: `#!+9` (Win+Alt+Shift+9) named **AI Quick Download**. v1 does not bind new hotkeys at runtime.

## Context Rule

Per-Shortcut behavior that applies to every Click unless a Click overrides it.

- `searchOrder=bottomUp` — pick the matching control lowest on screen (chat feeds). Default for AI Quick Download.
- `searchOrder=topDown` — pick the matching control highest on screen.
- `searchOrder=firstMatch` — first match in UIA tree order.

## Slot

One ordered step in the Shortcut chain. **Every Slot must succeed.** Types:

- **Hardcoded Script** — named function from the script registry.
- **Sequence Group** — ordered Sibling Sequences.

## Hardcoded Script

A registered AutoHotkey function. Users pick an id; they cannot author new AHK in the GUI.

| id | Effect |
|----|--------|
| `scrollFeedBottom` | Scroll the companion chat feed to the bottom |
| `desktopWait` | Wait for a new file on the Desktop |
| `desktopCut` | Cut the newest Desktop item (skipped for finance import) |
| `focusCompanion` | Focus the AI companion window (#!+9 also does this before the chain) |

## Sequence Group

A Slot that holds Sibling Sequences. Sequences in the group are still filtered by Context at run time (`gemini` / `enterprise` / `copilot` / `*`).

## Sibling Sequence

An alternative Dynamic Sequence in the same Sequence Group. Tried in order. If any Click in a Sibling fails, that Sibling is aborted and the next Sibling starts from its first Click. There is no undo of Clicks already invoked.

## Click

One ordered target inside a Sequence. All Clicks in a Sibling must succeed. Fields: control type, optional ClassContains AND-filter, settle ms, prefer-newest override.

## Alias

Fallback identifier for one Click, tried in order until one hits. Primary kind is **name** (UIA `Name` string). Also: `automationId`, `classContains`, `region` (window-relative click), `icon` (stored, not executed).

Example: `["concluded", "finished", "terminated"]` as name Aliases on one Click.

## Context

Companion filter on a Sibling Sequence: `gemini`, `enterprise`, `copilot`, or `*` (any).
