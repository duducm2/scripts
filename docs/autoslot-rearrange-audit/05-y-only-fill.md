# Y-only background fill (policy)

**Date:** 2026-07-22

## Rule

**Only** `Ctrl+Alt+Win+Y` (and Window tools menu `[3]`, same handler) may promote background windows into AutoSlot free capacity.

Close, minimize, move, maximize-on-monitor, and post-swap quiet must **not** import backgrounds. Companion **heal** (maximize leftover half after close/minimize) stays automatic and silent.

## Related UX fixes

| Item                       | Change                                         |
| -------------------------- | ---------------------------------------------- |
| Place vs auto-fill         | Solved by removing auto import                 |
| Maximize half → companion  | Pair unregisters; companion not auto-maximized |
| Heal silent / fill toasted | Fill toasts only on Y; heal silent             |
| Swap `[F]` replace-skip    | Removed                                        |

## Hotkey

[`WindowManagement/hotkeys.ahk`](../../WindowManagement/hotkeys.ahk): `^!#y:: WM_WindowTools_OnTileBackground()` — sole user gesture for background fill.
