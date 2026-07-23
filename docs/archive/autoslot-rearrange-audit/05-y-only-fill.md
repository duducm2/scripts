# Y-only background fill (policy)

**Date:** 2026-07-22  
**Superseded for full Y semantics by:** [`docs/canon/windows-rearrange.md`](../../canon/windows-rearrange.md)

## Rule

**Only** `Ctrl+Alt+Win+Y` (and Window tools menu `[3]`, same handler) may promote background windows into AutoSlot free capacity **or** expand a lone half-window into a free slot on the same monitor (maximize both slots). Close, minimize, move, maximize-on-monitor, and post-swap quiet must **not** import backgrounds.

Companion **heal** after close/minimize (maximize leftover half) stays automatic and silent.

On **Y** specifically:

| Monitor state | Y action |
| --- | --- |
| Empty | Import from background |
| Lone half + free slot | Maximize that window (no BG import into the free half) |
| Lone maximized | 50/50 SnapPair with a background window |

## Related UX fixes

| Item                       | Change                                         |
| -------------------------- | ---------------------------------------------- |
| Place vs auto-fill         | Solved by removing auto import                 |
| Maximize half → companion  | Pair unregisters; companion not auto-maximized |
| Heal silent / fill toasted | Fill toasts only on Y; heal silent             |
| Swap `[F]` replace-skip    | Removed                                        |
| Y lone half                | Expand (maximize), do not 50/50 with BG        |

## Hotkey

[`WindowManagement/hotkeys.ahk`](../../../WindowManagement/hotkeys.ahk): `^!#y:: WM_WindowTools_OnTileBackground()` — sole user gesture for free-capacity fill (BG import and/or lone-half expand).
