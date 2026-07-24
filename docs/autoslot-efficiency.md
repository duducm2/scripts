# AutoSlot efficiency pass

**Date:** 2026-07-24  
**Canon:** [`docs/efficiency-canon.md`](efficiency-canon.md) (§3 repeated enumeration / polling, §4 single authority, §9 temp-file IPC)

## Changes

| Canon                | Change                                                                                                                        |
| -------------------- | ----------------------------------------------------------------------------------------------------------------------------- |
| Dead paths           | Removed no-op `ScheduleRearrange` / `ScheduleFill` / `ProcessFill*` / `RearrangeUnderfilled` and suite call sites             |
| Repeated enumeration | `Ctrl+Alt+Win+Y` collects `WM_CollectBackgroundWindows` **once** per pass (`g_AutoSlotYBgRows`); picks consume from that list |
| Heal accounting      | Fill returns `"healed"` for lone-half expand so Y need not re-partition before/after                                          |
| Polling / file IPC   | Place-request file poll slowed **200 ms → 1000 ms** (PostMessage remains preferred; file is Shift keys fallback)              |
| Timer pile-up        | Destroy/minimize: dropped redundant late `HealLoneCompanion` (covered by `ScheduleHealOnly`)                                  |

## Unchanged policy

Y-only background import, Place empty/half rules, heal-on-close/minimize, no paired auto-max, no swap `[F]`.

## Verification

- Y: same fill outcomes; ideally one background collect per press
- Close/minimize: companion still heals; no background import
- Suite leave: no `ScheduleRearrange` call
- Cross-process place: PostMessage first; file poll still works slowly
