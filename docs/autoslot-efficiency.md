# AutoSlot efficiency and work-PC Place latency

**Date:** 2026-07-24 (efficiency pass); **2026-07-27** (work-PC latency fix)  
**Canon:** [`docs/efficiency-canon.md`](efficiency-canon.md) (§3 repeated enumeration / polling, §4 single authority, §9 temp-file IPC)  
**Place behavior:** [`docs/canon/windows-rearrange.md`](canon/windows-rearrange.md)

## Work-PC latency fix (2026-07-27)

### Symptom

On a **4-monitor work PC** with a busy desktop, opening a new window felt **~3–4 s slow**: the foreground partner would shrink first, then a long pause, then the new window finally snapped in. The same scripts felt fine at home.

### How we diagnosed it

1. Added optional perf logging (`AutoSlot_PerfLog` → `.cursor/autoslot_perf.log`) and enabled it on the work machine during testing.
2. Reproduced with Notepad, Explorer, browser tabs, and the worst offenders; copied the log to the personal repo for analysis.
3. Correlated log phases (`ShellCREATED`, `ProcessPending_*`, `Place_*`, `ScheduleFromShow_*`) with wall-clock gaps vs internal `ms=` timings.

**Teams UIA was ruled out:** occupancy scans on work showed `teamsUiaCalls=0`; Teams presenter-toolbar detection was not the bottleneck.

### Root causes (stacked)

| Issue                                        | Evidence in log                                                                                                                       | Fix                                                                                                                                                   |
| -------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Eligibility one-shot abandon**             | `ProcessPending_elig_fail` with no follow-up Place until a later SHOW                                                                 | Dense **~100 ms** eligibility settle (`AutoSlot_ScheduleEligRetry`, **~2 s** budget)                                                                  |
| **Double occupancy scan in Place**           | Two `WinGetList` passes per Place                                                                                                     | Single `BuildOccupancyByMonitor` snapshot per Place                                                                                                   |
| **WinEvent SHOW reentrancy during Place**    | `Place_enter` → `Place_freeze_done` **20+ s** wall gap with internal `ms≈797`; sibling `ProcessPending_elig_ok` without `Place_enter` | `Critical` + `g_AutoSlotPlaceDepth`; defer SHOW while Place active                                                                                    |
| **SHOW storm + full partition per callback** | Dozens of `ScheduleFromShow_skip occupied_mon=N`                                                                                      | **150 ms** occupancy cache, `g_AutoSlotHwndMon`, `MonitorHasOtherOccupant` early-exit — not `PartitionOccupancy` on every SHOW                        |
| **SHOW blocking debounce timers**            | CREATE → `ProcessPending` **~4.6 s** vs ~250 ms at home                                                                               | Queue SHOW: `AutoSlot_QueueScheduleFromShow` (**50 ms**, batch **12**)                                                                                |
| **Perf log FileAppend storm**                | Flood of routine occupied-skip lines starved the main thread                                                                          | Removed routine occupied-skip logging from the hot path                                                                                               |
| **Snap “thinking” feel**                     | High `snapMs`, partner demax before new window visible                                                                                | Place 50/50: `acceptUnvalidated` (skip ~400 ms validate poll); shorter `EnsureRestoredForSnap` sleeps via `AutoSlot_TrySnapNewWithPartner(..., true)` |

### Result after fixes

Healthy path on work: `Place_enter` → `Place_freeze_done` in **~16 ms**; total snap **~1.5–1.7 s** (occupancy ~700 ms + snap ~850 ms). User confirmed the lag was resolved.

### Constants (do not revert casually)

```ahk
AutoSlot_ELIG_RETRY_POLL_MS := 100
AutoSlot_ELIG_RETRY_BUDGET_MS := 2000
AutoSlot_OCC_CACHE_MS := 150
AutoSlot_SHOW_DEFER_MS := 50
AutoSlot_SHOW_BATCH_MAX := 12
```

Key helpers: `AutoSlot_BeginPlaceCritical` / `EndPlaceCritical`, `AutoSlot_QueueScheduleFromShow`, `AutoSlot_MonitorOccupiedFromCache`, `MonitorHasTrackedOccupant`, `MonitorHasOtherOccupant`, `BuildOccupancyByMonitor` (writes `g_AutoSlotOccSnap`).

### Anti-patterns (do not repeat)

- Setting `STANDARD_BUSY_ALL_MONITORS_DISABLED := true` to “speed up” rearrange — removes feedback without fixing Place timing.
- Removing eligibility settle or reverting to sparse 300/800/1500 delays — reintroduces multi-second sit-then-snap lag.
- Remembering hwnd in `Schedule` before eligibility OK — races with SHOW skips and settle.
- Treating file IPC poll **1000 ms** as the normal-window Place delay — that poll only affects cross-process / QL file fallback, not shell CREATE/SHOW Place.
- Running full `PartitionOccupancy` on every WinEvent SHOW on a busy multi-monitor desktop.

---

## Efficiency pass (2026-07-24)

| Canon                | Change                                                                                                                                                                                                                                 |
| -------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Dead paths           | Removed no-op `ScheduleRearrange` / `ScheduleFill` / `ProcessFill*` / `RearrangeUnderfilled` and suite call sites                                                                                                                      |
| Repeated enumeration | `Ctrl+Alt+Win+6` collects fill candidates **once** per pass (`AutoSlot_CollectFillCandidates` → `g_AutoSlotYBgRows`: hidden first, then visible unslotted); picks consume from that list; one QC backfill for remaining empty/lone-max |
| Place occupancy      | `AutoSlot_Place` / `TryPlaceBackgroundHwnd` use **one** `WinGetList` snapshot (`BuildOccupancyByMonitor`) for empty + free-half search                                                                                                 |
| Heal accounting      | Fill returns `"healed"` for lone-half expand so Y need not re-partition before/after                                                                                                                                                   |
| Polling / file IPC   | Place-request file poll slowed **200 ms → 1000 ms** (PostMessage remains preferred; file is Shift keys fallback)                                                                                                                       |
| Timer pile-up        | Destroy/minimize: dropped redundant late `HealLoneCompanion` (covered by `ScheduleHealOnly`)                                                                                                                                           |
| Place reentrancy     | `Critical` + `g_AutoSlotPlaceDepth` during Place; SHOW deferred while Place active; occupied-mon check uses cache / tracked / early-exit scan (not full `PartitionOccupancy` per SHOW)                                                 |

---

## Optional diagnosis

Perf logging is **off by default**. To trace Place timing on any machine:

1. Set user env **`AUTOSLOT_PERF_LOG=1`**.
2. Reload `WindowManagement.ahk` (log truncates on reload; first line is `session_start`).
3. Reproduce slow opens; inspect `<scripts-repo>\.cursor\autoslot_perf.log`.

Useful patterns: large gap before `ProcessPending_elig_ok` (eligibility settle); `Place_enter` → `Place_freeze_done` wall gap with low internal `ms=` (SHOW reentrancy — should not recur after the fix); high `total=` with low occupancy `ms=` (delay before Place, not scan cost).

---

## Unchanged policy

Explicit fill only (Ctrl+Alt+Win+6) for background import, Place empty/half rules, heal-on-close/minimize, no paired auto-max, no swap `[F]`. Busy-all-monitors arranging banner stays on during maximize/snap/swap.

## Verification

- Fill (Ctrl+Alt+Win+6): same fill outcomes; ideally one background collect per press; visible unslotted floats fill empty/lone-max; QC backfill once if needed
- Close/minimize: companion still heals; no background import
- Suite leave: no `ScheduleRearrange` call
- Cross-process place: PostMessage first; file poll still works slowly
- New window Place: dense eligibility settle (~100 ms); Place soon after title ready; no multi-second dead wait from one-shot abandon or sparse retry gaps
- Work PC: Place completes in ~1–2 s total; no 20+ s freeze during snap
