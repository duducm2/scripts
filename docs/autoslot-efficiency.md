# AutoSlot efficiency pass

**Date:** 2026-07-24  
**Canon:** [`docs/efficiency-canon.md`](efficiency-canon.md) (§3 repeated enumeration / polling, §4 single authority, §9 temp-file IPC)  
**Place latency (required behavior):** [`docs/canon/windows-rearrange.md`](canon/windows-rearrange.md) — eligibility retry.

## Changes

| Canon                | Change                                                                                                                                 |
| -------------------- | -------------------------------------------------------------------------------------------------------------------------------------- |
| Dead paths           | Removed no-op `ScheduleRearrange` / `ScheduleFill` / `ProcessFill*` / `RearrangeUnderfilled` and suite call sites                      |
| Repeated enumeration | `Ctrl+Alt+Win+Y` collects `WM_CollectBackgroundWindows` **once** per pass (`g_AutoSlotYBgRows`); picks consume from that list          |
| Place occupancy      | `AutoSlot_Place` / `TryPlaceBackgroundHwnd` use **one** `WinGetList` snapshot (`BuildOccupancyByMonitor`) for empty + free-half search |
| Heal accounting      | Fill returns `"healed"` for lone-half expand so Y need not re-partition before/after                                                   |
| Polling / file IPC   | Place-request file poll slowed **200 ms → 1000 ms** (PostMessage remains preferred; file is Shift keys fallback)                       |
| Timer pile-up        | Destroy/minimize: dropped redundant late `HealLoneCompanion` (covered by `ScheduleHealOnly`)                                           |

## Place latency fix (keep this)

**Symptom after efficiency-era Place/SHOW hardening:** new windows waited ~3–4 s before rearrange (previously ~1–2 s). Busy-all-monitors banners were **not** the cause.

**Root cause:** `ProcessPending` failed `IsEligibleNewWindow` once (empty title / HWND not ready), called `ForgetHwndMon`, and **returned with no retry**. Place only happened later via another SHOW (or never felt timely).

**Required fix (maintained):** `AutoSlot_ScheduleEligRetry` — on eligibility miss, **dense ~100 ms polls** until eligible or **~2 s** from first miss (`AutoSlot_ELIG_RETRY_POLL_MS` / `AutoSlot_ELIG_RETRY_BUDGET_MS`). `Schedule` is pending-only (no `RememberHwndMon` before elig OK; coalesce while settle armed). Do **not** revert to one-shot abandon or sparse 300/800/1500 delays.

**Anti-patterns (do not repeat):**

- Setting `STANDARD_BUSY_ALL_MONITORS_DISABLED := true` to “speed up” rearrange — removes the arranging indicator without fixing Place timing.
- Removing eligibility settle or reverting to sparse 300/800/1500 delays — reintroduces multi-second sit-then-snap lag.
- Remembering hwnd in `Schedule` before eligibility OK — races with SHOW skips and settle.
- Treating file IPC poll **1000 ms** as the normal-window Place delay — that poll only affects cross-process / QL file fallback, not shell CREATE/SHOW Place.

Optional diagnose: `AutoSlot_PERF_LOG := true` → `.cursor/autoslot_perf.log` (`ProcessPending_elig_fail`, `EligRetry_*`, `Place_*` phases). Turn off when done.

## Unchanged policy

Y-only background import, Place empty/half rules, heal-on-close/minimize, no paired auto-max, no swap `[F]`. Busy-all-monitors arranging banner stays on during maximize/snap/swap.

## Verification

- Y: same fill outcomes; ideally one background collect per press
- Close/minimize: companion still heals; no background import
- Suite leave: no `ScheduleRearrange` call
- Cross-process place: PostMessage first; file poll still works slowly
- New window Place: dense eligibility settle (~100 ms); Place soon after title ready; no multi-second dead wait from one-shot abandon or sparse retry gaps
