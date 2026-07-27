# AutoSlot efficiency pass

**Date:** 2026-07-24  
**Canon:** [`docs/efficiency-canon.md`](efficiency-canon.md) (§3 repeated enumeration / polling, §4 single authority, §9 temp-file IPC)  
**Place latency (required behavior):** [`docs/canon/windows-rearrange.md`](canon/windows-rearrange.md) — eligibility retry.

## Changes

| Canon                | Change                                                                                                                                                                                 |
| -------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Dead paths           | Removed no-op `ScheduleRearrange` / `ScheduleFill` / `ProcessFill*` / `RearrangeUnderfilled` and suite call sites                                                                      |
| Repeated enumeration | `Ctrl+Alt+Win+Y` collects `WM_CollectBackgroundWindows` **once** per pass (`g_AutoSlotYBgRows`); picks consume from that list                                                          |
| Place occupancy      | `AutoSlot_Place` / `TryPlaceBackgroundHwnd` use **one** `WinGetList` snapshot (`BuildOccupancyByMonitor`) for empty + free-half search                                                 |
| Heal accounting      | Fill returns `"healed"` for lone-half expand so Y need not re-partition before/after                                                                                                   |
| Polling / file IPC   | Place-request file poll slowed **200 ms → 1000 ms** (PostMessage remains preferred; file is Shift keys fallback)                                                                       |
| Timer pile-up        | Destroy/minimize: dropped redundant late `HealLoneCompanion` (covered by `ScheduleHealOnly`)                                                                                           |
| Place reentrancy     | `Critical` + `g_AutoSlotPlaceDepth` during Place; SHOW deferred while Place active; occupied-mon check uses cache / tracked / early-exit scan (not full `PartitionOccupancy` per SHOW) |

## Place reentrancy fix (keep this)

**Symptom (work log):** `Place_enter` then 20+ s wall gap before `Place_freeze_done`, while internal `ms=797`. `ProcessPending_elig_ok` without `Place_enter` on sibling hwnds. SHOW storm: dozens of `ScheduleFromShow_skip occupied_mon=N`.

**Root cause:** WinEvent `EVENT_OBJECT_SHOW` re-entered during `AutoSlot_Place` and ran full `PartitionOccupancy` → `OccupancyOnMonitor` per callback, starving Place.

**Required fix (maintained):** `AutoSlot_BeginPlaceCritical` / `EndPlaceCritical` wrap `Place`, `TryPlaceBackgroundHwnd`, and post-elig `ProcessPending`. `ScheduleFromShow` returns immediately when `g_AutoSlotPlaceDepth > 0`. Occupied-mon gate: fresh `BuildOccupancyByMonitor` cache (**150 ms**), else `g_AutoSlotHwndMon` scan, else `MonitorHasOtherOccupant` early-exit — **not** `PartitionOccupancy` on every SHOW. **SHOW defer:** WinEvent queues SHOW to `AutoSlot_ProcessShowPending` (50 ms, batch 12) so debounce timers are not starved by SHOW storms.

## Place latency fix (keep this)

**Symptom after efficiency-era Place/SHOW hardening:** new windows waited ~3–4 s before rearrange (previously ~1–2 s). Busy-all-monitors banners were **not** the cause.

**Root cause:** `ProcessPending` failed `IsEligibleNewWindow` once (empty title / HWND not ready), called `ForgetHwndMon`, and **returned with no retry**. Place only happened later via another SHOW (or never felt timely).

**Required fix (maintained):** `AutoSlot_ScheduleEligRetry` — on eligibility miss, **dense ~100 ms polls** until eligible or **~2 s** from first miss (`AutoSlot_ELIG_RETRY_POLL_MS` / `AutoSlot_ELIG_RETRY_BUDGET_MS`). `Schedule` is pending-only (no `RememberHwndMon` before elig OK; coalesce while settle armed). Do **not** revert to one-shot abandon or sparse 300/800/1500 delays.

**Anti-patterns (do not repeat):**

- Setting `STANDARD_BUSY_ALL_MONITORS_DISABLED := true` to “speed up” rearrange — removes the arranging indicator without fixing Place timing.
- Removing eligibility settle or reverting to sparse 300/800/1500 delays — reintroduces multi-second sit-then-snap lag.
- Remembering hwnd in `Schedule` before eligibility OK — races with SHOW skips and settle.
- Treating file IPC poll **1000 ms** as the normal-window Place delay — that poll only affects cross-process / QL file fallback, not shell CREATE/SHOW Place.

Optional diagnose: perf log → `.cursor/autoslot_perf.log`. **Auto-on at work** (`IS_WORK_ENVIRONMENT`). **Personal:** create empty `.cursor/autoslot_perf_debug` or set env `AUTOSLOT_PERF_LOG=1`. Disable: `AUTOSLOT_PERF_LOG=0`. Reload WindowManagement = fresh log (truncated). See **Work debugging** below.

## Work debugging

Use this when rearrange feels slow on the work PC but not at home.

**Log path:** `<scripts-repo>\.cursor\autoslot_perf.log`  
(e.g. `C:\Users\fie7ca\Documents\01 - Scripts\.cursor\autoslot_perf.log`)

**Enable:**

- **Work:** sync latest repo, reload `WindowManagement.ahk` (automatic).
- **Personal (dry run):** create empty file `.cursor/autoslot_perf_debug` in the scripts repo, then reload WindowManagement.

First line after reload must be `session_start env=work …` or `env=personal …`. If you only see an old `ShellCREATED` line with no `session_start`, logging was off on that reload — enable as above and reload again.

**Test protocol:**

1. Reload WindowManagement (fresh log).
2. Open 3–5 windows that felt slow (Notepad/Explorer baseline, browser tab, Teams if used, worst offender).
3. Copy `autoslot_perf.log` to personal and paste/attach in chat for analysis.

**Log interpretation:**

| Pattern                                                                        | Likely cause                                                    |
| ------------------------------------------------------------------------------ | --------------------------------------------------------------- |
| Large gap `ShellCREATED` → `ProcessPending_elig_ok` with many `EligRetry_fire` | App slow to get title (eligibility settle)                      |
| `ProcessPending_elig_fail reason=excluded` repeating                           | Window wrongly excluded; never Places                           |
| `ScheduleFromShow_skip occupied_mon=N`                                         | SHOW path blocked                                               |
| `HandlePlaceRequest ipc path=file` without `ShellCREATED`                      | Cross-process / QL file IPC (1 s poll)                          |
| `Place_after_occ_scan` with high `hwndTotal` / `teamsUiaMs`                    | Desktop enumeration or Teams UIA during occupancy               |
| `Place_enter` → `Place_freeze_done` gap 20+ s, low internal `ms=`              | SHOW reentrancy during Place (fixed by Critical + SHOW defer)   |
| `ScheduleFromShow_skip place_active` during Place                              | Expected — SHOW deferred until Place completes                  |
| `ScheduleFromShow_skip occupied_mon=N via=cache`                               | Fast occupied path (cache hit)                                  |
| `Place_exit path=snap` with high `snapMs`                                      | Snap validate / demax path                                      |
| High `total=` but low `occBuildMs`                                             | Delay is **before** Place (debounce/eligibility), not occupancy |

**Turn off after diagnosis:** set user env `AUTOSLOT_PERF_LOG=0` or sync when the debug pass is merged off.

## Unchanged policy

Y-only background import, Place empty/half rules, heal-on-close/minimize, no paired auto-max, no swap `[F]`. Busy-all-monitors arranging banner stays on during maximize/snap/swap.

## Verification

- Y: same fill outcomes; ideally one background collect per press
- Close/minimize: companion still heals; no background import
- Suite leave: no `ScheduleRearrange` call
- Cross-process place: PostMessage first; file poll still works slowly
- New window Place: dense eligibility settle (~100 ms); Place soon after title ready; no multi-second dead wait from one-shot abandon or sparse retry gaps
