# Windows Rearrange (AutoSlot) — Canon

**Status:** Authoritative for current behavior.  
**Audience:** Future AI runs and humans changing window placement.  
**Code:** [`AutoSlot/AutoSlot.ahk`](../../AutoSlot/AutoSlot.ahk), [`WindowManagement/tile_snap.ahk`](../../WindowManagement/tile_snap.ahk), [`WindowManagement/move_monitor.ahk`](../../WindowManagement/move_monitor.ahk).

Prefer this document over [`docs/archive/autoslot-rearrange-audit/`](../archive/autoslot-rearrange-audit/) (historical workbench; may describe pre-change policy). Short enablement notes also live in [`AutoSlot/README.md`](../../AutoSlot/README.md).

---

## Gates

- AutoSlot **ON** (Window tools **Win+Alt+Shift+W** → **[5]**; preference in `assets/data/wm_autoslot.ini`).
- **More than one** monitor.
- Capacity: **2 slots** per ordinal monitor; up to **4** ordinals → **8** max.
- Lone maximized / work-area-filled window = **1 slot** (free half remains for Place / Ctrl+Alt+Win+6 fill / Ctrl+Alt+Win+Y list). Windows hidden behind a maximized window do not fill the free half.
- Exclusions: ClipAngel, tool windows, dialog noise, AutoHotkey GUIs / own PID — not moved and not occupancy. **Teams:** chat (`Chat | … | Microsoft Teams`) and meeting (`{name} | Microsoft Teams`) are eligible; **sharing control bar** and **meeting compact view** are hard-excluded via `WM_IsTeamsChromeHwnd` (Win32 title + UIA fallback for compact when title looks like a meeting) for tile/move/cycle and AutoSlot Place/fill/occupancy — not by ignoring whole `ms-teams.exe`. User ignore list via **Win+Alt+Shift+L** → **[R]** add / **[I]** manage (`assets/data/autoslot_user_excludes.ini`) has the same gate effect — do **not** add `ms-teams.exe` via **[R]** to hide chrome (that also blocks chat/meeting).

---

## Place (new window opens)

`AutoSlot_Place` (shell create / show → debounced):

1. **Empty ordinal** → maximize new window onto the first empty monitor (ordinal order).
2. Else **free half** (lone maximized or lone half-pane) → **50/50 SnapPair** with that partner; first such monitor in ordinal order (**M1**, then M2, …). After register, both hwnds get **PairSuppress** so settle `LOCATIONCHANGE` (stale OS-max bit) does not immediately unpair.
3. Else → **leave as-is** (“grid full”) — do not maximize over true full monitors (two filled or an existing 50/50 pair); window stays where the OS opened it.

**SHOW vs CREATE:** Shell `WINDOWCREATED` always may Place. `EVENT_OBJECT_SHOW` uses `AutoSlot_ScheduleFromShow` — if the hwnd already shares its monitor with another occupant (e.g. a 50/50 half activated by **`^!#q/w/e/r`**), it does **not** Place/maximize. It only **`RememberHwndMon`** when the hwnd is an eligible / occupancy candidate (Teams/#32770 chrome is not cached, so dismiss cannot false-heal a lone half). Cycle activate must never yank a half onto an empty ordinal. Manual **`^!#x`** 50/50 also applies **PairSuppress** after register (same settle mute as Place). **Performance:** while Place is active (`g_AutoSlotPlaceDepth > 0`), SHOW callbacks are deferred (`place_active`). The occupied-mon check uses a **~150 ms occupancy cache**, `g_AutoSlotHwndMon`, or an early-exit scan — not full `PartitionOccupancy` per SHOW. WinEvent SHOW is **queued** (`AutoSlot_QueueScheduleFromShow`, 50 ms coalesce, batch 12) so CREATE debounce timers are not blocked by SHOW storms.

**Eligibility settle (required for Place latency):** After the debounce, `AutoSlot_ProcessPending` calls `AutoSlot_IsEligibleNewWindow`. Many apps open with an empty title / not-yet-ready HWND, so the first check fails. **Do not** one-shot abandon: start a **dense settle** (`AutoSlot_ScheduleEligRetry`) that re-checks about every **100 ms** until eligible or a **~2 s** budget from the first miss (`AutoSlot_ELIG_RETRY_POLL_MS` / `AutoSlot_ELIG_RETRY_BUDGET_MS`). Do **not** use sparse 300/800/1500 gaps (those wait for the next long timer even after the title is ready and reintroduce multi-second lag). `AutoSlot_Schedule` must **not** `RememberHwndMon` before eligibility succeeds (pending-only until then; coalesce while settle is armed). Giving up after a single miss left windows unplaced until a later SHOW.

**Do not “fix” Place delay by removing busy banners.** `StandardLoadingBar_BusyAllMonitors_*` (“Arranging window…”) is intentional feedback during maximize/snap/swap. Disabling it does not fix eligibility/Place timing and must not be treated as a performance improvement.

Feedback: INFO toast on successful empty/half place. After SnapPair, optional **[M]** undo modal where that path still uses it. Grid-full is silent.

---

## Close / minimize (heal)

When one window of a 50/50 pair **closes** or **minimizes**:

- Snap pair is cleared.
- Leftover companion is **healed** (maximized) when applicable — **intentional**.
- Minimize heal matches destroy: resolve monitor via cache → live → partner; dual `HealKnownCompanion` + `ScheduleHealOnly` / late `HealLoneCompanion` (occupancy gates still block false maximize).
- Heal only when the monitor has a **true lone uncovered half** (no uncovered filled occupant). Do not maximize a half beside an already-filled window (would stack two fulls).
- Occupancy for heal gates uses a **cover filter** (window center inside a higher z-order occupant’s rect, same idea as cycle’s visible list). Covered-behind noise does **not** block heal; only **uncovered** peers do.
- **`HealKnownCompanion`** aborts if another uncovered occupant remains (guards false/premature DESTROY that would bury the other half behind a maximized companion).
- Premature WinEvent **`EVENT_OBJECT_DESTROY`** while the hwnd is still a visible occupancy candidate is ignored (no unregister / no heal); shell `WINDOWDESTROYED` remains the primary real-close path.
- **No** automatic background import onto empty/half monitors.

---

## Maximize one half of a pair

Maximizing one half of a registered pair **unregisters** the pair and **does not** maximize the companion (companion stays half-sized).

---

## Explicit fill

| Action                                    | Behavior                                                                                                                                                                                                                                                                                                                                                                                                                                                                  |
| ----------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Ctrl+Alt+Win+6** / Window tools **[3]** | Free-capacity fill (`forceImport`): empty → import BG; **same-monitor extras pair first** (e.g. Teams behind/beside Excel on M3 → 50/50 there before stealing BG to another ordinal free half); then remaining free halves in ordinal order. Lone half with no BG → after BG pass, **expand to full monitor**. Already-slotted pairs are not relocated across monitors. Candidates: Y-list/hidden first, then visible unslotted. QC: `.cursor/autoslot_fill_quality.log`. |
| **Ctrl+Alt+Win+Y** open                   | Places the **chosen** background HWND into empty / free half (not a full scan). List remains **hidden-only**. Lone half still pairs with the chosen window via Place/Ctrl+Y path.                                                                                                                                                                                                                                                                                         |
| **Study Topic QuickLook** (`#!+X` / open) | After layout, **`QuickLook_ScheduleAutoSlotPlace`** (deferred ~400 / 1200 / 2500 / 4000 ms; PostMessage + file IPC) → `TryPlaceBackgroundHwnd`; AutoSlot **sticky-retries** until max/half geometry sticks (QL `PositionWindow` undoes size unless Maximized). Same empty / free-half policy as Ctrl+Alt+Win+Y list open.                                                                                                                                                 |

---

## Suite move / snap swap

With AutoSlot ON, suite move-to-monitor (`Ctrl+Alt+Win+A/S/D/F` → `MoveWinToMonitor`) or snap (`Ctrl+Alt+Win+X`) may **swap whole-monitor foreground layouts**. Manual drag does not swap. After swap: INFO toast; **no** `[F]` replace. User fills freed capacity with **Ctrl+Alt+Win+6** if desired.

---

## Busy loading (all monitors)

While AutoSlot / WindowManagement resizes or snaps (Place maximize/SnapPair, heal maximize, gapless snap, suite move, foreground swap), nestable:

- `StandardLoadingBar_BusyAllMonitors_Begin` / `_End`
- Default message: arranging-window loading indication on **every** monitor
- Visual only (does not block input)

See [`docs/canon/standard_information_display.md`](standard_information_display.md).

Rearrange INFO toasts / fill loading use `BANNER_ACCENT_INFO`.

---

## Scenario map (quick)

| Scenario                   | Outcome                                                                |
| -------------------------- | ---------------------------------------------------------------------- |
| New window                 | Empty → max; else free half → 50/50; else leave as-is                  |
| Close / minimize of a pair | Heal leftover companion; no BG import                                  |
| Maximize one half          | Unregister pair; companion unchanged                                   |
| Ctrl+Alt+Win+6 / menu [3]  | Same-mon BG pair first; then ordinal fill; expand leftover lone halves |
| Ctrl+Alt+Win+Y open        | Chosen HWND into free capacity                                         |
| Suite move / ^!#x          | May swap FG layouts                                                    |

---

## AI rule

When changing or explaining Windows rearrangement:

1. Read **this canon** and the code paths above.
2. Do **not** treat archived audit docs or old plans as current policy unless the task is historical analysis.
3. Do not reintroduce automatic background import on close/move/minimize unless the user explicitly requests it.
