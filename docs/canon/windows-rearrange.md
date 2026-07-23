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
- Lone maximized / work-area-filled window = **1 slot** (free half remains for Place / Y / Ctrl+6). Windows hidden behind a maximized window do not fill the free half.
- Exclusions: ClipAngel, tool windows, dialog/Teams chrome noise, AutoHotkey GUIs / own PID — not moved and not occupancy.

---

## Place (new window opens)

`AutoSlot_Place` (shell create / show → debounced):

1. **Empty ordinal** → maximize new window onto the first empty monitor (ordinal order).
2. Else **free half** (lone maximized or lone half-pane) → **50/50 SnapPair** with that partner; first such monitor in ordinal order (**M1**, then M2, …).
3. Else → **MaximizeInPlace** (“grid full”) — do not demax true full monitors (two filled or an existing 50/50 pair).

Feedback: INFO toast. After SnapPair, optional **[M]** undo modal where that path still uses it.

---

## Close / minimize (heal)

When one window of a 50/50 pair **closes** or **minimizes**:

- Snap pair is cleared.
- Leftover companion is **healed** (maximized) when applicable — **intentional**.
- **No** automatic background import onto empty/half monitors.

---

## Maximize one half of a pair

Maximizing one half of a registered pair **unregisters** the pair and **does not** maximize the companion (companion stays half-sized).

---

## Explicit fill

| Action                                    | Behavior                                                                                                                      |
| ----------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------- |
| **Ctrl+Alt+Win+Y** / Window tools **[3]** | **Only** full background→slot scan fill (`forceImport`): empty ordinals first, then free halves (including lone max → 50/50). |
| **Ctrl+Alt+Win+6** open                   | Places the **chosen** background HWND into empty / free half (not a full scan).                                               |

---

## Suite move / snap swap

With AutoSlot ON, suite move-to-monitor (`Ctrl+Alt+Win+A/S/D/F` → `MoveWinToMonitor`) or snap (`Ctrl+Alt+Win+X`) may **swap whole-monitor foreground layouts**. Manual drag does not swap. After swap: INFO toast; **no** `[F]` replace. User fills freed capacity with **Y** if desired.

---

## Busy loading (all monitors)

While AutoSlot / WindowManagement resizes or snaps (Place maximize/SnapPair, heal maximize, gapless snap, suite move, foreground swap), nestable:

- `StandardLoadingBar_BusyAllMonitors_Begin` / `_End`
- Default message: arranging-window loading indication on **every** monitor
- Visual only (does not block input)

See [`docs/canon/standard_information_display.md`](standard_information_display.md).

Rearrange INFO toasts / Y-fill loading use `BANNER_ACCENT_INFO`.

---

## Scenario map (quick)

| Scenario                   | Outcome                                                   |
| -------------------------- | --------------------------------------------------------- |
| New window                 | Empty → max; else free half → 50/50; else MaximizeInPlace |
| Close / minimize of a pair | Heal leftover companion; no BG import                     |
| Maximize one half          | Unregister pair; companion unchanged                      |
| Y / menu [3]               | Background scan fill                                      |
| Ctrl+6 open                | Chosen HWND into free capacity                            |
| Suite move / ^!#x          | May swap FG layouts                                       |

---

## AI rule

When changing or explaining Windows rearrangement:

1. Read **this canon** and the code paths above.
2. Do **not** treat archived audit docs or old plans as current policy unless the task is historical analysis.
3. Do not reintroduce automatic background import on close/move/minimize unless the user explicitly requests it.
