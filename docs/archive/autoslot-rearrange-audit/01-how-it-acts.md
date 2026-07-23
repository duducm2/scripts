# How rearrange acts (as implemented)

**Purpose:** Describe how Windows rearrangement **acts today** — triggers, guards, outcomes, and banners — from the code path.

**Source of truth:** [`AutoSlot/AutoSlot.ahk`](../../AutoSlot/AutoSlot.ahk) and WindowManagement call sites.  
**Policy note:** Background import is **Y-only** — see [`05-y-only-fill.md`](05-y-only-fill.md).  
**Problems / risks:** see [`00-findings-report.md`](00-findings-report.md) (some entries predate Y-only).  
**Accent:** rearrange Information Only / Y fill loading use `BANNER_ACCENT_INFO`.

---

## Capacity model

- **2 slots** per ordinal monitor, up to 4 ordinals → max **8**.
- Lone maximized / work-area = **1 slot**; free half remains for Place / Y / Ctrl+6.
- Covered-behind maximized ignored for free-half checks.
- Exclusions (ClipAngel, tool windows, dialogs/Teams noise, own PID) skipped.

---

## Scenario map

| Scenario                    | What fires                                        | Typical outcome                                           | User sees              |
| --------------------------- | ------------------------------------------------- | --------------------------------------------------------- | ---------------------- |
| **New window (Place)**      | `ProcessPending` → `Place`                        | Empty → max; else free half → 50/50; else MaximizeInPlace | INFO toast             |
| **Window close / destroy**  | Heal known companion + heal-only + late lone heal | Partner maximized; **no** background import               | Heal silent            |
| **Move / maximize end**     | Remembers monitor only                            | **No** rearrange import                                   | —                      |
| **Minimize**                | Unregister pair; heal companion                   | Heal only; **no** background import                       | Heal silent            |
| **Maximize one 50/50 half** | `OnPairedMaximize`                                | Unregister pair; companion **unchanged**                  | —                      |
| **Explicit Y / menu `[3]`** | `RunTileBackground` → Fill `forceImport`          | Empty → BG; lone half → maximize; lone max → 50/50 BG | INFO bar + slot toasts |
| **Ctrl+6 open**             | `TryPlaceBackgroundHwnd`                          | Empty / free-half place of chosen HWND                    | INFO on success        |
| **Foreground swap**         | `TryForegroundSwap` + quiet + toast               | Layouts exchange; **no** `[F]`                            | INFO swap toast        |

---

## Mute matrix (still relevant)

| Gate                                | Blocks                                             | Still allows   |
| ----------------------------------- | -------------------------------------------------- | -------------- |
| **SwapQuiet**                       | Paired-max processing during swap                  | Heal timers; Y |
| **PlaceFreeze** / **ClaimCooldown** | Background import inside Fill unless `forceImport` | Heal; Y        |
| **PairSuppress**                    | Short event storms after snap/heal                 | Y              |

---

## Policy split

| Path                            | Background import?                                                     |
| ------------------------------- | ---------------------------------------------------------------------- |
| Place / close / minimize / move | **No** (heal only on close/minimize)                                   |
| **Y / menu `[3]`**              | **Yes** (`forceImport`)                                                |
| Ctrl+6 open                     | Places the **chosen** window into free capacity (not a full scan fill) |
