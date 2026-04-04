# Structured Windows display configuration (4 monitors)

**Purpose:** Environment baseline for this machine’s display layout. Read-only context for automation and AI runs. Do not treat as portable constants in scripts.

---

## Binding to automation

In [WindowManagement.ahk](../WindowManagement.ahk), hotkeys **1–4** mean **ordinal position left-to-right** (`GetMonitorIndexByOrder`), sorted by monitor rectangle center `(cx, cy)` in screen coordinates—not necessarily the same index as Windows Settings labels. This layout (small landscape → large primary → two portrait) matches **monotonic left-to-right `cx`**, so **Settings Monitor 1–4 correspond to script order 1–4** under typical alignment.

[efficiency-canon.md](efficiency-canon.md) treats **hardcoded user-specific geometry** as a portability risk; this document is **environment baseline / AI context**, not a mandate to embed literals in scripts.

---

## Spatial order (left to right)

| Ord | Role      | Native resolution | Scale | Orientation |
| --- | --------- | ----------------- | ----- | ----------- |
| 1   | Secondary | 1920 × 1080       | 125%  | Landscape   |
| 2   | **Primary** | 3840 × 2160     | 150%  | Landscape   |
| 3   | Secondary | 1080 × 1920       | 100%  | Portrait    |
| 4   | Secondary | 1080 × 1920       | 100%  | Portrait    |

---

## Per-display facts (positive, exhaustive)

- **Display 1:** Resolution 1920 × 1080. Scale 125%. Orientation landscape. Not primary.
- **Display 2:** Resolution 3840 × 2160. Scale 150%. Orientation landscape. Primary display.
- **Display 3:** Resolution 1080 × 1920. Scale 100%. Orientation portrait. Not primary.
- **Display 4:** Resolution 1080 × 1920. Scale 100%. Orientation portrait. Not primary.

**Desktop mode:** Extended desktop (not mirror). **Alignment (from Settings diagram):** Top edges aligned; display 1 sits slightly lower than display 2 in the layout diagram—ordering still follows horizontal center as in the script.

---

## Global system settings (enabled)

**Window management**

- Remember window locations based on monitor connection: **on**.
- Minimize windows when a monitor is disconnected: **on**.

**Navigation**

- Facilitate cursor movement between displays: **on**.

---

## Automation-relevant notes (concise)

- **Work areas:** [WindowManagement.ahk](../WindowManagement.ahk) uses `MonitorGet` for placement and cycling; rectangles are in **Windows screen coordinates** on the virtual desktop. Per-monitor **DPI differs** (125% / 150% / 100%); Win32/AHK calls are **DPI-aware** where documented—do not assume one scale for all monitors.
- **Ordinal stability:** Order 1–4 holds while **left-to-right geometry** is unchanged. Reordering displays in Settings or cable swaps can remap ordinals without changing the script.
- **Not captured here:** Exact virtual-screen pixel offsets for each monitor; add only if a future task needs numeric bounds (e.g. regression tests or logging).
