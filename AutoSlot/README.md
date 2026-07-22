# AutoSlot (experimental)

Optional auto-positioning for newly opened windows on multi-monitor setups. Included from `WindowManagement.ahk`.

Detection and placement policy live in this folder. **50/50 snaps reuse** the gapless algorithm in `WindowManagement\tile_snap.ahk` (same engine as `Ctrl+Alt+Win+X`).

AutoSlot ignores **AutoHotkeyGUI** windows, AutoHotkey host processes, and windows owned by this script’s PID (prompts, selectors, overlays). They stay where they opened and do not count toward monitor occupancy. ToolWindow loading banners were already skipped via `WS_EX_TOOLWINDOW`.

## Enable / disable

From **Win+Alt+Shift+W** (Window tools), press **[5]** to toggle AutoSlot ON/OFF. The banner title and footer show the current state. Preference is stored in `assets/data/wm_autoslot.ini` (`[AutoSlot] Enabled=1|0`) and applied without reloading hooks: when OFF, place and Y-fill are no-ops.

Default is ON when the ini is missing.

## New-window placement

Capacity model: **2 slots per ordinal monitor** (up to 8). A single maximized window uses one full-screen presentation but only **one** slot — the other half-slot stays available for 50/50 (via **Y** or Ctrl+6, not on open).

On show (multi-monitor only):

| Condition                                    | Action                                           |
| -------------------------------------------- | ------------------------------------------------ |
| Any empty ordinal (0 visible excl. new hwnd) | Maximize new window onto the first empty monitor |
| No empty ordinal                             | Maximize new window in place (“grid full”)       |

Place never 50/50 or demaximizes existing slotted windows. During Place freeze, automatic companion heal on close/minimize still runs; background import does not (and is never automatic).

After a successful 50/50 snap (from Y / Ctrl+6 / explicit snap), a **2-second** interactive banner may offer **[M]** undo where that path still uses the undo modal.

Maximizing one half of a registered 50/50 pair **does not** maximize the companion — the pair is unregistered and the other half stays. Closing or minimizing one half still **heals** (maximizes) the leftover companion.

## Foreground monitor swap

When AutoSlot is ON, suite move-to-monitor (`Ctrl+Alt+Win+A/S/D/F` → `MoveWinToMonitor`) **or** snap (`Ctrl+Alt+Win+X`) may **swap whole-monitor foreground layouts**. Manual drag does not swap.

After a successful swap, an INFO toast reports the exchange. There is **no** `[F]` replace option. To fill freed capacity from background, use **Ctrl+Alt+Win+Y**.

## Close / minimize (heal only)

When a window closes or minimizes:

- Snap pair is cleared; leftover companion is **healed** (maximized) when applicable — silent.
- **No** automatic background promote onto empty/half monitors.

## Fill free slots (Ctrl+Alt+Win+Y) — only background import

When AutoSlot is ON, **Ctrl+Alt+Win+Y** (Window tools **[3]**) is the **only** path that imports background windows into free ordinal capacity: empty monitors first, then half-slots including a **lone maximized** window (50/50 in place). Existing 50/50 pairs and multi-filled monitors are not reshuffled. When AutoSlot is OFF, the legacy background tile runs with at most **2 per monitor** and still skips already-slotted windows.

## Background list open (Ctrl+Alt+Win+6)

When AutoSlot is ON, picking a window from the hidden-background list places it into free capacity: empty ordinal → maximize; one free half → 50/50. If no free slot (or AutoSlot OFF), it restores in place.

## Disable

Comment out or remove this line in `WindowManagement.ahk`:

```ahk
#include %A_ScriptDir%\AutoSlot\AutoSlot.ahk
```

Reload `WindowManagement.ahk`.

## Delete after testing

1. Remove the `#include` line above from `WindowManagement.ahk`.
2. Delete the `AutoSlot\` folder.
3. Reload `WindowManagement.ahk`.
