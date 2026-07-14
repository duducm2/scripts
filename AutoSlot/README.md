# AutoSlot (experimental)

Optional auto-positioning for newly opened windows on multi-monitor setups. Included from `WindowManagement.ahk`.

Detection and placement policy live in this folder. **50/50 snaps reuse** the gapless algorithm in `WindowManagement\tile_snap.ahk` (same engine as `Ctrl+Alt+Win+X`).

AutoSlot ignores **AutoHotkeyGUI** windows, AutoHotkey host processes, and windows owned by this script’s PID (prompts, selectors, overlays). They stay where they opened and do not count toward monitor occupancy or fill-on-close. ToolWindow loading banners were already skipped via `WS_EX_TOOLWINDOW`.

## Enable / disable

From **Win+Alt+Shift+W** (Window tools), press **[5]** to toggle AutoSlot ON/OFF. The banner title and footer show the current state. Preference is stored in `assets/data/wm_autoslot.ini` (`[AutoSlot] Enabled=1|0`) and applied without reloading hooks: when OFF, place and fill-on-close are no-ops.

Default is ON when the ini is missing.

## New-window placement

Capacity model: **2 slots per ordinal monitor** (up to 8). A single maximized window uses one full-screen presentation but only **one** slot — the other half-slot stays available for 50/50.

On show (multi-monitor only):

| Condition                                            | Action                                                  |
| ---------------------------------------------------- | ------------------------------------------------------- |
| Any empty ordinal (0 visible excl. new hwnd)         | Maximize new window onto the first empty monitor        |
| No empty; **origin** has exactly 1 other             | 50/50 on origin — partner keeps its pane/side           |
| No empty; origin full; another monitor has exactly 1 | 50/50 on that half-full monitor (incl. maximized-alone) |
| All ordinals have 2+ windows                         | Maximize new window in place (“grid full”)              |

Empty monitors always win over 50/50. Origin is preferred before any remote half-full so unrelated layouts are not reshaped when the origin can partition. During Place, fill-on-close will not SnapPair a background window into a half-full monitor (heal-only under place freeze).

After a successful 50/50 snap, a **2-second** interactive banner (`StandardLoadingBar_ShowWithKeys`) offers **[M]** to maximize the new window on ordinal monitor 2 and restore the partner to its pre-snap state. Timeout keeps the snap.

Maximizing either half of a registered 50/50 pair (Win+Up, title-bar, or script maximize) also maximizes the companion so no orphaned half-pane remains. Disabled when AutoSlot is OFF.

## Fill empty slot on close

When a window closes and leaves a monitor empty or with exactly one visible window:

| Occupancy after close | Action                                                               |
| --------------------- | -------------------------------------------------------------------- |
| 0 visible             | Maximize first eligible background window onto that monitor          |
| 1 visible             | Maximize the remaining companion to 100% (never import a background) |
| 2+                    | No-op                                                                |

Empty-monitor background candidates skip windows that still have a living 50/50 snap-pair partner, and windows whose home monitor has another window in **F11 fullscreen** (so an F11-covered companion is not stolen into another slot). True minimized / unrelated background windows remain eligible.

Only the monitor that lost the window is considered. No undo modal on this path. Companions already maximized/work-area-filled skip work. Restoring a background window is marked recent so the new-window placer does not run again on the same HWND. Background fill is suppressed for a monitor recently claimed by new-window placement, but companion heal still runs during that cooldown and is not blocked when the leftover window is tagged recent. If occupancy still shows 2+ windows briefly (e.g. Chrome closing), fill retries once (~400 ms) before giving up.

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

`tile_snap.ahk` is unchanged by removal; no other modules reference AutoSlot symbols.
