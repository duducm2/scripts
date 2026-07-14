# AutoSlot (experimental)

Optional auto-positioning for newly opened windows on multi-monitor setups. Included from `WindowManagement.ahk`.

Detection and placement policy live in this folder. **50/50 snaps reuse** the gapless algorithm in `WindowManagement\tile_snap.ahk` (same engine as `Ctrl+Alt+Win+X`).

## Enable / disable

From **Win+Alt+Shift+W** (Window tools), press **[5]** to toggle AutoSlot ON/OFF. The banner title and footer show the current state. Preference is stored in `assets/data/wm_autoslot.ini` (`[AutoSlot] Enabled=1|0`) and applied without reloading hooks: when OFF, place and fill-on-close are no-ops.

Default is ON when the ini is missing.

## New-window placement

On show (multi-monitor only), ordinal monitors are scanned left-to-right:

| Occupancy (excluding new hwnd)  | Action                                              |
| ------------------------------- | --------------------------------------------------- |
| Any empty (0 visible)           | Maximize new window onto the first empty monitor    |
| No empty; half-full (1 visible) | 50/50 snap with the resident on the first half-full |
| No empty and no half-full       | Maximize new window in place                        |

Empty monitors always win over 50/50: AutoSlot never splits an occupied screen when another monitor is free.

After a successful 50/50 snap, a **2-second** interactive banner (`StandardLoadingBar_ShowWithKeys`) offers **[M]** to maximize the new window on ordinal monitor 2 and restore the partner to its pre-snap state. Timeout keeps the snap.

## Fill empty slot on close

When a window closes and leaves a monitor empty or with exactly one visible window, AutoSlot promotes the first eligible **background** window (`WM_CollectBackgroundWindows`: covered and/or minimized):

| Occupancy after close            | Action                                                 |
| -------------------------------- | ------------------------------------------------------ |
| 0 visible                        | Maximize background window onto that monitor           |
| 1 visible + background available | 50/50 snap background window with the remaining window |
| 1 visible + no background        | Maximize the remaining companion                       |
| 2+                               | No-op                                                  |

Only the monitor that lost the window is considered. No undo modal on this path. Fill snaps use gapless placement **without** the ~400 ms strict validate wait (new-window place still validates). Companions already maximized/work-area-filled skip background enumeration. Restoring a background window is marked recent so the new-window placer does not run again on the same HWND. Background fill is suppressed for a monitor recently claimed by new-window placement (avoids a deferred fill racing Place and stealing the slot/foreground), but **companion heal still runs** during that cooldown and is not blocked when the leftover window is tagged recent. If occupancy still shows 2+ windows briefly (e.g. Chrome closing), fill retries once (~400 ms) before giving up.

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
