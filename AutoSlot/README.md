# AutoSlot (experimental)

Optional auto-positioning for newly opened windows on multi-monitor setups. Included from `WindowManagement.ahk`.

Detection and placement policy live in this folder. **50/50 snaps reuse** the gapless algorithm in `WindowManagement\tile_snap.ahk` (same engine as `Ctrl+Alt+Win+X`).

After a successful 50/50 snap, a **2-second** interactive banner (`StandardLoadingBar_ShowWithKeys`) offers **[M]** to maximize the new window on ordinal monitor 2 and restore the partner to its pre-snap state. Timeout keeps the snap.

## Fill empty slot on close

When a window closes and leaves a monitor empty or with exactly one visible window, AutoSlot promotes the first eligible **background** window (`WM_CollectBackgroundWindows`: covered and/or minimized):

| Occupancy after close | Action                                                 |
| --------------------- | ------------------------------------------------------ |
| 0 visible             | Maximize background window onto that monitor           |
| 1 visible             | 50/50 snap background window with the remaining window |
| 2+                    | No-op                                                  |

Only the monitor that lost the window is considered. No undo modal on this path. Restoring a background window is marked recent so the new-window placer does not run again on the same HWND.

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
