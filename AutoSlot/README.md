# AutoSlot (Windows Rearrange)

Optional auto-positioning for newly opened windows on multi-monitor setups. Included from `WindowManagement.ahk`.

**Authoritative behavior:** [`docs/canon/windows-rearrange.md`](../docs/canon/windows-rearrange.md) — read that before changing Place, heal, Y-fill, or swap logic.

**Efficiency notes:** [`docs/autoslot-efficiency.md`](../docs/autoslot-efficiency.md) (Y one-scan fill, place-request poll).

Detection/placement live in this folder. **50/50 snaps** reuse `WindowManagement\tile_snap.ahk` (same engine as `Ctrl+Alt+Win+X`).

AutoSlot ignores **AutoHotkeyGUI** windows, AutoHotkey host processes, and windows owned by this script’s PID (prompts, selectors, overlays). ToolWindow loading banners are skipped via `WS_EX_TOOLWINDOW`.

## Enable / disable

From **Win+Alt+Shift+W** (Window tools), press **[5]** to toggle AutoSlot ON/OFF. Preference: `assets/data/wm_autoslot.ini` (`[AutoSlot] Enabled=1|0`). Default ON when the ini is missing. When OFF, place and Y-fill are no-ops.

## Quick reference

| Action                    | Hotkey / UI                                                                   |
| ------------------------- | ----------------------------------------------------------------------------- |
| Toggle AutoSlot           | Win+Alt+Shift+W → **[5]**                                                     |
| Background fill           | Ctrl+Alt+Win+Y / Window tools **[3]**                                         |
| Open from background list | Ctrl+Alt+Win+6                                                                |
| Manual 50/50              | Ctrl+Alt+Win+X                                                                |
| Move to monitor           | Ctrl+Alt+Win+A/S/D/F                                                          |
| User ignore list          | Win+Alt+Shift+L → **[R]** add / **[I]** manage (`autoslot_user_excludes.ini`) |

Place: empty monitor → maximize; else free half → 50/50; else leave as-is (do not cover). After debounce, Place may **retry eligibility** briefly (300 / 800 / 1500 ms) when a new window is not yet titled — required; do not one-shot abandon (see canon). Close/minimize of a pair **heals** (maximizes) the leftover companion. **Y**: lone half + free slot → maximize that window; lone max → 50/50 with background. Busy-all-monitors rearrange banners stay on.

User ignore list (via **#!+L**): same effect as built-in ClipAngel exclusion — no place/fill/occupancy. **[R]** arms pick-to-ignore; **[I]** opens a digit-remove list.

## Disable module

Comment out or remove in `WindowManagement.ahk`:

```ahk
#include %A_ScriptDir%\AutoSlot\AutoSlot.ahk
```

Reload `WindowManagement.ahk`.
