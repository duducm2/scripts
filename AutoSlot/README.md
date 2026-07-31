# AutoSlot (Windows Rearrange)

Optional auto-positioning for newly opened windows on multi-monitor setups. Included from `WindowManagement.ahk`.

**Authoritative behavior:** [`docs/canon/windows-rearrange.md`](../docs/canon/windows-rearrange.md) — read that before changing Place, heal, fill, or swap logic.

**Efficiency / Place latency:** [`docs/autoslot-efficiency.md`](../docs/autoslot-efficiency.md) — one-scan fill (`Ctrl+Alt+Win+6`: hidden + visible unslotted), place-request poll, dense eligibility settle, and the **2026-07-27 work-PC latency fix** (SHOW reentrancy, occupancy cache, SHOW queue). Do not one-shot abandon eligibility or sparse 300/800/1500 retries.

Detection/placement live in this folder. **50/50 snaps** reuse `WindowManagement\tile_snap.ahk` (same engine as `Ctrl+Alt+Win+X`).

AutoSlot ignores **AutoHotkeyGUI** windows, AutoHotkey host processes, and windows owned by this script’s PID (prompts, selectors, overlays). ToolWindow loading banners are skipped via `WS_EX_TOOLWINDOW`. Keep the all-monitors **Arranging window…** busy banner during Place/snap/swap — it is feedback, not the Place-delay fix.

## Enable / disable

From **Win+Alt+Shift+W** (Window tools), press **[5]** to toggle AutoSlot ON/OFF. Preference: `assets/data/wm_autoslot.ini` (`[AutoSlot] Enabled=1|0`). Default ON when the ini is missing. When OFF, place and fill are no-ops.

## Quick reference

| Action                    | Hotkey / UI                                                                   |
| ------------------------- | ----------------------------------------------------------------------------- |
| Toggle AutoSlot           | Win+Alt+Shift+W → **[5]**                                                     |
| Background fill           | Ctrl+Alt+Win+6 / Window tools **[3]**                                         |
| Open from background list | Ctrl+Alt+Win+Y                                                                |
| Manual 50/50              | Ctrl+Alt+Win+X                                                                |
| Move to monitor           | Ctrl+Alt+Win+A/S/D/F                                                          |
| User ignore list          | Win+Alt+Shift+L → **[R]** add / **[I]** manage (`autoslot_user_excludes.ini`) |

Place: empty monitor → maximize; else free half → 50/50; else leave as-is (do not cover). Close/minimize of a pair **heals** (maximizes) the leftover companion. **Ctrl+Alt+Win+6**: prefer same-monitor background pairing, then fill other free halves; leftover lone halves expand to full. Already-slotted pairs stay on their monitor. Details and busy overlays: canon doc above.

User ignore list (via **#!+L**): same effect as built-in ClipAngel exclusion — no place/fill/occupancy. **[R]** arms pick-to-ignore by **process exe**; **[I]** opens a digit-remove list. Do **not** ignore `ms-teams.exe` for chrome — that blocks chat and meeting too. Hard-coded `WM_IsTeamsChromeHwnd` (plus AutoSlot) skips **Sharing control bar** and **Meeting compact view** for tile/move/cycle and Place/fill; chat/meeting stay manageable.

## Disable module

Comment out or remove in `WindowManagement.ahk`:

```ahk
#include %A_ScriptDir%\AutoSlot\AutoSlot.ahk
```

Reload `WindowManagement.ahk`.
