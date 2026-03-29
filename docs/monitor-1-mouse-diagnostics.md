# Monitor 1: mouse and interaction diagnostics

Use this checklist when the pointer cannot reach Monitor 1 or interactions (e.g. closing) fail only on that display. Typical cause: **vertical misalignment** in Windows Display settings—a smaller monitor centered beside a taller one leaves **dead zones** where no shared edge exists between displays.

## Phase 1 — Confirm geometry vs software (5–10 minutes)

| Step | Action | Result |
|------|--------|--------|
| 1.1 | On **Monitor 2**, move the mouse to the **vertical middle** of the **left edge** (aligned with Monitor 1’s center in the layout diagram), then move **left**. | Pointer should enter Monitor 1 if geometry allows. |
| 1.2 | Repeat from the **top** and **bottom** third of Monitor 2’s left edge. | If crossing works only in the **middle**, you have **dead zones**—fix layout (Phase 3), not drivers first. |
| 1.3 | **Settings → System → Display**: drag **Monitor 1** so its **top** aligns with **Monitor 2’s top** (or align bottoms). **Apply**. Repeat 1.1–1.2. | Dead zones should shrink or disappear if alignment was the issue. |

**Optional:** Temporarily set Monitor 1 and Monitor 2 to the **same scale** (e.g. both 125% or both 150%) only to see if hit-testing feels different. Mixed scaling does not remove geometric dead zones.

## Phase 2 — Rule out cursor locking and overlays

- Exit **fullscreen games**, **remote desktop**, and **VM** sessions that may lock the cursor.
- Suspend **third-party mouse/display overlays** and retest.
- If you use **tablet mode** or unusual snap behavior, toggle and retest.

`WindowManagement.ahk` does not use `ClipCursor` or global `BlockInput` for the pointer. If the pointer still cannot reach Monitor 1 along the clearly shared vertical band after Phase 1, focus on layout and other apps before blaming this script.

## Phase 3 — Layout and scaling fixes

- **A (recommended):** Align Monitor 1’s **top** with Monitor 2’s **top** in the Display layout diagram.
- **B:** Alternatively align **bottoms** if that matches your physical desk.
- **C:** Drag monitors in Settings so the on-screen layout matches **physical** placement.
- **D:** After geometry is fixed, tune **per-monitor scale** and, for legacy apps, **Compatibility → Change high DPI settings** on specific executables.

## Phase 4 — Keyboard shortcuts (no mouse required)

Monitor index is **left-to-right** by monitor center X (`GetMonitorIndexByOrder` in `WindowManagement.ahk`). For a normal four-monitor row, the leftmost screen is **order 1**.

| Goal | Hotkey |
|------|--------|
| Close topmost window on monitor 1 | **Ctrl+Alt+Shift+Win+A** |
| Cycle windows on monitor 1 | **Ctrl+Alt+Win+Q** |
| Minimize topmost on monitor 1 | **Ctrl+Alt+Shift+Win+Q** |

**Diagnostic:** If the pointer still cannot reach Monitor 1 but **Ctrl+Alt+Shift+Win+A** closes a window that is visibly on Monitor 1, the issue is **cursor pathing** (geometry or lock), not wrong monitor targeting.

If the **wrong** window closes, verify monitor count/order (e.g. `MonitorGetCount()` in a small test script) or temporarily disconnect a display to see if Windows renumbers monitors.

## Phase 5 — Hardware / drivers

Suspect drivers or cables if failure persists **even** when crossing at the clearly shared vertical band **after** layout correction, or if you see flicker / unstable output on Monitor 1.

- Update **GPU drivers**.
- Try another **cable or port** for Monitor 1.
- Test with **one fewer monitor** connected to isolate.

## Summary

1. Confirm **dead zones** with Phase 1; **realign** Monitor 1 vs 2 (Phase 3).
2. **Rule out** pointer-locking apps (Phase 2).
3. Use **keyboard close/cycle** on monitor 1 to separate “can’t reach with mouse” from “wrong target” (Phase 4).
4. **Drivers/cables** only if geometry and software are ruled out (Phase 5).
