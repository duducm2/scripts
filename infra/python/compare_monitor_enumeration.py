#!/usr/bin/env python3
"""
One-shot monitor enumeration for comparison with AutoHotkey MonitorGet / GetMonitorIndexByOrder.

Uses the same Win32 API family as wm_hooks._get_monitor_handles (EnumDisplayMonitors order).

Run from repo:
  cd python
  python compare_monitor_enumeration.py

Pair with: AutoHotkey64.exe ..\\infra\\tools\\MonitorEnumerationSnapshot.ahk
"""

from __future__ import annotations

import sys


def main() -> int:
    try:
        import win32api
        import win32con
    except ImportError:
        print(
            "compare_monitor_enumeration: install pywin32 (pip install pywin32)",
            file=sys.stderr,
        )
        return 1

    monitors = win32api.EnumDisplayMonitors(None, None)
    print("Python EnumDisplayMonitors order (1-based index = wm_daemon mon_index):")
    print("idx | HMONITOR(hex) | L T R B (work area via GetMonitorInfo)")
    for i, entry in enumerate(monitors, start=1):
        hmon = entry[0]
        hmon_i = int(hmon)
        try:
            info = win32api.GetMonitorInfo(hmon)
            wr = info["Work"]
            # wr is (left, top, right, bottom)
            l, t, r, b = wr
            cx = (l + r) // 2
            cy = (t + b) // 2
            print(f"  {i} | 0x{hmon_i:X} | work {l} {t} {r} {b} | center ({cx},{cy})")
        except Exception as e:  # noqa: BLE001
            print(f"  {i} | 0x{hmon_i:X} | <GetMonitorInfo error: {e}>")
    print()
    print(
        "Compare: AHK ordinal 1 = leftmost by center-x, then center-y (see tools snapshot)."
    )
    print(
        "If AHK MonitorGet index K != EnumDisplayMonitors[K], daemon GetVisibleWindowsByMonitor(K) may be wrong."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
