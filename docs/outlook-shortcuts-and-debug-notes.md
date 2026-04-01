# Outlook shortcuts (New Outlook + Classic) and recent fixes

This doc summarizes the Outlook-related hotkeys in this repo, plus the compatibility work done for **New Outlook (Monarch)**.

## New Outlook vs Classic Outlook

- **New Outlook (Monarch / Store)**
  - **Process**: `olk.exe`
  - **Window class**: `Outlook Host`
  - UI is WebView-based; many controls show up as UIA `Button` / `ComboBox` elements inside web content.
- **Classic Outlook**
  - **Process**: `OUTLOOK.EXE`
  - **Main window class**: `rctrl_renwnd32`
  - More traditional Win32/UIA control surface.

## Global Outlook hotkeys (Win+Alt+Shift…)

Defined in [`C:/Users/fie7ca/Documents/scripts/Outlook.ahk`](C:/Users/fie7ca/Documents/scripts/Outlook.ahk):

- **Win+Alt+Shift+B**: Open **Mail**
- **Win+Alt+Shift+G**: Open **Calendar**
- **Win+Alt+Shift+V**: Open **Reminders**

### What was changed for New Outlook compatibility

The original failures were caused by New Outlook using different window/process identifiers:

- **Process name mismatch**: New Outlook runs as `olk.exe` (not only `OUTLOOK.EXE`).
- **Window class mismatch**: New Outlook top-level windows are `Outlook Host` (not `rctrl_renwnd32`).
- **Window title changes**: Titles look like `Mail - … - Outlook`, `Calendar - … - Outlook`, `Reminders - … - Outlook`.
- **Module switch UIA**: Left rail items are UIA **Buttons** (type `50000`) with toggle state, not legacy list items.

Key improvements in `Outlook.ahk`:

- **Dual-process enumeration**: support both `OUTLOOK.EXE` and `olk.exe`.
- **Dual-class targeting**: accept both `rctrl_renwnd32` and `Outlook Host`.
- **Title-based mailbox/calendar selection** aligned to the new UIA captures.
- **UIA module switching** via `Mail` / `Calendar` buttons using TogglePattern when available.

## Shift-layer Outlook workflows (Shift keys.ahk)

Defined in [`C:/Users/fie7ca/Documents/scripts/Shift keys.ahk`](C:/Users/fie7ca/Documents/scripts/Shift%20keys.ahk).

### Cheat sheets and process normalization

New Outlook uses `olk.exe`, so cheat-sheet selection was updated to treat `olk.exe` as Outlook:

- When the active process is `olk.exe`, it is normalized to `OUTLOOK.EXE` for cheat-sheet routing.

### Appointment / “New event” (New Outlook) – Start date popover

Problem: In New Outlook “New event”, **Start date** and related controls may be inside a popover that is **not always open**. In scheduler views, you may need to open the popover by activating the **date/time range summary button** first.

UIA evidence (example capture):

- The date-range summary is a UIA **Button (50000)** with a dynamic name such as:
  - `"Wed 4/1/2026 2:00 PM - 2:30 PM"`

The appointment hotkeys include logic to:

- **Open the popover if needed**, using stable anchors / button selection around the date-range summary area.
- Then **invoke** the `Start date` dropdown inside the popover.

If New Outlook updates the UIA structure/labels/AutomationIds, capture a fresh UIA tree and update criteria in the appointment helpers.

### Appointment Wizard (Shift+W)

The Appointment “Wizard” (`Shift+W`) was updated to support **New Outlook** only and to follow the shared banner standards in `docs/standard_information_display.md`:

- **Interactive steps**: Each step uses `StandardLoadingBar_ShowWithKeys` (via `Appt_SelectFromModal`) to pick options with number keys and `Esc` to cancel.
- **Apply phase**: Uses `StandardLoadingBar_Show` / `StandardLoadingBar_Update` / `StandardLoadingBar_Hide` to show progress while applying Status, Privacy, All-day, Category, and Reminder using the same UIA selection helpers as the individual shortcuts (e.g., `Shift+V`, `Shift+P`, `Shift+A`, `Shift+G`, `Shift+Q`).
- **Flow**: 5 numbered steps (Status → Privacy → Category → Reminder → All-day). The prior NOTE step was removed.

## Troubleshooting / maintenance notes

- If a shortcut suddenly stops working after an Outlook update:
  - Re-capture UIA trees for the relevant surface (Mail, Calendar, Reminders, New event).
  - Confirm the **process name** (still `olk.exe`?) and window class (`Outlook Host`?).
  - Update selectors to prefer **AutomationId** and stable named anchors over full dynamic labels when possible.

