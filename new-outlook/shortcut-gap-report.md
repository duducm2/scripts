# New Outlook Shortcut Gap Report

This report tracks Outlook shortcuts from `Shift keys.ahk` that currently have **No Reliable Match** in New Outlook based on the captured UI trees.

## Main Window

- `Shift+N` (Send to Newsletter)
  - Attempted target: Quick steps item containing `newsletter`.
  - Why it failed: `new-outlook/main 3_27_2026 1_25_48 PM.txt` only shows `Move to Gerais` / `Move to general`; no stable newsletter target was captured.
  - Workaround: Keep legacy fallback sequence (`Alt+5`, `O`, `01`, Enter) for now.

## Reminders

- `Shift+H` (Snooze 1 hour)
- `Shift+F` (Snooze 4 hours)
- `Shift+D` (Snooze 1 day)
  - Attempted target: reminder window snooze controls / snooze picker.
  - Why it failed: `new-outlook/Reminders 3_27_2026 1_35_40 PM.txt` exposes reminder items plus `Dismiss all`, but no reliable Snooze button/combobox target in this capture.
  - Workaround: keep existing keyboard-based flow, marked as not available for New Outlook reliability.

- `Shift+J` (Join Online)
  - Attempted target: `Join Online` button in reminders.
  - Why it failed: no `Join Online` control appears in the New Outlook reminders tree capture.
  - Workaround: open the reminder/event and join from the event surface manually.

## New Event / Appointment

- `Shift+S` (Start date combo)
- `Shift+P` (Start date picker)
- `Shift+T` (Start time combo)
- `Shift+E` (End date combo)
- `Shift+H` (End time combo)
  - Attempted target: legacy date/time field AutomationIds (`4098`, `4096`, `4099`, `4097`, `4352`, `4353`).
  - Why it failed: New Outlook event compose exposes a combined date/time range button (`Fri 3/27/2026 2:00 PM - 2:30 PM`) instead of stable separate field controls.
  - Workaround: manual adjustment in the date/time range button until a stable automation pattern is captured.

- `Shift+A` (All-day toggle)
  - Attempted target: classic `All day` checkbox (`AutomationId: 4226` / name `All day`).
  - Why it failed: no matching all-day checkbox is exposed in the captured New Outlook event tree.
  - Workaround: manual toggle in event options.

- `Shift+C` (Make Recurring)
  - Attempted target: classic `Make Recurring` button (`AutomationId: 4364`).
  - Why it failed: no `Make Recurring` button is exposed in the captured New Outlook event tree.
  - Workaround: use recurrence controls from the event form UI manually.
