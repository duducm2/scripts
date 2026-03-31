# Outlook shortcut migration plan

## Required window distinctions (new requirement)

The previous version required window distinctions based on specific names. The current version requires primary distinctions only for the following three window types:

1. Reminders window.
2. Main Outlook window, including popped-up messages.
3. Appointment window.

This plan reframes the existing Outlook shortcuts around **only** those three window types.

---

## Mapping rules (how we regroup existing shortcuts)

- **Reminders window**: keep Reminders-only shortcuts (snooze/dismiss/join) scoped exclusively to the Reminders window.
- **Main Outlook window (including popped-up messages)**:
  - Merge current “main” and “message inspector” concepts into one bucket.
  - Any shortcuts currently defined for a separate “Message” window should be treated as part of the **Main Outlook window** bucket going forward.
- **Appointment window**: keep appointment/meeting/event inspector shortcuts scoped exclusively to the Appointment window.

---

## Shortcut inventory, regrouped by the 3 window types

### 1) Reminders window

- **Shift+H**: Snooze 1 hour
- **Shift+F**: Snooze 4 hours
- **Shift+T**: Snooze 10 minutes
- **Shift+Y**: Snooze 1 day
- **Shift+W**: Snooze 1 week
- **Shift+D**: Dismiss reminder (selected item)
- **Shift+X**: Dismiss all reminders
- **Shift+J**: Join Online

---

### 2) Main Outlook window (including popped-up messages)

#### Navigation / triage / view

- **Shift+I**: Go to Inbox
- **Shift+F**: Toggle Focused / Other
- **Shift+M**: Toggle Mail / Calendar
- **Shift+W**: Calendar Week view
- **Shift+O**: Calendar Month view
- **Shift+P**: Pop Out current item
- **Shift+K**: Cycle backward pane (send Shift+F6)
- **Shift+L**: Cycle forward pane (send F6)

#### Message field focus (applies to both main + popped-up messages)

- **Shift+S**: Subject / Title
- **Shift+T**: To / Required
- **Shift+B**: Body (tab from Subject → Body)

#### Responses

- **Shift+D**: “Don’t send any response”
- **Shift+E**: “Send response”

#### Categorize / move

- **Shift+G**: Send to General
- **Shift+N**: Send to Newsletter

---

### 3) Appointment window

#### General layout changes (new assessment)

- The classic “Date Picker” button/flow is no longer a standalone concept. The previous `Shift+P` “date picker” shortcut is removed.
- The date/time controls are now surfaced via a compact header area and a clock/calendar affordance that opens a **popover/context menu** (UIA `Window`/`dialog`) containing:
  - Start date (combo)
  - Start time (combo)
  - End time (combo)
  - Time zone
  - All day toggle
  - Make recurring toggle/button
  - Time suggestions list (list items)
- Any shortcut that targets items inside this popover must be resilient:
  - If popover is closed: open it first.
  - If popover is already open: reuse it and focus the target control.

#### Context menu / popover (clock icon field)

- **Shift+S**: Start date (combo) (in popover)
- **Shift+T**: Start time (combo) (in popover)
- **Shift+E**: End time (combo) (in popover)
- **Shift+A**: All day toggle (in popover)
- **Shift+C**: Recurring / Make recurring toggle (in popover) (new: replaces “Make recurring” classic flow)
- **(new mnemonic)**: Time suggestions — select 1st suggestion
- **(new mnemonic)**: Time suggestions — select 2nd suggestion

#### Appointment fields

- **Shift+I**: Title field
- **Shift+R**: Required / To
- **Shift+B**: Body (from Location → Body)

#### Confirmed working (no changes)

- **Shift+I**: Title field
- **Shift+R**: Required attendees
- **Shift+L**: Location/Add a room

#### Requires fixes/updates

- **Shift+B**: Focus the main message body (large empty text field) (currently broken)
- **Shift+W**: Wizard (configure) — should use Standard Information Display modals to apply presets quickly

#### New shortcuts (top command bar)

- **(new mnemonic)**: Create a Teams meeting
- **(new mnemonic)**: Series (recurring) (equivalent to recurring toggle)
- **(new mnemonic)**: Status/Busy selection modal (Free / Working elsewhere / Tentative / Busy / Out of office)
- **(new mnemonic)**: Reminder selection modal:
  - Don't remind me
  - 15 minutes before
  - 1 hour before
  - 12 hours before
  - 1 day before
  - 1 week before
- **(new mnemonic)**: Category selection modal (Aniversário / Importante / Pessoal)
- **(new mnemonic)**: Private toggle (Private / Not private)

#### Right sidebar (Attendee schedules)

- **Shift+L**: Previous/Back day (navigate schedule arrows) (per updated requirement)
- **Shift+K**: Next/Forward day (navigate schedule arrows) (per updated requirement)
- **(new mnemonic)**: Today button
- **(new mnemonic)**: Current date (button showing date/month/year)
- **(new mnemonic)**: Scheduling assistant button (opens Scheduling Assistant view)

#### Scheduling Assistant view (separate sub-mode)

- **(new mnemonic)**: Back button
- **(new mnemonic)**: Options menu (show all options)
- **Shift+L**: Previous/Back (time suggestions navigation fallback)
- **Shift+K**: Next/Forward (time suggestions navigation fallback)
- **(new mnemonic)**: Start date
- **(new mnemonic)**: Start time
- **(new mnemonic)**: End time
- **(new mnemonic)**: All day toggle
- **(new mnemonic)**: Time zone
- **(new mnemonic)**: Add required attendee
- **(new mnemonic)**: Add optional attendee

---

## Migration execution steps (implementation checklist)

1. **Replace title/name-based branching** with a single classifier that returns exactly one of:
   - Reminders
   - Main (includes popped-up messages)
   - Appointment
2. **Move/merge “message inspector” bindings** into the Main bucket (no separate message window type).
3. **Ensure Reminders shortcuts cannot trigger** in Main/Appointment contexts (and vice-versa).
4. **Update the cheatsheet/UX surfaces** to show only the three buckets above.

