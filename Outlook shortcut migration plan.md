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
- **Shift+D**: Snooze 1 day
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

#### Date/time fields

- **Shift+S**: Start date (combo)
- **Shift+P**: Start date picker
- **Shift+T**: Start time (combo)
- **Shift+E**: End date (combo)
- **Shift+H**: End time (combo)

#### Appointment fields

- **Shift+A**: All-day toggle
- **Shift+I**: Title field
- **Shift+R**: Required / To
- **Shift+L**: Location (then tab forward)
- **Shift+B**: Body (from Location → Body)

#### Recurrence / configuration

- **Shift+C**: Make recurring
- **Shift+W**: Wizard (configure)

---

## Migration execution steps (implementation checklist)

1. **Replace title/name-based branching** with a single classifier that returns exactly one of:
   - Reminders
   - Main (includes popped-up messages)
   - Appointment
2. **Move/merge “message inspector” bindings** into the Main bucket (no separate message window type).
3. **Ensure Reminders shortcuts cannot trigger** in Main/Appointment contexts (and vice-versa).
4. **Update the cheatsheet/UX surfaces** to show only the three buckets above.

