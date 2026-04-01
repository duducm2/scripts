# Outlook main window shortcut mapping plan (New Outlook)

This document is the working plan to map, validate, and migrate shortcuts for the **main Outlook window** (the unified shell that hosts **Mail**, **Calendar**, and **Message/Reading Pane** views) for **New Outlook (Monarch)**.

Primary UIA context sources:

- `C:\Users\fie7ca\Documents\scripts\outlook-mail.md`
- `C:\Users\fie7ca\Documents\scripts\outlook-caledar.md`
- The attached screenshots (Mail + Calendar, showing the unified shell and the top overlay banners)

## Goals

- Provide a **complete shortcut map** for the unified main Outlook window (Mail + Calendar + Reading Pane).
- **Preserve existing Shift shortcuts** where they still make sense and can be supported in New Outlook.
- Expand capacity using the agreed overflow modifier: **Shift first, then Ctrl+Alt**.
- Migrate away from Classic-only mechanisms (e.g., ribbon Alt sequences) and prefer **UIA** against New Outlook’s WebView surface.
- Ensure shortcuts are **view-aware** (Mail vs Calendar vs Message inspector) and do not trigger on inspectors (New event / Message compose).

## Non-goals (for this pass)

- Re-designing the Appointment (New event) inspector shortcuts (already handled elsewhere).
- Implementing a brand-new shortcut runtime/engine framework. This plan stays within the current `Shift keys.ahk` architecture.

## Constraints and conventions

- **Modifier policy** (confirmed): use **Shift** as primary layer; when Shift capacity runs out, use **Ctrl+Alt** as overflow layer in the **same conceptual key order**.
- **Backward compatibility policy** (confirmed): preserve existing mappings unless impossible or broken in New Outlook.
- Prefer stable UIA selectors:
  - **AutomationId** when present.
  - **Name + ControlType** when stable (e.g., left-rail buttons “Mail”, “Calendar”).
  - Avoid dynamic labels that include dates/times or message subjects.
- Avoid expensive desktop-wide UIA searches; scope to active window (`UIA.ElementFromHandle(WinExist("A"))`) and use known anchors.

## What “main Outlook window” means (target surface)

From `outlook-mail.md` / `outlook-caledar.md`:

- Top-level window:
  - **ClassName**: `Outlook Host`
  - **Document**: `AutomationId: "RootWebArea"`
- Cross-view shared anchors:
  - **Top search**: `AutomationId: "topSearchInput"` (ComboBox)
  - **Left rail app bar**: group `Name: "left-rail-appbar"` containing toggle buttons:
    - “Mail” (`Button`, toggle)
    - “Calendar” (`Button`, toggle)
  - **My Day**: `AutomationId: "owaTimePanelBtn_container"`
  - **Notifications**: `AutomationId: "owaActivityFeedButton_container"`
  - **Settings**: `AutomationId: "owaSettingsBtn_container"`
- Mail view key regions:
  - **Message list region**: `AutomationId: "Skip to message list-region"`
  - **Reading pane region**: `AutomationId: "Skip to message-region"`
  - Mail list controls:
    - **Filter**: `AutomationId: "mailListFilterMenu"`
    - **Sort**: `AutomationId: "mailListSortMenu"`
- Calendar view key regions:
  - View mode toggles have known AutomationIds in capture (examples):
    - “Week”: `AutomationId: "2519"`
    - “Month”: `AutomationId: "2505"`
    - “Work week”: `AutomationId: "2520"` (shown in UIA label as `Ctrl+Alt+2`)
  - Date selector/navigation:
    - “Go to previous month …”, “Go to next month …” (Buttons)
    - Date grid (DataGrid) and day buttons (Buttons)

## Existing shortcut inventory (current configuration)

This is the baseline to preserve/migrate.

### Global Outlook launchers (Win+Alt+Shift…)

From `docs/outlook-shortcuts-and-debug-notes.md`:

- Win+Alt+Shift+B: Open Mail
- Win+Alt+Shift+G: Open Calendar
- Win+Alt+Shift+V: Open Reminders

### Shift layer — main Outlook window (current)

From `Shift keys.ahk` cheat sheet `cheatSheets["OUTLOOK.EXE"]` and `#HotIf IsOutlookMainActive()` handlers:

- Shift+G: Move to General folder (New Outlook UIA path, classic fallback via Alt ribbon sequence)
- Shift+N: Move to Newsletter folder (New Outlook UIA path, classic fallback via Alt ribbon sequence)
- Shift+I: Go to Inbox (New Outlook UIA path, classic fallback)
- Shift+S: Focus Subject (where applicable)
- Shift+T: Focus Required/To (where applicable)
- Shift+B: Subject → Body (tab from Subject)
- Shift+F: Toggle Focused/Other
- Shift+K / Shift+L: Pane cycling via Shift+F6 / F6
- Shift+M: Toggle Mail/Calendar (New Outlook uses left-rail buttons by Name)
- Shift+W: Calendar Week view (New Outlook: `AutomationId: "2519"`)
- Shift+O: Calendar Month view (New Outlook: `AutomationId: "2505"`)
- Shift+P: Pop Out current item (needs re-validation in New Outlook; likely Name-based “Open in new window” vs “Pop Out”)
- Shift+D / Shift+E: Response flows (currently keyboard-driven; needs New Outlook confirmation)

### Shift layer — message inspector (compose) (current)

From `cheatSheets["OutlookMessage"]` and `#HotIf IsOutlookMessageActive()`:

- Shift+S: Subject
- Shift+T: To / Required
- Shift+B: Subject → Body

## Compatibility gaps found during review (needs migration plan)

- `IsOutlookMainActive()` currently gates on `WinActive("ahk_exe OUTLOOK.EXE")` and title heuristics. New Outlook can run as `olk.exe`. We should align gating with `IsNewOutlookActive()` / `Outlook Host` class.
- Some existing actions still rely on classic sequences (Alt ribbon navigation). These must be replaced with UIA equivalents (or explicitly marked “classic-only” and moved behind a classic-only gate).
- The main-window shortcut set today is **not comprehensive** for Mail triage, navigation, and Calendar navigation. The overflow layer (Ctrl+Alt) must fill these gaps.

**Implementation status (this repo)**:

- `IsOutlookMainActive()` now accepts **`olk.exe`** and prefers the **`Outlook Host`** shell characteristics.
- Added a main-window activation helper that targets the `Outlook Host` window titled `Mail - … - Outlook` or `Calendar - … - Outlook`.
- Added a first batch of **Ctrl+Alt** overflow shortcuts (Mail triage + focus + view switching + search), implemented via UIA anchors present in `outlook-mail.md` / `outlook-caledar.md`.

## Proposed overall shortcut architecture (high level)

### 1) Focus/activation: always anchor to the main window first

Add a dedicated “bring main Outlook shell to front” helper used by all main-window hotkeys:

- Locate/activate any visible New Outlook window with:
  - `ahk_class Outlook Host`
  - title containing `Mail -` or `Calendar -` (or ` - Outlook`)
  - process `olk.exe` or `OUTLOOK.EXE` (depending on install)

This ensures shortcuts work even when a child region has focus (reading pane, message list, nav tree).

### 2) View detection: decide which mapping applies

Use fast heuristics:

- **Mail view** if the main window contains `AutomationId: "Skip to message list-region"` and/or a document URL containing `/mail/`.
- **Calendar view** if the document URL contains `/calendar/` or if calendar view toggles are present.

Shortcuts should be:

- **global-in-main-window** (work in both Mail and Calendar)
- **mail-only**
- **calendar-only**

### 3) Two-layer keyspace plan

- **Layer A (primary)**: Shift+Key — preserve current assignments; add only if there’s still room and no conflict.
- **Layer B (overflow)**: Ctrl+Alt+Key — new mappings for coverage.

Key selection principles:

- Prefer mnemonic letters (e.g., `Ctrl+Alt+R` = Reply).
- Prefer consistency across views (same key triggers analogous action in Mail/Calendar where reasonable).
- Avoid collisions with global Windows/system shortcuts.

## Shortcut mapping backlog (what “comprehensive” means)

This list defines the full scope we will cover (mapped either via Shift or Ctrl+Alt).

### Global (Mail + Calendar)

- App focus/switching:
  - Focus main Outlook window
  - Toggle Mail/Calendar
  - Focus top search
  - Open Settings / Notifications / My Day
- Layout:
  - Toggle reading pane / split view (if exposed)
  - Switch layouts (ribbon mode toggle exists in UIA in both captures)

### Mail view — core

- Navigation:
  - Focus navigation pane (folders)
  - Focus message list
  - Focus reading pane
  - Go to Inbox / Sent / Drafts / Archive (at least Inbox preserved; others optional but recommended)
  - Focused/Other toggle (preserved)
- Triage actions on selected message/conversation:
  - Reply / Reply all / Forward
  - Delete / Archive
  - Mark as read/unread
  - Flag / Unflag
  - Categorize (open category menu + pick commonly used categories)
  - Move (open move menu, optionally quick targets)
  - Pin (if available) / Snooze
  - Open in new window
- List controls:
  - Filter menu
  - Sort menu
  - Jump to / Select (if useful)

### Calendar view — core

- Create:
  - New event
  - (Optional) new meeting / new appointment variants if distinct
- View modes:
  - Day / Work week / Week / Month (Week/Month preserved)
  - Today
  - Next/previous day/week (depending on current view)
- Navigation:
  - Open date selector / jump to date
  - Next/previous month in date selector (buttons exist)
  - Focus the calendar grid
- Event actions (in main Calendar surface, not the “New event” inspector):
  - Open selected event
  - Delete / categorize / private (if available at this surface)

## Proposed mappings (draft)

This section is the “first-pass” mapping proposal. It will be refined after checking for collisions and verifying UIA selectors.

### Preserve existing Shift mappings (as-is)

- Shift+G: Move to General
- Shift+N: Move to Newsletter
- Shift+I: Go to Inbox
- Shift+F: Toggle Focused/Other
- Shift+M: Toggle Mail/Calendar
- Shift+W: Week view (Calendar)
- Shift+O: Month view (Calendar)
- Shift+K / Shift+L: Cycle panes (F6 / Shift+F6)
- Shift+P: Pop out / open in new window (re-validate label in New Outlook)
- Shift+S / Shift+T / Shift+B: field focus helpers (message/compose contexts)

### New Ctrl+Alt mappings (overflow layer)

Global:

- Ctrl+Alt+F: Focus top Search (`AutomationId: "topSearchInput"`)
- Ctrl+Alt+M: Force switch to Mail (left rail “Mail” toggle)
- Ctrl+Alt+G: Force switch to Calendar (left rail “Calendar” toggle)
- Ctrl+Alt+, / Ctrl+Alt+. : Previous/Next major navigation (reserved for later if needed)

Mail view:

- Ctrl+Alt+R: Reply (Reading Pane toolbar/menu item “Reply”)
- Ctrl+Alt+A: Reply all
- Ctrl+Alt+W: Forward
- Ctrl+Alt+D: Delete (Ribbon “Delete” button when present; otherwise reading pane “More items” menu path)
- Ctrl+Alt+E: Archive
- Ctrl+Alt+U: Mark read/unread
- Ctrl+Alt+C: Categorize
- Ctrl+Alt+V: Move (note: Outlook already advertises Ctrl+Shift+V for “Move to folder”; we keep Ctrl+Alt+V for our automation layer only if it doesn’t conflict in practice)
- Ctrl+Alt+O: Open in new window (or pop out)
- Ctrl+Alt+L: Focus message list (`AutomationId: "Skip to message list-region"`)
- Ctrl+Alt+P: Focus reading pane (`AutomationId: "Skip to message-region"`)
- Ctrl+Alt+N: New mail (if a stable “New” button exists in mail ribbon; verify in UIA capture beyond the excerpt)
- Ctrl+Alt+H: Toggle Focused/Other (optional duplicate; only if it helps ergonomics)
- Ctrl+Alt+J: Jump to (button exists in message list header; verify usefulness)
- Ctrl+Alt+I: Filter menu (`AutomationId: "mailListFilterMenu"`)
- Ctrl+Alt+S: Sort menu (`AutomationId: "mailListSortMenu"`)

Calendar view:

- Ctrl+Alt+N: New event (Calendar ribbon “New event” button)
- Ctrl+Alt+1/2/3/4: Day / Work week / Week / Month view (only if stable AutomationIds exist; Week/Month already on Shift so these may be optional)
- Ctrl+Alt+T: Today
- Ctrl+Alt+Left / Ctrl+Alt+Right: Previous/Next day or week (to be defined per view mode)
- Ctrl+Alt+D: Day view toggle (if we avoid numeric mapping)
- Ctrl+Alt+W: Work week view toggle (if distinct)
- Ctrl+Alt+K / Ctrl+Alt+L: Previous/Next month in date selector (buttons exist in date selector nav)

**Implemented now (Ctrl+Alt)**:

- Global:
  - Ctrl+Alt+F: focus Search
  - Ctrl+Alt+M: switch to Mail
  - Ctrl+Alt+G: switch to Calendar
- Mail:
  - Ctrl+Alt+L: focus message list region
  - Ctrl+Alt+P: focus reading pane region
  - Ctrl+Alt+I: open Filter menu (`mailListFilterMenu`)
  - Ctrl+Alt+S: open Sort menu (`mailListSortMenu`)
  - Ctrl+Alt+R/A/W: Reply / Reply all / Forward (prefers reading-pane command surface)
  - Ctrl+Alt+D/E: Delete / Archive (prefers ribbon AutomationIds `519` / `505`)
  - Ctrl+Alt+U: Read/Unread (AutomationId `552`)
  - Ctrl+Alt+C: Categorize (AutomationId `509`)
  - Ctrl+Alt+V: Move (AutomationId `540`)
- Calendar:
  - Ctrl+Alt+N: New event if available; otherwise falls back to built-in `Ctrl+N` for Mail compose
  - Ctrl+Alt+T: Today (Name-based; needs confirmation that a stable Today button exists in your Calendar surface)

## Implementation plan (phased)

### Phase 0 — inventory and collision audit

- Enumerate all existing Outlook-related hotkeys across:
  - Main window (`IsOutlookMainActive()`)
  - Message inspector (`IsOutlookMessageActive()`)
  - Appointment inspector (`IsOutlookAppointmentActive()`)
  - Global launchers (`Outlook.ahk`)
- Identify collisions with:
  - Built-in Outlook shortcuts
  - Windows global shortcuts
  - Existing repo-wide hotkeys

### Phase 1 — make the “main window” gate robust for New Outlook

- Update/replace `IsOutlookMainActive()` logic so New Outlook is reliably detected:
  - include `olk.exe`
  - include `Outlook Host` class
  - exclude “New event”, “Message (compose)”, and “Reminders” windows by title/class where possible

### Phase 2 — add main-window focus helpers

- Implement helpers that:
  - Activate the main Outlook window
  - Focus major regions (navigation pane, message list, reading pane, calendar grid)
  - Focus top search reliably

### Phase 3 — migrate any remaining classic-only sequences

- Replace remaining `Send "!5" ...` style flows for actions that are expected to work in New Outlook with UIA equivalents.
- Keep classic-only fallbacks behind a strict classic-only gate (if still desired).

### Phase 4 — implement Ctrl+Alt overflow mappings

- Implement the proposed Ctrl+Alt mapping set.
- Each action must:
  - Activate the main window first
  - Verify expected view or fallback gracefully
  - Use UIA anchors in `outlook-mail.md` / `outlook-caledar.md`

### Phase 5 — validate against the UI structure files and iterate

- Validate each mapping against:
  - Control existence (AutomationId/Name)
  - Stability across Mail/Calendar
  - Speed + reliability (avoid deep tree traversals)

## Test plan (manual, pragmatic)

- Mail:
  - With focus in each region (nav tree / message list / reading pane), trigger each Mail mapping and confirm correct target action.
  - With no message selected, confirm actions fail safely (no random clicks).
- Calendar:
  - From week view and month view, validate view toggles, Today, and navigation.
  - Validate “New event” opens (main surface), and shortcuts do not accidentally apply to the inspector unless explicitly intended.
- Cross-view:
  - From Mail, switch to Calendar and back using both Shift+M and the dedicated Ctrl+Alt toggles.

## Open items (need a short follow-up capture or confirmation)

- Confirm the New Outlook Mail ribbon contains a stable “New mail” button (name/AutomationId), or decide to rely on existing built-in `Ctrl+N`.
- Confirm whether “Pop Out” is still labeled as “Open in new window” in New Outlook reading pane toolbar.
- Confirm whether message triage buttons (Delete/Archive/Move/Categorize) are consistently present in the reading pane toolbar vs the top ribbon, to choose the most stable anchor.

