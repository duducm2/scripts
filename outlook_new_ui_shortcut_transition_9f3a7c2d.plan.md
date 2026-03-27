---
name: Outlook New UI Shortcut Transition Plan
overview: Migrate Outlook shortcuts in `Shift keys.ahk` from legacy desktop UI selectors to New Outlook web-hosted UI selectors using the captured UI trees in `new-outlook/`. Preserve existing shortcut behavior where possible and explicitly document each shortcut that has no exact New Outlook equivalent.
todos:
  - id: collect_current_outlook_shortcuts
    content: Create a complete inventory table of all Outlook shortcuts currently defined in `Shift keys.ahk` (main, reminder, appointment, message contexts), including key combo, current behavior, and current locator strategy.
    status: pending
    dependencies: []
  - id: map_new_outlook_targets
    content: Build a mapping table from each inventory item to candidate New Outlook UI targets using `new-outlook/main 3_27_2026 1_25_48 PM.txt`, `new-outlook/calendar 3_27_2026 1_27_20 PM.txt`, `new-outlook/new-event 3_27_2026 1_30_09 PM.txt`, and `new-outlook/Reminders 3_27_2026 1_35_40 PM.txt`.
    status: pending
    dependencies: [collect_current_outlook_shortcuts]
  - id: define_matchability_status
    content: Assign each shortcut one of three statuses: Exact Match, Adapted Match, or No Reliable Match; include rationale per item and the exact UIA Name/AutomationId/ControlType evidence.
    status: pending
    dependencies: [map_new_outlook_targets]
  - id: add_new_outlook_detection_helpers
    content: Add helper predicates in `Shift keys.ahk` to detect New Outlook contexts using stable signals such as ClassName `Outlook Host`, title tokens (Inbox/Calendar/New event/Reminders), and robust UIA root checks.
    status: pending
    dependencies: [define_matchability_status]
  - id: add_new_uia_lookup_wrappers
    content: Add reusable helper wrappers in `Shift keys.ahk` that perform prioritized UIA lookup by AutomationId then Name+ControlType, with explicit fallback arrays for New Outlook labels.
    status: pending
    dependencies: [add_new_outlook_detection_helpers]
  - id: migrate_main_send_to_general
    content: Replace legacy send-sequence logic for Shift+G with New Outlook action based on `Quick steps` target (e.g., quick step names containing `Move to Gerais` or `Move to general`) in the main window context.
    status: pending
    dependencies: [add_new_uia_lookup_wrappers]
  - id: migrate_main_send_to_newsletter
    content: Replace legacy send-sequence logic for Shift+N with New Outlook quick-step or folder-target action for newsletter route, using stable names from navigation/list controls.
    status: pending
    dependencies: [add_new_uia_lookup_wrappers]
  - id: migrate_main_go_to_inbox
    content: Replace legacy Alt/typed-navigation logic for Shift+I with direct UIA selection of Inbox tree item in New Outlook Navigation pane.
    status: pending
    dependencies: [add_new_uia_lookup_wrappers]
  - id: migrate_main_focus_tabs
    content: Update Shift+F logic to toggle `Focused` and `Other` using New Outlook Message list tab items rather than old button assumptions.
    status: pending
    dependencies: [add_new_uia_lookup_wrappers]
  - id: migrate_main_mail_calendar_toggle
    content: Update Shift+M to toggle `Mail` and `Calendar` using `left-rail-appbar` toggle buttons instead of `NetUIListViewItem` lookups.
    status: pending
    dependencies: [add_new_uia_lookup_wrappers]
  - id: migrate_main_week_month_views
    content: Update Shift+W and Shift+O to select New Outlook `Week` and `Month` toggle buttons in calendar ribbon/tooling, with fallback-safe behavior.
    status: pending
    dependencies: [add_new_uia_lookup_wrappers]
  - id: migrate_reminder_actions
    content: Update reminder shortcuts: Shift+X to `Dismiss all`, and rework snooze/join logic for Shift+H/Shift+F/Shift+D/Shift+J to New Outlook reminder UI; mark any unsupported interval or join behavior when no exact element exists.
    status: pending
    dependencies: [add_new_uia_lookup_wrappers]
  - id: migrate_new_event_field_shortcuts
    content: Replace old appointment field IDs and names for Shift+S/P/T/E/H/A/I/R/L/B/C in event compose context using New Outlook controls (`Add title`, `Invite required attendees`, `Add a room or location`, `Teams meeting`, `Response options`, etc.).
    status: pending
    dependencies: [add_new_uia_lookup_wrappers]
  - id: preserve_unaffected_shortcuts
    content: Keep shortcuts that still function through generic key navigation (for example pane cycle via F6) unchanged and explicitly annotate as intentionally unchanged.
    status: pending
    dependencies: [define_matchability_status]
  - id: update_cheat_sheet_only_for_gaps
    content: Update cheat sheet sections in `Shift keys.ahk` to preserve current entries for adapted shortcuts and add a concise `Not available in New Outlook` marker only for shortcuts with No Reliable Match.
    status: pending
    dependencies: [migrate_main_send_to_general, migrate_main_send_to_newsletter, migrate_main_go_to_inbox, migrate_main_focus_tabs, migrate_main_mail_calendar_toggle, migrate_main_week_month_views, migrate_reminder_actions, migrate_new_event_field_shortcuts, preserve_unaffected_shortcuts]
  - id: add_gap_report_document
    content: Create `new-outlook/shortcut-gap-report.md` documenting every No Reliable Match shortcut with attempted target, why it failed, and proposed workaround or manual alternative.
    status: pending
    dependencies: [update_cheat_sheet_only_for_gaps]
  - id: validate_shortcut_matrix
    content: Execute a manual validation pass per Outlook window type (main, reminders, calendar, new event) and record Pass/Fail results for every shortcut in the inventory matrix.
    status: pending
    dependencies: [add_gap_report_document]
  - id: final_cleanup_and_consistency
    content: Ensure naming, comments, and helper usage in `Shift keys.ahk` are consistent, remove obsolete locator constants for legacy Outlook where replaced, and keep backward-safe fallbacks only where justified.
    status: pending
    dependencies: [validate_shortcut_matrix]
---

# Outlook New UI Shortcut Transition Plan

## Analysis / Context:
`Shift keys.ahk` currently contains Outlook shortcuts that depend on classic Outlook controls, legacy `AutomationId` values (for example `4098`, `4101`, `4226`, `4364`) and `NetUI*` class names. The New Outlook captures in `new-outlook/` show a web-hosted structure (`Outlook Host`, `BrowserRootView`, Fluent UI control names, and different IDs), so many selectors no longer resolve. The migration needs to preserve current hotkeys, adapt implementation to the new UI tree, and explicitly document unsupported mappings.

## Proposed Changes:
- Build an explicit shortcut inventory and matchability matrix before editing behavior.
- Introduce New Outlook-aware detection and UIA helper wrappers to avoid repeating brittle lookup logic.
- Migrate shortcut handlers in focused groups:
  - Main window shortcuts (`+G`, `+N`, `+I`, `+F`, `+M`, `+W`, `+O`, and other Outlook-main actions).
  - Reminder shortcuts (`+H`, `+F`, `+D`, `+X`, `+J`).
  - Event/appointment shortcuts (`+S`, `+P`, `+T`, `+E`, `+H`, `+A`, `+I`, `+R`, `+L`, `+B`, `+C`).
- Keep working generic shortcuts unchanged when they still map cleanly (for example F6-based pane cycling).
- Update cheat sheet text only where exact parity is not possible, and generate a dedicated gap report file for transparency.

## Files to Modify:
- `Shift keys.ahk`
  - Cheat sheet Outlook regions near current `cheatSheets["OUTLOOK.EXE"]`, `cheatSheets["OutlookReminder"]`, `cheatSheets["OutlookAppointment"]`, `cheatSheets["OutlookMessage"]`.
  - Reminder hotkey block near `#HotIf ... Reminder`.
  - Main Outlook hotkey block near `#HotIf IsOutlookMainActive()`.
  - Appointment hotkey block near `#HotIf IsOutlookAppointmentActive()`.
  - Shared helper region (UIA lookup and focus/click helpers).
- `new-outlook/shortcut-gap-report.md` (new file)
- Reference-only inputs (do not modify):
  - `new-outlook/main 3_27_2026 1_25_48 PM.txt`
  - `new-outlook/calendar 3_27_2026 1_27_20 PM.txt`
  - `new-outlook/new-event 3_27_2026 1_30_09 PM.txt`
  - `new-outlook/Reminders 3_27_2026 1_35_40 PM.txt`
  - `new-outlook/copilot 3_27_2026 1_28_09 PM.txt`

## Implementation Strategy:
1. Inventory every Outlook shortcut and classify by window context.
2. For each shortcut, define target evidence from the New Outlook UI trees (Name, AutomationId, ControlType, class clues).
3. Mark each shortcut as Exact Match, Adapted Match, or No Reliable Match before implementation begins.
4. Add New Outlook context-detection helpers and consolidated UIA lookup wrappers with ordered fallbacks.
5. Migrate main window shortcuts first, then reminder shortcuts, then new-event/appointment shortcuts.
6. Preserve any still-valid generic flows and mark them as intentionally retained.
7. Update cheat sheet text only for No Reliable Match shortcuts; do not alter wording for successfully adapted shortcuts.
8. Create `new-outlook/shortcut-gap-report.md` with failure rationale and fallback guidance.
9. Run a full shortcut validation matrix across all four captured window types and capture results.
10. Perform cleanup and consistency pass in `Shift keys.ahk` without removing needed compatibility fallbacks.
