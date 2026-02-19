# Shortcut audit: Win+Alt+Shift+X

## Scope

Audit searched the following for `#!+x` / `Win+Alt+Shift+X`:

- Shift keys.ahk
- Utils.ahk
- AppLaunchers.ahk
- Gemini.ahk
- Microsoft Teams.ahk
- Outlook.ahk
- Spotify.ahk
- WindowManagement.ahk
- Grep over `*.ahk` for `#!+x` and `Win+Alt+Shift+X`

## Result

**Available.** As of the follow-up change, Win+Alt+Shift+X is no longer bound. All code related to `#!+x` (Hunt and Peck hotkey and helpers) was removed from Utils.ahk; the cheatsheet was updated to mark the shortcut as available.

## References (historical)

Previously: Utils.ahk had hotkey `#!+x::` and Hunt and Peck logic (ActivateHuntAndPeck, CloseHuntAndPeckProcess, ScheduleHnPCleanup, etc.). Shift keys.ahk cheatsheet listed `[Win+Alt+Shift+X] > Activate hunt and Peck`. env.ahk contained `HNP_EXE_PATH_*` and `GetHnPExePath()` (still present; only Utils.ahk’s Hunt and Peck block and its `#Include` of env.ahk were removed).

## Action

Done. Shortcut is available for new assignments. Cheatsheet entry: `[Win+Alt+Shift+X] > (available)`.
