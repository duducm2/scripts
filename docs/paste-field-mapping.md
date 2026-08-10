# Paste field mapping (Win+Alt+Shift+L)

Learn-and-persist **main text fields** for windows reached via the visible-window paste picker. After the first successful paste into a target app, you can save the focused UIA field; later pastes to the same process + title needle focus that field before `Ctrl+V`.

## Purpose

| Entry                        | Behavior                                                                                |
| ---------------------------- | --------------------------------------------------------------------------------------- |
| **Win+Alt+Shift+L** (`#!+L`) | Paste OS clipboard (`^v`) to a picked visible window (same as D2C submit menu **[W]**). |
| Mapped windows               | Before paste, focus the saved main text field via UIA.                                  |
| Unmapped windows             | Paste blindly, then ask whether to save the focused field as the main field.            |

Core logic: [`Utils/paste_field_mapping.ahk`](../Utils/paste_field_mapping.ahk). Wired from [`Utils/d2c_flow_manager.ahk`](../Utils/d2c_flow_manager.ahk) (`PasteClipboardToVisibleWindow`). Hotkey: [`Utils/utility_shortcuts.ahk`](../Utils/utility_shortcuts.ahk). Included from [`Utils.ahk`](../Utils.ahk) before the D2C module.

Dictation **[W]** uses the same deferred focus+paste path as `#!+L` so INI-mapped main fields apply reliably after leaving the Send dictation? overlay.

## How to use

### First time (learn)

1. Put text on the clipboard.
2. In the target app, click (or Tab to) the **intended** compose / prompt field so it has keyboard focus.
3. Press **Win+Alt+Shift+L**, then press the **slot key** on that window in the picker.
4. Content pastes. If there is no saved mapping for that window, an Interactive Input banner appears:
   - `❓ Set this text field as the main field for {label}?`
   - Optional second line: the focused field’s UIA `Name` (shortened).
5. **[Y]** — save mapping to INI and show `✅ Main text field saved for …`
6. **[N]**, **Esc**, or ~5s timeout — skip; no file write.

Learning captures whatever has focus **after** paste (signature is taken before the banner steals focus). If the wrong control was focused, answer **N**, focus the right field, paste again, and answer **Y**.

### Later (mapped)

1. Clipboard → **Win+Alt+Shift+L** → slot key.
2. Script matches **exe + title needle** in the INI, focuses the stored UIA field, then `^v`.
3. No Y/N prompt when a mapping already exists.

### Picker controls (unchanged)

- **Slot key** — select window and paste.
- **[R]** then slot — AutoSlot user ignore (process exe); not field learning.
- **[I]** — manage AutoSlot ignore list.
- **Esc** — cancel picker.

## Persistence

File (created on first **Y**): [`assets/data/paste_field_mappings.ini`](../assets/data/paste_field_mappings.ini)

Example section:

```ini
[Mapping_1]
Exe=chrome.exe
TitleNeedle=gemini
Name=Enter a prompt for Gemini
AutomationId=
ClassName=ql-editor
Type=50004
UiaAttach=browser
```

| Key                                            | Role                                                                                                                                                        |
| ---------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `Exe`                                          | Process name match (case-insensitive).                                                                                                                      |
| `TitleNeedle`                                  | Optional; window title must contain this (`InStr`, case-insensitive). Derived at learn time (browser suffixes stripped; prefers segment before `-` / `\|`). |
| `Name` / `AutomationId` / `ClassName` / `Type` | UIA field signature. Lookup prefers AutomationId, then Name+Type, then ClassName+Type.                                                                      |
| `UiaAttach`                                    | `browser` (`UIA_Browser`) for Chromium-style windows; otherwise `element` (`UIA.ElementFromHandle`).                                                        |

First matching `[Mapping_N]` wins. Sections are numbered sequentially (`Mapping_1`, `Mapping_2`, …).

### Re-learn

Delete the relevant `[Mapping_N]` block (or the whole file), then **reload Utils** (or restart the host script) so the in-memory mapping cache refreshes. Paste again with the correct field focused and answer **Y**.

## UI standard

- Learn prompt: **Interactive Input** — `StandardLoadingBar_ShowWithKeys` with `promptKeys` `[Y] Yes  [N] No`, `BANNER_ACCENT_INTERMEDIATE`.
- Save success: **Information Only** — `ShowCenteredOverlay_Utils` + `BANNER_ACCENT_SUCCESS`.

Details: [standard_information_display.md](standard_information_display.md).

## Related

- [dictation-to-gemini-cursor-flow.md](dictation-to-gemini-cursor-flow.md) — D2C flow; `#!+L` / menu **[W]** entry.
- Cheat sheet GLOBAL string: [`Shift keys/cheat_sheet_registry.ahk`](../Shift%20keys/cheat_sheet_registry.ahk) (`[Win+Alt+Shift+L]`).
