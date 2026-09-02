# Paste field mapping (Win+Alt+Shift+L)

Learn-and-persist **main text fields** for windows reached via the visible-window paste picker. After the first successful paste into a target app, you can save the focused UIA field; later pastes to the same process + title/URL needle focus that field before `Ctrl+V`.

## Purpose

| Entry                        | Behavior                                                                                                                                                                                |
| ---------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Win+Alt+Shift+L** (`#!+L`) | Paste OS clipboard (`^v`) to a picked visible window (same as D2C submit menu **[W]**). After pick: **Y** = paste+Enter, **N** = paste only, **Esc** = abort, ~3s timeout = paste only. |
| Mapped windows               | Before paste, focus the saved main text field via UIA.                                                                                                                                  |
| Unmapped windows             | Paste blindly, then ask whether to save the focused field as the main field.                                                                                                            |

Core logic: [`Utils/paste_field_mapping.ahk`](../Utils/paste_field_mapping.ahk). Wired from [`Utils/d2c_flow_manager.ahk`](../Utils/d2c_flow_manager.ahk) (`PasteClipboardToVisibleWindow`, post-pick auto-send banner). Hotkey: [`Utils/utility_shortcuts.ahk`](../Utils/utility_shortcuts.ahk). Included from [`Utils.ahk`](../Utils.ahk) before the D2C module.

Dictation **[W]** uses the same deferred focus+paste path as `#!+L` so INI-mapped main fields apply reliably after leaving the Send dictation? overlay.

## How to use

### After window pick (auto-send banner)

Immediately after you press a **slot key** in the picker (before focus/paste), an Interactive Input banner appears:

- `❓ Paste to window? (3s)`
- **[Y]** — paste (`^v`) then **Enter**
- **[N]** — paste only (no Enter); same outcome as waiting out the timeout
- **[Esc]** — abort; no paste (distinct from picker **Esc**, which cancels the picker before a window is chosen)
- ~3s timeout — paste only (no Enter)

### First time (learn)

1. Put text on the clipboard.
2. In the target app, click (or Tab to) the **intended** compose / prompt field so it has keyboard focus.
3. Press **Win+Alt+Shift+L**, then press the **slot key** on that window in the picker.
4. Answer the auto-send banner (**Y** / **N** / **Esc** / timeout). If paste proceeds and there is no saved mapping for that window, an Interactive Input banner appears:
   - `❓ Set this text field as the main field for {label}?`
   - Optional second line: the focused field’s UIA `Name` (shortened).
   - For Chromium windows, `{label}` prefers the page **host+path** URL needle when available.
5. **[Y]** — save mapping to INI and show `✅ Main text field saved for …`
6. **[N]**, **Esc**, or ~5s timeout — skip; no file write.

Learning captures whatever has focus **after** paste (signature is taken before the banner steals focus). If the wrong control was focused, answer **N**, focus the right field, paste again, and answer **Y**.

### Later (mapped)

1. Clipboard → **Win+Alt+Shift+L** → slot key → auto-send banner (**Y** / **N** / timeout).
2. Script matches the INI row (see [Match rules](#match-rules)), focuses the stored UIA field, then `^v` (and **Enter** if you pressed **Y**).
3. No field-learn Y/N prompt when a mapping already exists.

### Picker controls (unchanged)

- **Slot key** — select window (then auto-send banner).
- **[R]** then slot — AutoSlot user ignore (process exe); not field learning.
- **[I]** — manage AutoSlot ignore list.
- **[M]** — manage/remove main text-field mappings.
- **Esc** — cancel picker (before a window is chosen).

## Persistence

File (created on first **Y**): [`assets/data/paste_field_mappings.ini`](../assets/data/paste_field_mappings.ini)

Example section (browser):

```ini
[Mapping_1]
Exe=chrome.exe
TitleNeedle=Gemini
UrlNeedle=gemini.google.com/app
Name=Enter a prompt for Gemini
AutomationId=
ClassName=ql-editor
Type=50004
UiaAttach=browser
```

Example section (non-browser):

```ini
[Mapping_2]
Exe=Cursor.exe
TitleNeedle=research.md
UrlNeedle=
Name=
AutomationId=
ClassName=aislash-editor-input
Type=50004
UiaAttach=browser
```

| Key                                            | Role                                                                                                                                                                                                      |
| ---------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `Exe`                                          | Process name match (case-insensitive).                                                                                                                                                                    |
| `TitleNeedle`                                  | Optional; window title must contain this (`InStr`, case-insensitive). Derived at learn time (browser suffixes stripped; prefers segment before `-` / `\|`). Used for non-browser and legacy browser rows. |
| `UrlNeedle`                                    | Optional; host + path (lowercase), no scheme/query/hash, no trailing `/`. Set at learn time for Chromium windows via `UIA_Browser.GetCurrentURL`. Primary match key for browsers.                         |
| `Name` / `AutomationId` / `ClassName` / `Type` | UIA field signature. Lookup prefers AutomationId, then Name+Type, then ClassName+Type.                                                                                                                    |
| `UiaAttach`                                    | `browser` (`UIA_Browser`) for Chromium-style windows; otherwise `element` (`UIA.ElementFromHandle`).                                                                                                      |

First matching `[Mapping_N]` wins (for URL matches, the **longest** matching `UrlNeedle` wins). Sections are numbered sequentially (`Mapping_1`, `Mapping_2`, …).

### Match rules

1. Same `Exe` only.
2. **Browser hwnd** (Chromium class): if the current page URL needle can be captured, match rows whose `UrlNeedle` is a path-prefix of the current page (prefer longest). If none match, fall back to **TitleNeedle** only for legacy rows with **empty** `UrlNeedle`.
3. **Non-browser**, or browser when URL capture fails: **TitleNeedle** gate (empty needle matches any title for that exe).
4. **Pass 2** (strict UIA identity when title/url gate missed):
   - Non-browser: unchanged (AutomationId or ClassName+Type).
   - Browser: only rows whose `UrlNeedle` matches the current page. Legacy browser rows with empty `UrlNeedle` are **excluded** from Pass 2 so one Chrome field cannot apply site-wide.

### Re-learn

Delete the relevant `[Mapping_N]` block via picker **[M]** (or edit the INI), then **reload Utils** (or restart the host script) so the in-memory mapping cache refreshes. Paste again with the correct field focused and answer **Y**. Legacy Chrome rows without `UrlNeedle` keep title-only matching until re-learned (which stores `UrlNeedle`).

## UI standard

- Picker scan (before grid modal): **Loading Indication** — `StandardLoadingBar_Show` `⏳ Scanning visible windows...` during `Dictation_BuildMonitorGrid`, then `Hide(0)` before the picker or the no-windows overlay (`Dictation_ShowVisiblePasteSelector` in `dictation_visible_paste.ahk`).
- Post-pick automation (after auto-send choice): **Loading Indication** — `Show` → `Update` → `Hide` in `_FinishDeferredPaste` (`d2c_flow_manager.ahk`): `⏳ Activating window...`, `⏳ Focusing main field...` (when mapped), `⏳ Pasting...`, `⏳ Sending...` (when **Y**); bar hides before the learn prompt.
- Auto-send (post-pick): **Interactive Input** — `StandardLoadingBar_ShowWithKeys` with `promptKeys` `[Y] Send after paste  [N] Paste only  [Esc] Cancel`, `BANNER_ACCENT_INTERMEDIATE` (see `PasteWindow_ShowAutoSendOptionsAndWait` in `d2c_flow_manager.ahk`).
- Learn prompt: **Interactive Input** — `StandardLoadingBar_ShowWithKeys` with `promptKeys` `[Y] Yes  [N] No`, `BANNER_ACCENT_INTERMEDIATE`.
- Save success: **Information Only** — `ShowCenteredOverlay_Utils` + `BANNER_ACCENT_SUCCESS`.

Details: [standard_information_display.md](standard_information_display.md).

## Manual verification

1. `#!+L` → pick window → **N** pastes without Enter; **Esc** aborts; **Y** still paste+Enter; timeout still paste-only.
2. On Chrome site A, learn a field; on site B (different host/path) the mapping must not auto-focus; learn separately.
3. Revisit site A: saved field focuses before paste.
4. Non-browser (e.g. Cursor): title-based mapping still works.
5. Manage UI **[M]** lists/deletes rows; URL-backed rows show path in the Title / URL column.

## Related

- [dictation-to-gemini-cursor-flow.md](dictation-to-gemini-cursor-flow.md) — D2C flow; `#!+L` / menu **[W]** entry.
- Cheat sheet GLOBAL string: [`Shift keys/cheat_sheet_registry.ahk`](../Shift%20keys/cheat_sheet_registry.ahk) (`[Win+Alt+Shift+L]`).
