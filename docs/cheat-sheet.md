# Cheat sheet (ShiftKeys)

This document is the single reference for **authoring** cheat sheet strings, **using** the in-script overlays in [`Shift keys.ahk`](../Shift%20keys.ahk), and **resolving** which sheet applies. The implementation follows repository guidance in [efficiency-canon.md](efficiency-canon.md) (behavior parity, no extra hot-path IPC for the overlay).

**Canonical registry (all sheet strings):** [`Shift keys/cheat_sheet_registry.ahk`](../Shift%20keys/cheat_sheet_registry.ahk) — `cheatSheets` map entries and `GLOBAL_CHEAT_SHEET_RAW`.

**Canonical script:** [`Shift keys.ahk`](../Shift%20keys.ahk) — entry point; not `shiftkeys.autohotkey` (that filename is not used in this repo).

---

## How to open the overlays

| Gesture                            | Result                                                                                                           |
| ---------------------------------- | ---------------------------------------------------------------------------------------------------------------- |
| **Win+Alt+Shift+A** (quick tap)    | App-specific sheet for the foreground window (if registered). **ListView** (Section \| Shortcut \| Description). |
| **Win+Alt+Shift+A** (hold ~700ms+) | Global sheet: same **ListView** layout (system-wide chords + ZMK).                                               |
| **Win+Alt+Shift+/**                | Search across all registered cheat sheets and the global block (ListView; double-click a row to copy the line).  |

Closing:

- **Escape** while the overlay or search window has focus closes it (via `Gui.OnEvent("Escape", ...)` in [`Shift keys.ahk`](../Shift%20keys.ahk)).
- **Win+Alt+Shift+A** again toggles off the app or global sheet, and also closes the **Search cheat sheets** window if it is open.
- Otherwise close the search window normally (e.g. title bar).

On open, the **search/filter field is focused** so you can type immediately.

Both **app** (quick tap) and **global** (long hold) overlays use the same **ListView** shell in [`cheat_sheet_gui.ahk`](../Shift%20keys/cheat_sheet_gui.ahk):

- Columns: **Section**, **Shortcut**, **Description** (processed text parsed by `CheatSheet_ParseSheetRows()`).
- **Light theme:** white background, black text (`BackgroundFFFFFF` / `c000000` on ListView and filter Edit).
- Column widths: Section ~20%, Shortcut ~28% of overlay width (mins 80/100px; scales on portrait), Description fills remainder.
- Double-click a row to copy `Shortcut > Description`. Filter uses the same haystack/AND rules as before (`CheatSheet_LineMatchesQuery` on each row’s raw processed line).

Overlays are **centered** on the foreground window’s monitor (`GetActiveMonitorWorkArea_StandardBar` in [`Utils.ahk`](../Utils.ahk)) at **80% of that monitor’s work-area width and height** (`CHEAT_SHEET_WIDTH_FRAC` / `CHEAT_SHEET_HEIGHT_FRAC` in [`cheat_sheet_gui.ahk`](../Shift%20keys/cheat_sheet_gui.ahk)), with a **12px margin** on each edge so the window never spills into adjacent monitors (portrait-safe).

**WindowManagement** global chords (`Ctrl+Alt+Win`, close/cycle/minimize per monitor, MEH Alt+Tab, etc.) are documented in **`GLOBAL_CHEAT_SHEET_RAW`** under `=== WINDOW MANAGEMENT ===` in [`cheat_sheet_registry.ahk`](../Shift%20keys/cheat_sheet_registry.ahk) — that registry is the source of truth (not the comment block in `WindowManagement.ahk`).

---

## Search behavior (filter and cross-context search)

- The filter box is limited to **20 characters** per field.
- Matching is **case-insensitive** and runs on a **description haystack** aligned with what you read on screen: the line is stripped of leading `>>>` / `---`, then **each** `[...]` segment is replaced by its **inner text** (brackets removed, content kept). That yields the same words as the overlay’s plain text (e.g. `Toggle theDrawer` matches `drawer`).
- **Multiple words** (separated by spaces) use **AND** semantics: every term must appear in that haystack (e.g. `copy mail` requires both substrings in the description text).
- While the filter text is **non-empty**, each matching line is shown with a **`[Cluster label]`** prefix taken from the nearest preceding `=== Cluster label ===` header (same rules as modifier sections), so you can see which modifier group a hit belongs to without scrolling the full sheet.
- **`SearchCheatSheets(query, includeGlobal := true)`** uses the same rules and returns full **processed** lines for display. An empty query returns an empty map.

---

## Modifier clusters (standard layout)

App-specific sheets should group shortcuts **by modifier family** so scanning and maintenance stay consistent across apps. Use **section headers** on their own lines. Headers must **not** start with a mnemonic `[` bracket (so [`ProcessCheatSheetText()`](../Shift%20keys.ahk) leaves them unchanged—no `>>>`/`---` prefix on header lines).

**Official header syntax:** `=== Cluster label ===` (equals signs, space padding, short readable label).

**Canonical section order** (omit empty sections):

1. **Context line** — `AppName` or `AppName (Modifier)` (one line).
2. **Single-modifier groups** (alphabetically by modifier name): `Ctrl`, `Shift`, `Alt`, `Win`.
3. **Two-modifier groups** (fixed order): `Ctrl+Shift`, `Ctrl+Alt`, `Ctrl+Win`, `Alt+Shift`, `Alt+Win`, `Shift+Win`.
4. **Three or more modifiers** — e.g. `Ctrl+Alt+Shift` (rare; label clearly).
5. **Residual** — `Function keys & misc` (or similar) for bare `F2`/`F8`, `Shift+Delete`, or chords that do not fit a cluster without awkward duplication.

Within each section, keep the existing line format (emoji + `[KEY]` + description). When a chord fits multiple clusters (e.g. `Alt+F12`), place it under the **most specific** cluster you use in that sheet (e.g. `Alt (other chords)`), and document that rule here once.

**Display merge:** [`ProcessCheatSheetText()`](../Shift%20keys.ahk) expands the **first** `[KEY]` on each shortcut line using the active modifier cluster: the nearest preceding `=== Cluster label ===` header (or a leading **`AppName (Modifier)`** line such as `Explorer (Shift)`) supplies an implied prefix (`Shift+`, `Ctrl+`, etc.), so the overlay shows **`[Shift+M]`** instead of bare **`[M]`** when the section implies Shift. Brackets that already list a full chord (`+`, `/`, `...`, or parenthetical keys) are left unchanged. The **`>>>` / `---`** classification still uses the **original** first bracket (before merge) so built-in detection matches prior behavior.

**Reference implementation:** `cheatSheets["Cursor.exe"]` in [`Shift keys/cheat_sheet_registry.ahk`](../Shift%20keys/cheat_sheet_registry.ahk) is the **template** for modifier clustering (Cursor-first, complex). New app sheets should mirror this structure before adding one-off sections.

---

## Processing pipeline

1. **`GetCheatSheetText()`** picks raw text for the active context (process, window title, Chrome site, Teams mode, etc.).
2. **`NormalizeMojibake()`** fixes common UTF-8→ANSI display glitches (arrows, dashes).
3. **`ProcessCheatSheetText()`** merges the active section (or `AppName (Modifier)` context line) into the **first** `[shortcut]` when the key is implied (e.g. `[M]` under `=== Shift ===` becomes `[Shift+M]`), pads that bracket, and prefixes lines with `>>>` (custom/remapped) or `---` (built-in style chords; classification uses the **unmerged** bracket).
4. **`CheatSheet_ParseSheetRows()`** in [`cheat_sheet_gui.ahk`](../Shift%20keys/cheat_sheet_gui.ahk) splits processed lines into ListView rows; filter/search still operate on the underlying processed line text.

Filter and **Search all sheets** operate on this **processed** text (with haystack rules above) so the query matches readable words, not bracket notation.

---

## Registry: `cheatSheets` map keys

Each row is a key in `cheatSheets` in [`Shift keys/cheat_sheet_registry.ahk`](../Shift%20keys/cheat_sheet_registry.ahk). Line numbers drift with edits; search the file for `cheatSheets["Key"]`.

| Map key              | When it applies                                                                                                                                    |
| -------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------- |
| `Mercado Livre`      | `chrome.exe` and `IsMercadoLivreActive()` (URL/site, not title).                                                                                   |
| `Shopee`             | `chrome.exe`, resolver chain, and `IsShopeeActive()` when applicable.                                                                              |
| `WhatsApp`           | Chrome tab title contains `WhatsApp`.                                                                                                              |
| `OUTLOOK.EXE`        | Default Outlook when no reminder/message/appointment-specific sheet matches.                                                                       |
| `OutlookReminder`    | Outlook window title matches `Reminder`.                                                                                                           |
| `OutlookAppointment` | Outlook title matches `Appointment`, `Meeting`, or `Event`.                                                                                        |
| `OutlookMessage`     | Outlook title matches ` - Message (` (HTML inspector).                                                                                             |
| `TeamsMeeting`       | `IsTeamsMeetingActive()`.                                                                                                                          |
| `TeamsChat`          | `IsTeamsChatActive()`.                                                                                                                             |
| `Spotify.exe`        | Foreground process `Spotify.exe`.                                                                                                                  |
| `ONENOTE.EXE`        | Foreground process OneNote.                                                                                                                        |
| `chrome.exe`         | Chrome: general browser shortcuts; often combined with a site-specific sheet.                                                                      |
| `Chrome PDF Viewer`  | Chrome and `IsChromePdfViewerActive()`.                                                                                                            |
| `Cursor.exe`         | Foreground Cursor. **Reference layout** for [modifier clusters](#modifier-clusters-standard-layout).                                               |
| `Code.exe`           | Foreground VS Code. Separate sheet, initially derived from Cursor layout; diverges as migration proceeds to official VS Code and Copilot defaults. |
| `explorer.exe`       | File Explorer.                                                                                                                                     |
| `mspaint.exe`        | Paint.                                                                                                                                             |
| `ClipAngel.exe`      | Clip Angel.                                                                                                                                        |
| `Figma.exe`          | Figma.                                                                                                                                             |
| `Gmail`              | Chrome title contains `Gmail`.                                                                                                                     |
| `Google Keep`        | Chrome title contains `Google Keep` or `keep.google.com`.                                                                                          |
| `FileDialog`         | Class `#32770` and dialog text contains the namespace tree control (EN/PT).                                                                        |
| `Settings`           | Title `Settings` or `Configurações`.                                                                                                               |
| `Command Palette`    | Title contains `Command Palette`.                                                                                                                  |
| `EXCEL.EXE`          | Excel.                                                                                                                                             |
| `Power BI`           | `PBIDesktop.exe` or title contains `powerbi`.                                                                                                      |
| `UIATreeInspector`   | Chrome title or `AutoHotkey64.exe` + UIATreeInspector title.                                                                                       |
| `Settle Up`          | Chrome title contains `Settle Up`.                                                                                                                 |
| `Miro`               | Chrome title contains `Miro`.                                                                                                                      |
| `Wikipedia`          | Chrome title contains `Wikipedia` or `wikipedia.org`.                                                                                              |
| `YouTube`            | Chrome title contains `YouTube`.                                                                                                                   |
| `Google`             | Chrome, no other site sheet, and title is `Google` or ` - Google Search`.                                                                          |
| `ChatGPT`            | Chrome title contains `chatgpt`.                                                                                                                   |
| `Gemini`             | Chrome title contains `gemini` (`cheatSheets["Gemini"]`).                                                                                          |
| `Mobills`            | Chrome title contains `Mobills`.                                                                                                                   |

### Editor Alt+S stash and pull (Cursor / VS Code)

**Alt+S** in Cursor and VS Code is wrapped by AHK (`$!s` → `Editor_GitStashAndPull()` in [`Shift keys/cursor_predicates.ahk`](../Shift%20keys/cursor_predicates.ahk)) as **Git Stash and Pull**: re-sends **Alt+S** (native **Git: Stash**), **UIA-waits** for the stash-message QuickInput, sleeps **400ms**, sends **Enter** for an empty stash message, then sends **Shift+P** (native **Git: Pull**). When the status bar sync item (`status.scm.1`) no longer shows pending **Pull N commits**, plays **`quick-update-success.wav`** (respects global sound toggle). Listed on both overlays as `💾 [S][S]tash and Pull (Git) (ahk)`.

### Special resolution (not only `exe` match)

- **Chrome (`chrome.exe`):** Site-specific keys are chosen via **`PickChromeAppSheetKey()`** in [`Shift keys/cheat_sheet_gui.ahk`](../Shift%20keys/cheat_sheet_gui.ahk) (sequential overwrites, same semantics as the former `if` chain). The generic `chrome.exe` sheet may be appended for combined display.
- **Outlook:** Reminder / message / appointment inspectors override generic `OUTLOOK.EXE` where applicable.
- **Teams:** Meeting vs chat uses helper predicates, not raw `exe` only.

### Global shortcuts text

The long-hold overlay text lives in **`GLOBAL_CHEAT_SHEET_RAW`** in [`Shift keys/cheat_sheet_registry.ahk`](../Shift%20keys/cheat_sheet_registry.ahk). It lists **Win+Alt+Shift** and **Ctrl+Alt+Win** chords. Free **Alt+Shift** letters (starting with **W**) and free **Ctrl+Alt+Win** letters are listed in the same string under the available-slots sections. **`GetGlobalCheatSheetRawText()`** in [`Shift keys/cheat_sheet_gui.ahk`](../Shift%20keys/cheat_sheet_gui.ahk) returns that string. **`SearchCheatSheets()`** includes it when `includeGlobal` is true.

A **`=== ZMK KEYBOARD (eyelash_sofle.keymap) ===`** block at the end of that string documents the Sofle ZMK keymap by **layer and physical key** (e.g. `[ZMK L0 · L] hold > …`). Source of truth for bindings: `C:\Users\eduev\Documents\ZMK\zmk-sofle\config\eyelash_sofle.keymap`. Legend: **L0** = base, **L1** = hold left thumb, **L2** = hold right thumb, **L3** = sticky rapid arrows (activate from **L2·W**), **L4** = auto when L2+L3 are both active. Tap-dance lines use **1× / 2× / 3×** for single/double/triple tap. Update the registry manually when the keymap changes.

---

## Format structure (authoring)

### Header format

```
AppName (Modifier)
```

- **AppName**: The name of the application or context
- **Modifier**: The keyboard modifier used (e.g., `Shift`, `Ctrl+Alt`, `Win+Alt+Shift`)
- Format: `AppName (Modifier)` with a space before the opening parenthesis

### Line format

```
[KEY]ACTION_WITH_[KEY]HIGHLIGHTED
```

Each line follows this structure:

1. **Emoji** (optional but recommended) - Visual indicator for the action
2. **Bracketed Key** - The mnemonic key in brackets `[KEY]` (no spaces inside brackets in source)
3. **Action Description** - The action text with the mnemonic letter highlighted using `[KEY]` format

### Key rules

1. **Mnemonic Keys**: Use mnemonic letters that match the action (e.g., `[D]` for Drawer, `[S]` for Search)
2. **Double Highlighting**: The mnemonic letter appears twice:
   - At the start of the line: `[KEY]`
   - Within the action text: `[KEY]word` (e.g., `[D]rawer`, `[S]earch`)
3. **No Spaces in Brackets**: In the source code, brackets should have no spaces: `[D]` not `[ D ]`
4. **Spacing**: No space between the initial bracket and the action text: `[D]Toggle` not `[D] Toggle`
5. **No Space Before Mnemonic in Text**: The mnemonic bracket in the action should be directly attached: `the[D]rawer` not `the [D]rawer`

---

## Mnemonic key conventions

Guidelines (not strict rules) for consistency across applications.

### Primary conventions

- **`N`** - **New**: New chat, new document, new file, new tab, new item
- **`S`** - **Search**: Search, find, seek. In editor Git chords (e.g. Cursor/VS Code **Alt+S**), **`S`** = **Stash** (context override; app-specific meanings are OK).
- **`C`** - **Copy**: Copy, clipboard operations
- **`P`** - **Prompt/Input**: Prompt field, input focus, paste
- **`F`** - **Fullscreen/Focus**: Fullscreen mode, focus actions, find
- **`R`** - **Read/Reply**: Read aloud, reply, refresh, reload
- **`T`** - **Tools**: Tools menu, toggle, tab
- **`D`** - **Drawer/Delete**: Drawer, delete, duplicate
- **`M`** - **Model/Menu**: Model selection, menu, move
- **`G`** - **Go/Generate**: Go to, generate, Gemini (when context-specific)
- **`H`** - **Help/History**: Help, history, home
- **`E`** - **Edit/Export**: Edit, export, expand
- **`O`** - **Open/Options**: Open, options, organize
- **`U`** - **Undo/Update**: Undo, update, unfold
- **`I`** - **Insert/Import**: Insert, import, info
- **`K`** - **Keep/Keyboard**: Keep, keyboard shortcuts
- **`L`** - **List/Link**: List, link, location
- **`W`** - **Window/Write**: Window operations, write
- **`X`** - **Exit/eXport**: Exit, export, close
- **`Z`** - **Undo/Zoom**: Undo (common), zoom

### Secondary conventions

- **`Y`** - Toggle actions or yes/confirm when primary keys are taken
- **`V`** - View, verify, version
- **`B`** - Back, bookmark, bold
- **`J`** - Jump, join
- **`Q`** - Quit, query, quick

### Decision guidelines

1. Prioritize conventions when they fit
2. Consider context; app-specific meanings (e.g., `G` for Gemini) are OK
3. Avoid conflicts; pick the next best mnemonic if needed
4. Prefer clarity over strict convention
5. Be flexible

---

## Emoji guidelines

Choose emojis that:

- Clearly represent the action (e.g., 🔍 for Search, 📋 for Copy)
- Are visually distinct from each other
- Follow common conventions (e.g., 💬 for chat, 🔄 for change/refresh)
- Appear before the first bracket on each line

### Common emoji mappings

- 📂 / 📁 - Folders, drawers, navigation
- 💬 - Chat, messages, conversations
- 🔍 - Search, find
- 🔄 - Change, switch, update
- 🛠️ - Tools, settings
- ⌨️ - Input, prompt, keyboard
- 📋 - Copy, clipboard
- 🔊 - Audio, read aloud, sound
- 🤖 - AI, automation, Gemini
- ⛶ - Fullscreen, expand
- ✏️ - Edit, write
- 🗑️ - Delete, remove
- ✅ - Confirm, check
- ❌ - Cancel, close
- ⬆️ / ⬇️ - Up, down navigation
- ➡️ / ⬅️ - Next, previous

---

## Implementation notes (AutoHotkey)

### In code

The cheat sheet string should be formatted as:

```autohotkey
appShortcuts := "Gemini (Shift)`r`n📂 [D]Toggle the[D]rawer`r`n💬 [N][N]ew chat`r`n..."
```

### Processing

The `ProcessCheatSheetText()` function (with `PadShortcut()` in [`Shift keys/config.ahk`](../Shift%20keys/config.ahk)) will:

- Pad the **first** `[...]` on each line so the bracket group reaches a target width (default 18 characters) by centering spaces **inside** the brackets when the inner text is shorter than the target (longer keys are left unchanged).
- Add `>>> ` prefix for custom shortcuts
- Add `--- ` prefix for built-in shortcuts (detected via `IsBuiltInShortcut()`)
- Leave additional mnemonic brackets in the rest of the line unchanged (no padding)

### Display result (conceptual)

The overlay does **not** show square brackets. After processing, a line is rendered roughly like:

- `>>> 📂` then a padded key column (spaces + **bold** mnemonic letter(s)), then the action text with inline mnemonic letters also **bold** / larger (e.g. **D**rawer, **S**earch).

Authoring still uses `[KEY]` in source strings; padding still applies to the **first** bracket group only (`PadShortcut()`). Inline mnemonic brackets in the rest of the line are stripped visually; emphasis is formatting only.

---

## Example: Gemini (Shift)

### Source

```autohotkey
appShortcuts := "Gemini (Shift)`r`n📂 [D]Toggle the[D]rawer`r`n💬 [N][N]ew chat`r`n🔍 [S][S]earch`r`n🔄 [M]Change[M]odel`r`n🛠️ [T][T]ools`r`n⌨️ [P]Focus[P]rompt field`r`n📋 [C][C]opy last message`r`n🔊 [R][R]ead aloud last message`r`n🤖 [G]Send[G]emini prompt text`r`n⛶ [F][F]ullscreen input"
```

### After processing

```
Gemini (Shift)
>>> 📂 [ D ] Toggle the[D]rawer
>>> 💬 [ N ] [N]ew chat
>>> 🔍 [ S ] [S]earch
>>> 🔄 [ M ] Change[M]odel
>>> 🛠️ [ T ] [T]ools
>>> ⌨️ [ P ] Focus[P]rompt field
>>> 📋 [ C ] [C]opy last message
>>> 🔊 [ R ] [R]ead aloud last message
>>> 🤖 [ G ] Send[G]emini prompt text
>>> ⛶ [ F ] [F]ullscreen input
```

---

## Best practices

1. **Consistent Mnemonics**: Use the same mnemonic letter at the start and in the action text
2. **Clear Actions**: Use action verbs that clearly describe what the shortcut does
3. **Logical Grouping**: Group related shortcuts together
4. **Emoji Selection**: Choose emojis that are universally recognizable
5. **Test Display**: Verify the processed output looks correct in the cheat sheet GUI

## Action descriptions

- Use **action verbs** at the start when possible (e.g., "Toggle", "Change", "Focus", "Send")
- Keep descriptions **concise** but **clear**
- Use **camelCase** or **Title Case** for multi-word actions
- Include **context** when needed (e.g., "last message", "prompt field")

---

## Add a new cheat sheet (checklist)

1. Add the raw string to `cheatSheets["YourKey"]` in [`Shift keys/cheat_sheet_registry.ahk`](../Shift%20keys/cheat_sheet_registry.ahk) using the header and line rules above.
2. For non-trivial apps, follow [Modifier clusters (standard layout)](#modifier-clusters-standard-layout); compare with `Cursor.exe` as needed.
3. If the sheet is not selected by process name alone, extend `GetCheatSheetText()` and (for Chrome) `PickChromeAppSheetKey()` in [`Shift keys/cheat_sheet_gui.ahk`](../Shift%20keys/cheat_sheet_gui.ahk).
4. Press **Win+Alt+Shift+A** in the target app and confirm the overlay.
5. If the resolver is non-obvious, add a short note to this document’s registry table.

---

## Programmatic search

- **`SearchCheatSheets(query, includeGlobal := true)`** returns a `Map` of context label → array of matching **processed** lines (empty query returns an empty map).
