# Cheat Sheet Standard

This document defines the standard format for creating cheat sheet information in AutoHotkey scripts.

## Format Structure

### Header Format

```
AppName (Modifier)
```

- **AppName**: The name of the application or context
- **Modifier**: The keyboard modifier used (e.g., `Shift`, `Ctrl+Alt`, `Win+Alt+Shift`)
- Format: `AppName (Modifier)` with a space before the opening parenthesis

### Line Format

```
[KEY]ACTION_WITH_[KEY]HIGHLIGHTED
```

Each line follows this structure:

1. **Emoji** (optional but recommended) - Visual indicator for the action
2. **Bracketed Key** - The mnemonic key in brackets `[KEY]` (no spaces inside brackets in source)
3. **Action Description** - The action text with the mnemonic letter highlighted using `[KEY]` format

### Key Rules

1. **Mnemonic Keys**: Use mnemonic letters that match the action (e.g., `[D]` for Drawer, `[S]` for Search)
2. **Double Highlighting**: The mnemonic letter appears twice:
   - At the start of the line: `[KEY]`
   - Within the action text: `[KEY]word` (e.g., `[D]rawer`, `[S]earch`)
3. **No Spaces in Brackets**: In the source code, brackets should have no spaces: `[D]` not `[ D ]`
4. **Spacing**: No space between the initial bracket and the action text: `[D]Toggle` not `[D] Toggle`
5. **No Space Before Mnemonic in Text**: The mnemonic bracket in the action should be directly attached: `the[D]rawer` not `the [D]rawer`

## Example: Gemini (Shift)

```
Gemini (Shift)
📂 [D]Toggle the[D]rawer
💬 [N][N]ew chat
🔍 [S][S]earch
🔄 [M]Change[M]odel
🛠️ [T][T]ools
⌨️ [P]Focus[P]rompt field
📋 [C][C]opy last message
🔊 [R][R]ead aloud last message
🤖 [G]Send[G]emini prompt text
⛶ [F][F]ullscreen input
```

## Emoji Guidelines

Choose emojis that:

- **Clearly represent the action** (e.g., 🔍 for Search, 📋 for Copy)
- **Are visually distinct** from each other
- **Follow common conventions** (e.g., 💬 for chat, 🔄 for change/refresh)
- **Appear before the first bracket** on each line

### Common Emoji Mappings

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

## Implementation Notes

### In AutoHotkey Code

The cheat sheet string should be formatted as:

```autohotkey
appShortcuts := "Gemini (Shift)`r`n📂 [D]Toggle the[D]rawer`r`n💬 [N][N]ew chat`r`n..."
```

### Processing

The `ProcessCheatSheetText()` function will:

- Pad the first bracket (shortcut key) for alignment: `[D]` → `[ D ]`
- Add `>>> ` prefix for custom shortcuts
- Add `--- ` prefix for built-in shortcuts
- Leave mnemonic brackets in text unchanged (no padding)

### Display Result

After processing, the display will show:

```
>>> 📂 [ D ] Toggle the[D]rawer
>>> 💬 [ N ] [N]ew chat
>>> 🔍 [ S ] [S]earch
```

Note: The first bracket gets padded with spaces for alignment, but mnemonic brackets in the text remain unchanged.

## Complete Gemini Example

### Source Code Format

```autohotkey
appShortcuts := "Gemini (Shift)`r`n📂 [D]Toggle the[D]rawer`r`n💬 [N][N]ew chat`r`n🔍 [S][S]earch`r`n🔄 [M]Change[M]odel`r`n🛠️ [T][T]ools`r`n⌨️ [P]Focus[P]rompt field`r`n📋 [C][C]opy last message`r`n🔊 [R][R]ead aloud last message`r`n🤖 [G]Send[G]emini prompt text`r`n⛶ [F][F]ullscreen input"
```

### Display Format (after processing)

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

## Best Practices

1. **Consistent Mnemonics**: Use the same mnemonic letter at the start and in the action text
2. **Clear Actions**: Use action verbs that clearly describe what the shortcut does
3. **Logical Grouping**: Group related shortcuts together
4. **Emoji Selection**: Choose emojis that are universally recognizable
5. **Test Display**: Verify the processed output looks correct in the cheat sheet GUI

## Action Descriptions

- Use **action verbs** at the start when possible (e.g., "Toggle", "Change", "Focus", "Send")
- Keep descriptions **concise** but **clear**
- Use **camelCase** or **Title Case** for multi-word actions
- Include **context** when needed (e.g., "last message", "prompt field")
