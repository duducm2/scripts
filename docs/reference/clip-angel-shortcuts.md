# Clip Angel shortcuts

Clip Angel keeps a history of everything copied to the Windows clipboard and lets you paste any earlier item back. History is stored in a local SQLite database (path shown under Settings, Information). It records plain text, RTF, HTML, images, and files, including the source window and process.

On this PC (ClipAngel 2.13), open the history with the tray icon or **Ctrl+Alt+C**. That hotkey opens the window without changing the filter or selection. Press it again to close the window. The factory open hotkey, **Alt+V**, is unset here.

Window shortcuts below are fixed menu shortcuts. Global hotkeys are the ones in Settings on this PC.

Sources: this PC’s Settings window, the [Clip Angel description](https://sourceforge.net/p/clip-angel/wiki/Description/), and the [Clip Angel source](https://github.com/tormozit/ClipAngel) (`Properties/Settings.settings` factory defaults, `Main.resx` menu shortcuts, `Main.cs` key handling).

## Global hotkeys

These work from any program. **No** means unset. **This PC** is the binding in Settings. **Factory** is the value shipped with Clip Angel.

| This PC | Factory | Action |
| --- | --- | --- |
| Ctrl+Alt+C | No | Open the window without changing the filter or selection. Press again to close. |
| No | Alt+V | Open the window, clear filters, and select the latest clip. Press again to close. |
| No | Alt+B | Open the window filtered to favorites only. Press again to close. |
| Ctrl+Shift+V | No | Paste the latest clip as plain text. |
| Ctrl+Alt+V | No | Paste the current clip and select the next (newer) clip. The tray tooltip shows the newly selected clip. |
| Ctrl+Alt+B | No | Paste the current clip and select the previous (older) clip. The tray tooltip shows the newly selected clip. |
| No | No | Compare the last two clips in the configured diff application. |
| No | No | Simulate fast keyboard input of the latest clip. |
| No | No | Turn clipboard monitoring on or off. |
| No | No | Force-capture the current clipboard while monitoring is off. |
| Off | Off | **Capture selected text by Control+F3** is unchecked. The key cannot be changed. When the option is on, Ctrl+F3 captures the selection and is still delivered to the active app. |

Do not assign **Ctrl+V** as a Clip Angel hotkey. Paste sends Ctrl+V, so that binding reopens the window instead of pasting.

Menu names and shortcuts below match the Window, List, and Clip menus on this PC. **none** means the menu item has no shortcut.

## Window

| Shortcut | Menu item |
| --- | --- |
| F10 | Switch focus list/text |
| Tab | Set focus to clip text |
| Ctrl+T | Always on top |
| none | Show in taskbar |
| none | Monitoring clipboard |
| none | Clear clipboard |
| none | Connect clip channel |
| none | Exit |
| Escape | Close the window. This key is not on the menu. |

## List

| Shortcut | Menu item |
| --- | --- |
| none | Secondary columns |
| none | Move to top after use |
| Ctrl+Alt+Up | Move to top |
| Ctrl+Up | Move up |
| Ctrl+Down | Move down |
| none | Filter by date |
| none | Sort |
| Ctrl+Alt+K | Delete all non favorite |
| Ctrl+1 | Show all marks |
| Ctrl+2 | Show only favorite |
| Ctrl+3 | Show only used |
| Alt+1 | Show all types |
| Alt+2 | Show only images |
| Alt+3 | Show only files |
| Alt+4 | Show only texts |
| F7 | Clear filters |
| F8 | Clear filters and select prev. |
| F9 | Clear filters and select top |
| Alt+Down | Find next |
| Alt+Up | Find previous |
| Ctrl+Home | Select top |

## Clip

| Shortcut | Menu item |
| --- | --- |
| Ctrl+D | Search text in list |
| none | Send clip to channel |
| Del | Delete |
| F2 | Edit title |
| F3 | Find next |
| Shift+F3 | Find previous |
| Ctrl+N | Native formatting |
| Ctrl+W | Word wrap |
| none | Monospaced font |
| Alt+Q | Mark as favorite |
| Alt+W | Unmark as favorite |
| F4 | Edit text |
| Ctrl+F4 | Open file |
| Ctrl+S | Save as file... |
| Ctrl+Alt+J | Join text of clips |
| none | Duplicate text clip |
| none | Text compare |
| none | Translate text |
| none | Upload image to web |
| none | Text from picture |
| none | Ignore application |
| none | Export clips |
| none | Import clips |
| none | Decode URL |

### Paste

Clip → Paste. Paste goes into the window that was active before Clip Angel opened.

| Shortcut | Menu item |
| --- | --- |
| Enter | Paste |
| Ctrl+Enter | Paste text |
| Ctrl+Shift+Enter | Paste line |
| none | Paste file |
| Ctrl+P | Paste special... |
| none | Simulate fast char input |
| none | Simulate slow char input |
| Double-click | Same as Paste. This is not on the menu. |

## Keys that are not on the menus

| Shortcut | Action |
| --- | --- |
| Type | Filter the list. Character keys typed in the list go to the search field. |
| Up, Down, Page Up, Page Down | Move in the list. These also work while the search field has focus. |
| Ctrl+End | Move to the end of the list, including from the search field. |
| Ctrl+A | Select clips in the list, up to 300. In the HTML view, select all text. |
| Ctrl+Insert | Copy the current clip to the clipboard. |
| Ctrl+C | Copy the current clip from the list. In the clip text, copy the selection. |
| Alt+click | Open a marked hyperlink in the clip text. |
