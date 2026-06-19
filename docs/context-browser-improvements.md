# Context browser — improvements

Scratchpad for the modal Context browser (Win+Alt+Shift+N). Code: Utils.ahk.

- [x] Remember last folder between opens (today always resets to context/ root)
- [x] JSON preview snippet in pane (topic from _meta or first lines) — most entries are .json, not images
- [x] Type-to-filter list (narrow rows as you type; letter jump only matches first character today)
- [x] Global file index search when filter has text (matches filename and relative path across all of context/; list shows filename only, folder path in footer)
- [x] Secondary action: paste path as text vs attach file (Enter always attaches — use Ctrl+Enter for path)
- [x] Copy path / open in Explorer without closing modal (Ctrl+C copy · Ctrl+H explorer)
- [x] Clickable breadcrumb segments in path subtitle
- [ ] Hide or collapse minimized/ subfolders by default (reverted — show minimized/ folders in list)
- [x] Cross-link image-references/ rows to matching research JSON when names align
- [x] Free previous preview HBITMAP on each row change (minor memory leak risk in long sessions)
- [x] Support relative paths in reference files (first-line pointers must be absolute today)
- [x] Shift+Enter paste file contents as text (follows reference pointers to target file, e.g. characters.md → characters.json)
