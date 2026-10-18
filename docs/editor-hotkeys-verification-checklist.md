# Antigravity IDE hotkeys verification checklist

Source: `cheatSheets["Antigravity IDE.exe"]` in `Shift keys/cheat_sheet_registry.ahk`.

Check each box when the chord works as intended in **Antigravity IDE**. Leave unchecked if broken, missing, or not tested yet.

Scope tags:

- **shared-ahk** — `#HotIf IsEditorActive()` (should fire in Antigravity)
- **code-family** — `#HotIf IsCodeFamilyActive()` (VS Code–family path in Antigravity)
- **native / extension / keybinding** — not AHK (or only partly); depends on Antigravity install / extensions
- **N/A** — Cursor-only by design (do not expect it)

---

## Ctrl (no other modifiers)

- [ ] Ctrl+1 — Remove clustering / focus on code (shared-ahk)
- [ ] Ctrl+2 — Copy path via Explorer focus (shared-ahk)
- [ ] Ctrl+3 — CSV: Edit CSV (extension)
- [ ] Ctrl+4 — CSV: Apply changes + save (extension)
- [ ] Ctrl+5 — Markdown Enhanced: Export PDF (extension)
- [ ] Ctrl+6 — Marp export PDF (extension)
- [ ] Ctrl+7 — Build LaTeX project (extension)
- [ ] Ctrl+8 — View LaTeX PDF (extension)
- [ ] Ctrl+9 — Markdown Preview Enhanced: Insert Page Break (extension)
- [ ] Ctrl+M — Generate commit message + commit/push flow (code-family)
- [ ] Ctrl+G — Kill terminal (native / keybinding)
- [ ] Ctrl+Y — Fold all (native)
- [ ] Ctrl+U — Unfold all (native)
- [ ] Ctrl+O — Open Paste As… (native)
- [ ] Ctrl+H — Smart nav Editor↔Explorer / Reveal (shared-ahk)
- [ ] Ctrl+J — Select to Bracket (native)
- [ ] Ctrl+, — Fold all directories (native / keybinding)
- [ ] Ctrl+Q — Unfold all directories (native / keybinding)
- [ ] Ctrl+. — Agent modes / chat
- [ ] Ctrl+E — Default editor behavior
- [ ] Ctrl+R — File open Recent (native)
- [ ] Ctrl+T — Go to Type symbol (native)
- [ ] Ctrl+N — Chat session workflow
- [ ] Ctrl+Enter — Insert line below (native)
- [ ] Ctrl+P — Quick open / project search (native)
- [ ] Ctrl+; — Insert comment (native)
- [ ] Ctrl+D — Duplicate selection to next find match (native)- [ ] Ctrl+F — Find (native)
- [ ] Ctrl+Z — Undo (native)
- [ ] Ctrl+B — Toggle primary sidebar (native)
- [ ] Ctrl+Alt+. — Generate commit message (code-family)

---

## Shift

- [ ] Shift+F — Fold (shared-ahk)
- [ ] Shift+U — Unfold (shared-ahk)
- [ ] Shift+M — Markdown preview
- [ ] Shift+W — Move editor window
- [ ] Shift+T — Terminal (shared-ahk)
- [ ] Shift+N — New Terminal (shared-ahk)
- [ ] Shift+E — Explorer (shared-ahk)
- [ ] Shift+K — Markdown + window (shared-ahk)
- [ ] Shift+C — Command palette (shared-ahk)
- [ ] Shift+X — Expand selection (shared-ahk)
- [ ] Shift+S — Symbol in access view
- [ ] Shift+H — History (chat) (shared-ahk)
- [ ] Shift+I — Image paste
- [ ] Shift+G — Git repos fold SCM (shared-ahk)
- [ ] Shift+Q — Query Search (shared-ahk)
- [ ] Shift+R — Breadcrumbs menu (shared-ahk)
- [ ] Shift+O — Emoji selector (shared-ahk)
- [ ] Shift+D — Git section / Diff (shared-ahk)
- [ ] Shift+Z — Close all editors (shared-ahk)
- [ ] Shift+A — AI models switch (code-family branch)
- [ ] Shift+Y — Zen mode
- [ ] Shift+P — Pull or Sync Git (shared-ahk)
- [ ] Shift+V — Commit (Git)
- [ ] Shift+B — Push (Git)
- [ ] Shift+J — Quick commit timestamp + push (shared-ahk)

---

## Alt

- [ ] Alt+X — Shrink selection (shared-ahk)
- [ ] Alt+, — Classical Markdown Preview (extension)
- [ ] Alt+Y — Paste image to Markdown (extension)
- [ ] Alt+U — Scroll AI feed to bottom (code-family path)
- [ ] Alt+M — Quick shortcut menu (code-family)
- [ ] Alt+C — Add Context picker (code-family)
- [ ] Alt+A — Add file to AI Context (code-family)
- [ ] Alt+Q — Unpin current tab (native)
- [ ] Alt+P — Pin current tab (native)
- [ ] Alt+I — Reveal in Explorer + copy file (shared-ahk)
- [ ] Alt+H — Reveal in Explorer + open file (shared-ahk)
- [ ] Alt+6 — Share in OneDrive — Reveal + share link (shared-ahk)
- [ ] Alt+R — Refresh preview (native)
- [ ] Alt+E — File: Open File (shared-ahk)
- [ ] Alt+F — File: New File (native)
- [ ] Alt+O — File: New Folder (native)
- [ ] Alt+W — Focus other side / compare editor (native)
- [ ] Alt+S — Stash and Pull Git (shared-ahk)
- [ ] Alt+N — Review next file (shared-ahk)

---

## Ctrl+Shift / Ctrl+Alt / Alt+Shift

- [ ] Ctrl+Shift+L — Select all identical words (native)
- [ ] Ctrl+Shift+D — Debugging (native)
- [ ] Ctrl+Alt+L — Markdown Preview Enhanced: Toggle Live Update (extension)
- [ ] Ctrl+Alt+T — Markdown Preview Enhanced: Toggle Scroll Sync (extension)
- [ ] Ctrl+Alt+Up — Go to Parent Fold (native)
- [ ] Ctrl+Alt+Left — Sibling fold previous (native)
- [ ] Ctrl+Alt+Right — Sibling fold next (native)
- [ ] Ctrl+Alt+↑ — Add cursor above (native)
- [ ] Ctrl+Alt+↓ — Add cursor below (native)
- [ ] Shift+Alt+↑ — Copy line up (native)
- [ ] Shift+Alt+↓ — Copy line down (native)

---

## Alt (other) + function keys

- [ ] Alt+F12 — Peek Definition (native)
- [ ] Alt+↑ — Move line up (native)
- [ ] Alt+↓ — Move line down (native)
- [ ] Alt+Click — Multi-cursor (native)
- [ ] Alt+Z — Toggle word wrap (native)
- [ ] Alt+J — Jump to next review (native / keybinding)
- [ ] Alt+K — Previous review (native / keybinding)
- [ ] F2 — Rename symbol (native)
- [ ] F8 — Navigate problems (native)
- [ ] Shift+Delete — Delete line (native)

---

## Expected N/A (Cursor-only — do not expect in Antigravity)

- [x] Cursor Chat Alt+A UIA (Antigravity uses Code-family Alt+A instead)
- [x] Cursor permission Run / Allowlist / Proceed macros
- [x] Cursor Composer ToggleAiPane automation

---

## Priority smoke set (start here)

- [ ] Shift+E — Explorer
- [ ] Shift+C — Command palette
- [ ] Ctrl+H — Smart nav
- [ ] Alt+E — Open File
- [ ] Shift+V / Shift+B — Commit / Push
- [ ] Ctrl+M — Generate commit (+ push flow)
- [ ] Alt+A — Add file to AI context
- [ ] Alt+U — Scroll AI feed
- [ ] Shift+A — AI model picker
- [ ] Alt+C — Add Context
- [ ] Alt+M — Shortcut menu

---

## Broken / notes

| Chord | Symptom | Fix later? |
| ----- | ------- | ---------- |
|       |         |            |
|       |         |            |
|       |         |            |
