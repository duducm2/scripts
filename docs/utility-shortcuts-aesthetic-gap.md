# Utility Shortcuts aesthetic gap report

**Purpose:** Inventory key-driven selection modals relative to the Utility Shortcuts / project-selector ListView aesthetic.

**Context:** Handy AI (`#!+C`) and the former Catppuccin captionless key menus listed below were migrated to ListView chrome. Related doctrine: [efficiency-canon.md](efficiency-canon.md), [standard_information_display.md](standard_information_display.md). Companion models: [ai-companion-models.md](ai-companion-models.md).

---

## 1. Canonical aligned look

| Trait    | Detail                                                                                 |
| -------- | -------------------------------------------------------------------------------------- |
| Window   | `Gui("+AlwaysOnTop +ToolWindow", titled)` — captioned tool window                      |
| Font     | Segoe UI `s10` (default chrome; not bold Catppuccin title stack)                       |
| Layout   | Hint line + Char-first `ListView`                                                      |
| Keys     | Digit/letter select; Enter / double-click activate; Esc cancel                         |
| Not used | `-Caption`, `BackColor := "1E1E2E"`, per-row `cCDD6F4` / green highlight Text controls |

---

## 2. Already aligned

| UI                             | Entry                             | File                                                                                                                |
| ------------------------------ | --------------------------------- | ------------------------------------------------------------------------------------------------------------------- |
| Utility Shortcuts              | `#!+U` / `#!+W`                   | [`Utils/hotstring_selector_gui.ahk`](../Utils/hotstring_selector_gui.ahk)                                           |
| Project Selector               | `^!#0`                            | [`WindowManagement/cursor_window_select.ahk`](../WindowManagement/cursor_window_select.ahk) (`ShowProjectSelector`) |
| Import Management              | `#!+X` / Utility `[J]` / `#!+F`×2 | [`Utils/import_mgmt_launcher.ahk`](../Utils/import_mgmt_launcher.ahk)                                               |
| Handy AI models                | `#!+C`                            | [`Utils/handy_ai_model_gui.ahk`](../Utils/handy_ai_model_gui.ahk)                                                   |
| AI companion model list        | Shift+L                           | [`Utils/ai_companion_model_selector.ahk`](../Utils/ai_companion_model_selector.ahk)                                 |
| Study material / topic         | Study-topic / Peek–QuickLook path | [`Utils/peek_pdf_study_02.ahk`](../Utils/peek_pdf_study_02.ahk)                                                     |
| Transfer to Cursor window pick | D2C / Gemini→Cursor transfer      | [`Utils/gemini_cursor_transfer.ahk`](../Utils/gemini_cursor_transfer.ahk)                                           |
| Cursor shortcut menu           | Alt+M in Cursor                   | [`Shift keys/hotif_editor_01.ahk`](../Shift%20keys/hotif_editor_01.ahk)                                             |
| VS Code shortcut menu          | Parallel quick-shortcuts menu     | [`Shift keys/hotif_scroll_ai.ahk`](../Shift%20keys/hotif_scroll_ai.ahk)                                             |
| Mercado Livre sort menu        | ML sort picker (`+o`)             | [`Shift keys/hotif_mercado_livre.ahk`](../Shift%20keys/hotif_mercado_livre.ahk)                                     |

Finance, Memory Palace browse, click-sequence, and similar apps already use titled `+ToolWindow` + ListView chrome.

**Migration status:** The former §3 Catppuccin captionless key menus are migrated (behavior parity; chrome only).

---

## 3. Gap: Catppuccin captionless key menus

None remaining in the original gap set. (This section kept empty for historical clarity.)

---

## 4. Related but different

Captionless dark UIs that are **grids / previews**, not simple Char ListView menus. Do not treat as drop-in Handy-style migrations:

| UI                                    | File                                                                            |
| ------------------------------------- | ------------------------------------------------------------------------------- |
| Dictation visible-window paste picker | [`Utils/dictation_visible_paste.ahk`](../Utils/dictation_visible_paste.ahk)     |
| WM minimized / hidden window list     | [`WindowManagement/minimized_list.ahk`](../WindowManagement/minimized_list.ahk) |

---

## 5. Explicitly out of scope

Not selector aesthetic debt. Intentional surfaces per [standard_information_display.md](standard_information_display.md):

- `StandardLoadingBar` / `ShowWithKeys` banners (Interactive Input / Loading / Information Only)
- Language flag and other persistent indicators
- Focus-mode monitor blackout
- Mousemaster / square-jump / mouse-jump spatial overlays
- Transient dictation / status / confirmation banners (e.g. merge/cleanup overlays)
