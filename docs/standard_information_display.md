# Standard Information Display

A shared UI component for **loading indication**, **information-only messages**, and **interactive input** across all AHK scripts. Defined in `Utils.ahk` and used by Act, AppLaunchers, Gemini, Shift keys, WindowManagement, Microsoft Teams, and any script that includes Utils. This document is the **single source of truth** for banner behavior and the canonical way to show loading, information, and interactive banners across all scripts.

## Overview and Purpose

- **Single shared component** for loading, information, and user-input feedback across all AHK scripts
- **Replaces ad-hoc banners and overlays** with a consistent user experience
- **Monitor-aware positioning** – centers on the active window's monitor or the primary monitor; optional **active-monitor tracking** (`trackActiveMonitor` on interactive banners) recenters the bar when the foreground window moves to another display while the banner is open (dictation and Gemini transfer flows).
- **Four display categories** – Loading Indication (progress bar), Information Only (static message), Interactive Input (key press + optional timeout), Persistent Indicator (long-lived state pinned to a monitor corner)
- **Standard font size** – 17px for all banners and loading indicators (except `ShowSingleCharTabBanner_Utils`, which keeps 72px)
- **Emoji** – every banner message must start with an emoji (e.g. ⏳ loading, ✅ success, ❌ error, ❓ user input)

## Display Categories

Three categories define how information is shown to users. Use the same API and semantics for all consumers (human and tooling).

| Category                  | Label / prefix   | Implementation                                                                                                                          | Purpose                                                                           |
| ------------------------- | ---------------- | --------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------- |
| **Loading Indication**    | ⏳, 🔄           | `StandardLoadingBar_Show` with `passive: false` (default), optional `StandardLoadingBar_Update`; progress bar animates                  | Progress tracking; show → update → hide lifecycle                                 |
| **Information Only**      | ✅, ℹ, 📋, ❌, ⚠ | `StandardLoadingBar_Show` with `passive: true` and immediate or delayed `Hide`; or `ShowCenteredOverlay_Utils(text, duration, bgColor)` | Static message; no user input; auto-hide by duration                              |
| **Interactive Input**     | ❓, ⌨            | `StandardLoadingBar_ShowWithKeys(state, keyCallbacks, timeoutMs, …)` with optional `promptKeys` strip                                   | Wait for specific key presses; optional timeout; fixed bottom strip for key hints |
| **Persistent Indicator**  | 🇬🇧, 🇧🇷 (image)  | `LanguageFlag_Show(slot)` / `LanguageFlag_Hide()` (Utils.ahk); opaque chip, bottom-right of active monitor; independent monitor tracker | Long-lived state indicator (e.g. transcription language) that survives across transient banners and follows the foreground window |

- **Loading Indication:** Call `Show` at start, `Update` at milestones, `Hide` in all exit paths. The animated progress bar indicates ongoing work.
- **Information Only:** Short-lived message; use `ShowCenteredOverlay_Utils(text, duration, bgColor)` for the common case (show + auto-hide after `duration` ms), or `StandardLoadingBar_Show` with `passive: true` plus `Hide(duration)`.
- **Interactive Input:** Use `StandardLoadingBar_ShowWithKeys`; pass `promptKeys` (e.g. `"[Y] Confirm  [N] Cancel"`) for a fixed bottom strip so the main message and key hints stay clearly separated. Interactive confirmations can use `noBorder: true` and a neutral/dark `barColor` (e.g. `"1E1E2E"`) for a single clean banner (see `DictationGeminiConfirm_ShowAndWait` in Utils.ahk).
- **Chord-triggered menus** (`#!+…`, `^!+#…`): `ShowWithKeys` calls `StandardLoadingBar_WaitForTriggerKeyRelease()` so selection keys (`1`–`9`, letters) work after Win/Ctrl/Alt/Shift and the trigger key are released.
- **`[Esc] Cancel` on prompts:** Escape dismiss (`$*Escape`, poll, Gui `Escape`) is registered by default. Pass `"Escape"` in `keyCallbacks` only when cancel needs extra logic; set `skipEscapeDismiss := true` when Escape must not capture input (e.g. Command Palette bookmark prompt in Shift keys).
- **Selection menus:** Use `preserveUserFocus := false` so the overlay activates and owns key input; use `preserveUserFocus := true` only when the user must keep typing in the same field (dictation paste-target flows).
- **Digit/letter keys (`1`, `2`, `Y`, …):** `ShowWithKeys` registers hook hotkeys as `$*key` at `#InputLevel 10` and runs **`StandardLoadingBar_KeysSelectionPoll`** (edge-triggered `GetKeyState`, including `Numpad1`–`Numpad9`) so choices work when another script or the foreground app would otherwise consume the key. Poll ignores edges while Win/Ctrl/Alt are held (same rule as Outlook Reminders).
- **Persistent Indicator:** Long-lived, **opaque** overlay anchored to the bottom-right of the active monitor's work area. Does **not** auto-hide; remains visible until explicitly hidden or replaced. Uses an independent track timer so it does not collide with the centered `StandardLoadingBar_*` overlay (a transient banner showing/hiding never disturbs a persistent indicator). See dedicated subsection below.
- **Outlook activation-failed prompts:** For mailbox/calendar failures such as `"Outlook mailbox and calendar are not open (activation failed)"`, do **not** use `MsgBox` buttons. Use a standard **Interactive Input** banner that asks whether to activate Outlook (for example: `❓ Would you like to activate Outlook?`) with `ShowWithKeys` + prompt strip (e.g. `"[Y] Activate Outlook  [N] Cancel"`).

### Banner Types (by category)

| Category                  | Emoji / prefix   | Use case                                      |
| ------------------------- | ---------------- | --------------------------------------------- |
| **Loading Indication**    | ⏳, 🔄           | Progress bar + emoji; `passive: false`        |
| **Information Only**      | ✅, ℹ, 📋, ❌, ⚠ | Passive text-only; emoji + message            |
| **Interactive Input**     | ❓, ⌨            | `ShowWithKeys` + fixed bottom strip with keys |
| **Persistent Indicator**  | 🇬🇧, 🇧🇷 (image)  | Pinned image or text fallback in monitor corner; opaque; no auto-hide |

### Persistent Indicators

Persistent indicators are long-lived overlays that show **state** rather than transient progress or messages. The first instance is the language flag for the Handy transcription tool (slot 3 = Cohere English → 🇬🇧, slot 4 = Cohere Portuguese → 🇧🇷; slots 1/2 hide the flag).

- **Anchor:** bottom-right corner of the active monitor's work area, with a 20 px margin from the right and bottom edges.
- **Borderless rendering:** the GUI is a small `+ToolWindow +AlwaysOnTop -Caption -DPIScale` window with `MarginX = MarginY = 0` and no `Border` option on the `Picture` control, so the window sizes exactly to the bitmap (no outer frame, no inner frame, no padding). It must **not** use `WS_EX_TRANSPARENT` (formerly combined with layered as `+E0x80020`): that extended style tells Windows not to paint the window normally, which made the indicator invisible. Focus is not stolen because the window is tool-style and shown with `NA` (no activate); the chip is small and sits in a corner.
- **Opacity:** the language flag uses default opaque rendering (no `WinSetTransparent` on the indicator). Use `WinSetTransparent` only if a future design needs window-level alpha.
- **Multi-monitor tracking:** reuses `GetMonitorIndexForForeground_StandardBar` and `GetActiveMonitorWorkArea_StandardBar` (the same helpers `StandardLoadingBar_*` uses) so the flag follows the foreground window across all monitors. The poll interval matches the standard bar (`LANGUAGE_FLAG_TRACK_INTERVAL`, 115 ms).
- **Independent timer:** the indicator must own a **separate** track-timer global (e.g. `g_LanguageFlagTrackTimer`). Sharing `g_StandardLoadingBarTrackTimer` would let any transient `Show`/`Hide` cycle silently kill the persistent tracker.
- **Reload-safe:** the active slot is persisted in `data\handy_ai_model.ini`. On script load, `LanguageFlag_InitFromPersistedSlot()` (deferred via `SetTimer(..., -250)`) reads `Handy_GetPersistedAiModelSlot()` and restores the correct flag.
- **Image assets:** `images/flags/united-kingdom.png` and `images/flags/brazil.png` (shown at `LANGUAGE_FLAG_WIDTH` px wide in the `Picture` control, aspect preserved with `h-1`). Paths are resolved from `A_ScriptDir` first, then from the directory containing `Utils.ahk` (`A_LineFile`). If no image loads, a bold **EN** / **PT** text label is shown on the same panel.
- **Code-switch hook:** `ExecuteHandyAiModelSelection` calls `LanguageFlag_Show(3)` / `LanguageFlag_Show(4)` on success for slots 3/4 and `LanguageFlag_Hide()` for slots 1/2, immediately after `Handy_SetPersistedAiModelSlot`.

## Semantic Colors (Colorblind Accessibility)

Accent colors are applied to the **border** only; the overlay background stays dark (`1E1E2E`). Global constants in `Utils.ahk` define semantic accent colors suitable for common color vision deficiencies:

| Constant                     | Hex    | Meaning               | Use for                                                |
| ---------------------------- | ------ | --------------------- | ------------------------------------------------------ |
| `BANNER_ACCENT_SUCCESS`      | 27AE60 | Dark green (positive) | Success confirmations, "Done", "activated"             |
| `BANNER_ACCENT_ERROR`        | C0392B | Red (negative)        | Errors, "not found", failures, activation failed       |
| `BANNER_ACCENT_INTERMEDIATE` | F1C40F | Yellow (general)      | Loading, in-progress, actionable prompts, neutral info |
| `BANNER_ACCENT_INFO`         | 2980B9 | Blue (distinct hue)   | Modes that must read differently from green/yellow (e.g. color-vision–friendly pairs with success) |

- **Information Only:** Use **success** for success messages, **error** for error/warning messages, **intermediate** for neutral (e.g. "Selecting X", "Sound ON/OFF").
- **Interactive Input:** Use **intermediate** (or neutral/dark `barColor` with `noBorder: true` for a minimal look).
- **Loading Indication:** Use **intermediate**.

All new banner call sites should pass one of these constants (e.g. as `bgColor` for `ShowCenteredOverlay_Utils`, or as `passiveBgColor` / `barColor` for `StandardLoadingBar_Show`).

## API Reference

### Core Functions

| Function                              | Signature                                                                                                                                  | Purpose                                                                                                                                 |
| ------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------ | --------------------------------------------------------------------------------------------------------------------------------------- |
| `StandardLoadingBar_Show`             | `(state, barColor, options)`                                                                                                               | Show overlay; options: passive, centerOnHwnd, textWidth, fontSize, alpha, passiveBgColor, noBorder, **promptKeys**, **trackActiveMonitor** |
| `StandardLoadingBar_Update`           | `(state, barColor)`                                                                                                                        | Update text/progress of visible bar (main message only; prompt strip is not updated)                                                    |
| `StandardLoadingBar_Hide`             | `(delayMs)`                                                                                                                                | Hide bar; `delayMs > 0` shows briefly before hiding                                                                                     |
| `StandardLoadingBar_ShowWithKeys`     | `(state, keyCallbacks, timeoutMs, centerOnHwnd, timeoutCallback, barColor, textWidth, fontSize, passiveBgColor, noBorder, **promptKeys**, trackActiveMonitor)` | Show with hotkey handlers; optional **trackActiveMonitor** (default false) to follow the foreground window's monitor while visible |
| `StandardLoadingBar_CloseKeysOverlay` | (internal)                                                                                                                                 | Unregister keys, cancel timeout, destroy overlay                                                                                        |

### Options Reference

| Option           | Type    | Default | Description                                                                                                                                |
| ---------------- | ------- | ------- | ------------------------------------------------------------------------------------------------------------------------------------------ |
| `passive`        | boolean | false   | Text-only mode (no progress bar animation)                                                                                                 |
| `centerOnHwnd`   | integer | 0       | Window to center on; 0 = active monitor                                                                                                    |
| `textWidth`      | integer | 0       | Overlay width in pixels; 0 = auto (60% of monitor width)                                                                                   |
| `fontSize`       | integer | **17**  | Font size in points (standard for all banners)                                                                                             |
| `alpha`          | integer | 235     | Window transparency (0–255)                                                                                                                |
| `passiveBgColor` | string  | ""      | Border accent color. Prefer `BANNER_ACCENT_SUCCESS` / `BANNER_ACCENT_ERROR` / `BANNER_ACCENT_INTERMEDIATE`. Overlay background stays dark. |
| `noBorder`       | boolean | false   | Skip yellow border frame (single GUI); used for dictation confirm and other minimal interactive banners                                    |
| `promptKeys`     | string  | ""      | Optional fixed bottom strip text (e.g. "[Y] Confirm [N] Cancel"); Interactive Input category                                               |
| `trackActiveMonitor` | boolean | false | When true with `centerOnHwnd` 0, starts a short timer to reposition the overlay when the foreground window's monitor changes (dictation/Gemini/Cursor-transfer prompts). Stopped when the bar hides. |

## Helper Wrappers (Utils.ahk)

These wrap `StandardLoadingBar_*` with preset styles:

| Function                                                    | Purpose                                                                                                                                                          |
| ----------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `AiModelBanner_Show` / `AiModelBanner_Hide`                 | AI model selection (textWidth 450, fontSize 17)                                                                                                                  |
| `ClipAngelBanner_Show` / `ClipAngelBanner_Hide`             | Clip Angel (textWidth 200, fontSize 17)                                                                                                                          |
| `FastCopyModeBanner_Show` / `FastCopyModeBanner_Update` / `FastCopyModeBanner_Hide` | Fast Copy Mode in Shift keys (textWidth 480, fontSize 17, `BANNER_ACCENT_INFO`, `promptKeys`, `trackActiveMonitor`)                                              |
| `ShowSingleCharTabBanner_Utils(tabNumber)`                  | Tab number (1 or 2); auto-hides after 700 ms; **fontSize 72** (excluded from 17px standard)                                                                      |
| `ShowCenteredOverlay_Utils(text, duration, bgColor)`        | Short message with duration; Show + Hide(duration); fontSize 17; message should start with emoji                                                                 |
| `HotstringGeminiBanner_Show` / `HotstringGeminiBanner_Hide` | Gemini redirect (textWidth 280, fontSize 17); default text with emoji                                                                                            |
| `DictationGeminiConfirm_ShowAndWait()`                      | "❓ Send transcription to Gemini? (6s)" with Y/N keys, prompt strip `[Y] Confirm  [N] Cancel`, 6 s timeout; `noBorder: true`; `barColor` `"1E1E2E"`; fontSize 17 |
| `LanguageFlag_Show(slot)` / `LanguageFlag_Hide()`           | **Persistent Indicator**: pinned, borderless opaque flag (or EN/PT fallback) at `LANGUAGE_FLAG_WIDTH` px wide (~30% smaller than 64); bottom-right of active monitor (slot 3 = UK, slot 4 = Brazil); follows foreground window across monitors; independent of `StandardLoadingBar_*` |

## Implementation Instances

### Utils.ahk

| Lines     | Context                                                                                                                                                                                                                                             |
| --------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 1753–1761 | Semantic color globals (`BANNER_ACCENT_*`), `g_StandardLoadingBarGui`, `g_StandardLoadingBarValue`, `g_StandardLoadingBarIsKeysOverlay`, `g_StandardLoadingBarKeysHotkeys`, `g_StandardLoadingBarKeysTimeoutTimer`, `g_StandardLoadingBarBorderGui` |
| 1764–1823 | `GetWorkAreaForWindow_StandardBar`, `GetActiveMonitorWorkArea_StandardBar`                                                                                                                                                                          |
| 1826–1912 | `StandardLoadingBar_Show`, `StandardLoadingBar_Tick`                                                                                                                                                                                                |
| 1929–1972 | `StandardLoadingBar_Update`, `StandardLoadingBar_Hide`, `StandardLoadingBar_CloseKeysOverlay`                                                                                                                                                       |
| 2010–2076 | `StandardLoadingBar_ShowWithKeys`, `StandardLoadingBar_RegisterKeyHandler` (`$*key`, `#InputLevel 10`), `StandardLoadingBar_KeysSelectionPoll`, `StandardLoadingBar_KeyWrapper`, `StandardLoadingBar_KeysTimeoutFired`                              |
| 1409–1454 | `AiModelBanner_Show`/`Hide`, `ClipAngelBanner_Show`/`Hide`, `ShowSingleCharTabBanner_Utils`                                                                                                                                                         |
| 2303–2316 | `FastCopyModeBanner_Show` / `FastCopyModeBanner_Update` / `FastCopyModeBanner_Hide` (Shift keys Fast Copy Mode)                                                                                                                                     |
| 1738–1745 | `ShowCenteredOverlay_Utils`                                                                                                                                                                                                                         |
| 2081–2152 | `HotstringGeminiBanner_Show`/`Hide`, `DictationGeminiConfirm_ShowAndWait` (uses `StandardLoadingBar_ShowWithKeys` with `promptKeys` `"[Y] Confirm  [N] Cancel"`, `noBorder: true`, `barColor` `"1E1E2E"`)                                           |
| 1679–1690 | Persistent language flag globals (`g_LanguageFlagGui`, `g_LanguageFlagSlot`, `g_LanguageFlagTrackTimer`, `g_LanguageFlagLastForegroundMonitorIdx`) and constants (`LANGUAGE_FLAG_WIDTH` default 45, `LANGUAGE_FLAG_MARGIN`, `LANGUAGE_FLAG_TRACK_INTERVAL`); deferred `SetTimer(LanguageFlag_InitFromPersistedSlot, -250)` for Reload-safe restoration |
| 3171–3310 | `LanguageFlag_GetImagePath`, `LanguageFlag_Show`, `LanguageFlag_Hide`, `LanguageFlag_RepositionToActiveMonitor`, `LanguageFlag_TrackTick`, `LanguageFlag_InitFromPersistedSlot` (Persistent Indicator; borderless, no margins, no window-level transparency, `LANGUAGE_FLAG_WIDTH` default 45, bottom-right anchor, independent track timer, EN/PT fallback if images missing) |
| 3375–3381 | `ExecuteHandyAiModelSelection` code-switch hook: after `Handy_SetPersistedAiModelSlot(selection)`, calls `LanguageFlag_Show(3)` / `LanguageFlag_Show(4)` for Cohere slots and `LanguageFlag_Hide()` for slots 1/2 |
| 5493–5500 | Peek PDF flow                                                                                                                                                                                                                                       |

### Gemini.ahk

| Lines     | Context                                                                                                |
| --------- | ------------------------------------------------------------------------------------------------------ |
| 172–206   | `ShowSmallLoadingIndicator` / `HideSmallLoadingIndicator` wrappers (Information Only / Loading)        |
| 235–264   | Async TTS state display                                                                                |
| 390–500   | Read aloud flow                                                                                        |
| 742–828   | First-time init (Opening Gemini, Sending prompt)                                                       |
| 914–959   | `#!+8` pronunciation: Interactive Input (2s `ShowWithKeys` + progress); on auto-detect timeout, **Loading** `⏳ Detecting language…` during IPC/daemon; then `GeminiAsyncLookup` **Loading** `⏳ Loading…` |
| 946–1107  | Async lookup/TTS loading                                                                               |
| 1220      | `ShowWithKeys` for pronunciation completion (close keys)                                               |
| 1317–1318 | `ShowWithKeys` "❓ Copy response?" with `promptKeys` `"[Y] Copy  [N] No  [R] Copy+Read  [C] Transfer"` |
| 1328–1449 | Additional loading states                                                                              |

### Shift keys.ahk

| Lines       | Context                                          |
| ----------- | ------------------------------------------------ |
| 3150–3196   | Wikipedia restore scroll (Loading Indication)    |
| 3248–3463   | Wikipedia save scroll (Loading Indication)       |
| 9936–10336  | Fold/Unfold Explorer                             |
| 15034–15038 | `ShowCenteredOverlay` wrapper (Information Only) |
| ~1556–1647  | Fast Copy Mode (`#!+j`, `FastCopyModeBanner_*`, `#HotIf` copy hooks) |
| Outlook Reminders (`Reminders_SelectItem`) | **Interactive Input** (`StandardLoadingBar_ShowWithKeys`) for picking a reminder row. The main message may begin with an **Information Only**–style ℹ️ preamble (workaround when UI automation may not match on-screen rows), followed by the ❓ selection prompt and the numbered list. Do not use a second overlay on top of the keys modal for the same tip—embed the copy in the `ShowWithKeys` body so one surface stays readable. |

#### Outlook Reminders selection (Shift keys)

- **Category:** Interactive Input (`ShowWithKeys`, `promptKeys` e.g. `"[1-9/A-Z] Select  [Esc] Cancel"`).
- **Optional user guidance:** Prefix the main `state` text with a short ℹ️ line explaining that the list is derived from the Reminders window and what to do if rows look incomplete (focus/move window, retry). Keeps parity with the “emoji first” rule for the first line users read.
- **Pre-modal nudge:** Before opening the selection modal, Shift keys may apply a one-pixel `WinMove` and restore on the Reminders HWND so Outlook refreshes its accessible tree; the ℹ️ preamble can mention this so users are not surprised. Do not repeat the nudge on modal live-refresh timers.
- **Standalone information:** Use `ShowCenteredOverlay_Utils(..., BANNER_ACCENT_INFO)` only when there is no interactive banner yet (e.g. debug-only notices); avoid stacking two full-screen banners with `ShowWithKeys`.

### AppLaunchers.ahk

| Lines     | Context                          |
| --------- | -------------------------------- |
| `#!+h` YouTube focus | No banner on this hotkey (latency); `YouTube_PlayWhenOpened` uses `UIA_Browser("ahk_id " hwnd)` + `GetCurrentURL` then `Send("k")` when on a watch URL (assumes paused video on session start) |
| 522–563   | Restore scroll (short path)      |
| 842–1003  | Restore scroll (new window, UIA) |
| 1296–1392 | Restore scroll (existing window) |
| 1912–1941 | Save scroll position             |

### Act.ahk

| Lines | Context                                                                                        |
| ----- | ---------------------------------------------------------------------------------------------- |
| 20–65 | Startup sequence (Updating scripts, Updating notes, Launching apps, Done) – Loading Indication |

### WindowManagement.ahk

| Lines                                                                                                                                                                 | Context                                                                                                                                                                          |
| --------------------------------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 34–36                                                                                                                                                                 | `ShowNotification_WM(message, durationMs)` – wraps `ShowCenteredOverlay_Utils(message, durationMs, BANNER_ACCENT_ERROR)`; Information Only category for errors and notifications |
| `#!+w` (`WM_WindowTools_ShowMenu`)                                                                                                                                   | **Interactive Input** menu via `StandardLoadingBar_ShowWithKeys` (`[1]` lone-window maximize per monitor, `[2]` minimized background list); option 2 opens a dark **numbered text selector** (`WM_ShowMinimizedBackgroundList`: keys `1`–`9`, `0` = 10th, then `a`–`z` to close windows; after each key the list **re-scans** (`WM_CollectBackgroundWindows`), drops the closed HWND, **re-renders** with keys renumbered from `[1]`, and keeps the modal open until empty or Esc; modal uses **active-monitor tracking** (`WM_MinimizedList_TrackActiveMonitorTick`, 115 ms — same as `trackActiveMonitor` / `StudyTopicSelector`) so it stays centered on the monitor of the foreground window while `MonitorActiveWindow` recenters the cursor; Esc cancels via `g_OnEscapePressed`, `$*Escape` at `#InputLevel 10`, Esc poll, Gui `Escape`, and `.cursor/wm_minimized_list_open` IPC so **AppLaunchers** `Utils_GlobalEscapeHandler` requests close instead of forwarding Escape); empty list uses **Information Only** `ShowCenteredOverlay_Utils` with `BANNER_ACCENT_INFO` |
| 41–51, 116, 375, 383, 387, 432, 516, 644, 651, 666, 677, 684, 699, 802, 1039, 1061, 1124, 1133, 1155, 1349, 1374–1386, 1412, 1619, 1645, 1675, 1709, 1859, 1868, 2214 | Call sites for `ShowNotification_WM`                                                                                                                                             |

### Microsoft Teams.ahk

| Lines   | Context                                                   |
| ------- | --------------------------------------------------------- |
| 199–205 | `ShowCenteredOverlay` wrapper (Show + Hide with duration) |
| 254–255 | Usage (Information Only)                                  |

## Lifecycle and Best Practices

1. **Show → Update → Hide** – Call `Show` at start, `Update` at milestones, `Hide` in all exit paths (including `try`/`finally` and error branches).
2. **Delayed hide** – Use `Hide(delayMs)` to show a final message briefly before hiding.
3. **Keys overlay** – `ShowWithKeys` registers hotkeys; `CloseKeysOverlay` or `Hide(0)` unregisters and destroys.
4. **Include Utils** – Scripts that use the bar must include `Utils.ahk` (`#Include %A_ScriptDir%\Utils.ahk`).
5. **No stuck bar** – Ensure every code path that calls `Show` eventually calls `Hide`.
6. **Font size 17** – Use default `fontSize` 17 for all new banners; only `ShowSingleCharTabBanner_Utils` keeps 72.
7. **Emoji** – Start every banner message with an appropriate emoji (e.g. ⏳ loading, ✅ done, ❌ error, ❓ user input).
8. **Interactive Input** – When using `ShowWithKeys`, pass the 11th parameter `promptKeys` (e.g. `"[Y] Confirm  [N] Cancel"`) for a fixed bottom strip.
9. **Background work** – Any step that can take noticeable time without direct user input (daemon or IPC startup, language detection, network, browser automation) must show **Loading Indication**: call `StandardLoadingBar_Show` (animated bar, default `passive: false`) before the work begins, then `Update` at milestones if helpful, and always reach `Hide` on failure branches. Do not leave the screen empty between closing an interactive banner (e.g. `ShowWithKeys`) and the next visible outcome unless the transition is instantaneous.
10. **Persistent indicators** – Each persistent indicator (e.g. `LanguageFlag_*`) must own a **separate** track-timer global; do not reuse `g_StandardLoadingBarTrackTimer` or any other transient banner's tracker. Sharing would let any `Show`/`Hide` cycle on a transient banner silently kill the persistent indicator's monitor-follow behavior. Persistent indicators should render with **no border or padding** (window equals the inner control's pixel size: `MarginX = MarginY = 0`, no `Border` option on the `Gui(...)` call or the `Picture`/`Text` control). The language flag stays **opaque by default** (no `WinSetTransparent` unless a design explicitly needs alpha). Show with `NA` on a `+ToolWindow +AlwaysOnTop -Caption` window so the chip never activates the foreground app. Do **not** use `WS_EX_TRANSPARENT` (`+E0x80020`) for visibility — it suppresses painting entirely; click-through is optional and only acceptable if implemented without breaking normal client-area painting.

## Consumption by tools

This document is the single source of truth for banner and information-display behavior. When suggesting or generating code that shows loading, information, or interactive prompts, use the three display categories (Loading Indication, Information Only, Interactive Input) and the same function and option names documented here. Prefer `StandardLoadingBar_*` and `ShowCenteredOverlay_Utils` over ad-hoc overlays or `MsgBox`.

## Related Documentation

- [efficiency-canon.md](efficiency-canon.md) – Strategic guidelines; AI agents should read before refactors.
- [asynchronous_workflow_standards.md](asynchronous_workflow_standards.md) – Submit → monitor → retrieve pattern.
- [cheat-sheet.md](cheat-sheet.md) – ShiftKeys cheat sheet (authoring format, registry, search).
