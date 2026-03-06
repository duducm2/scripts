# Standard Loading Bar Component

A shared UI component for loading and progress feedback across all AHK scripts. Defined in `Utils.ahk` and used by scripts that include it.

## Overview and Purpose

- **Single shared component** for loading/progress feedback across all AHK scripts
- **Replaces ad-hoc banners and overlays** with a consistent user experience
- **Monitor-aware positioning** – centers on the active window's monitor or the primary monitor
- **Dual modes** – passive (text-only) display and interactive mode with key callbacks and timeouts
- **Standard font size** – 17px for all banners and loading indicators (except `ShowSingleCharTabBanner_Utils`, which keeps 72px)
- **Emoji** – every banner message must start with an emoji (e.g. ⏳ loading, ✅ success, ❌ error, ❓ user input)

## Banner Types

| Type            | Label / prefix   | Use case                                      |
| --------------- | ---------------- | --------------------------------------------- |
| **Loading**     | ⏳, 🔄           | Progress bar + emoji; passive: false          |
| **Information** | ✅, ℹ, 📋, ❌, ⚠ | Passive text-only; emoji + message            |
| **User input**  | ❓, ⌨            | `ShowWithKeys` + fixed bottom strip with keys |

User-input banners use an optional **fixed bottom strip** (e.g. `[Y] Confirm  [N] Cancel  [E] Close`) via the `promptKeys` option so the main message and key hints stay clearly separated.

## Semantic Colors (Colorblind Accessibility)

Accent colors are applied to the **border** only; the overlay background stays dark (`1E1E2E`). Three global constants in `Utils.ahk` define semantic accent colors suitable for common color vision deficiencies:

| Constant                     | Hex    | Meaning               | Use for                                                |
| ---------------------------- | ------ | --------------------- | ------------------------------------------------------ |
| `BANNER_ACCENT_SUCCESS`      | 27AE60 | Dark green (positive) | Success confirmations, "Done", "activated"             |
| `BANNER_ACCENT_ERROR`        | C0392B | Red (negative)        | Errors, "not found", failures, activation failed       |
| `BANNER_ACCENT_INTERMEDIATE` | F1C40F | Yellow (general)      | Loading, in-progress, actionable prompts, neutral info |

- **Informational banners**: Use **success** for success messages, **error** for error/warning messages, **intermediate** for neutral (e.g. "Selecting X", "Sound ON/OFF").
- **Actionable banners** (user input): Use **intermediate**.
- **Loading banners**: Use **intermediate**.

All new banner call sites should pass one of these constants (e.g. as `bgColor` for `ShowCenteredOverlay_Utils`, or as `passiveBgColor` / `barColor` for `StandardLoadingBar_Show`).

## API Reference

### Core Functions

| Function                              | Signature                                                                                                                                  | Purpose                                                                                                                                 |
| ------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------ | --------------------------------------------------------------------------------------------------------------------------------------- |
| `StandardLoadingBar_Show`             | `(state, barColor, options)`                                                                                                               | Show overlay; options: passive, centerOnHwnd, textWidth, fontSize, alpha, passiveBgColor, noBorder, **promptKeys** (fixed bottom strip) |
| `StandardLoadingBar_Update`           | `(state, barColor)`                                                                                                                        | Update text/progress of visible bar (main message only; prompt strip is not updated)                                                    |
| `StandardLoadingBar_Hide`             | `(delayMs)`                                                                                                                                | Hide bar; `delayMs > 0` shows briefly before hiding                                                                                     |
| `StandardLoadingBar_ShowWithKeys`     | `(state, keyCallbacks, timeoutMs, centerOnHwnd, timeoutCallback, barColor, textWidth, fontSize, passiveBgColor, noBorder, **promptKeys**)` | Show with hotkey handlers, optional timeout, and optional fixed bottom strip for key hints                                              |
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
| `noBorder`       | boolean | false   | Skip yellow border frame (single GUI); used for dictation confirm                                                                          |
| `promptKeys`     | string  | ""      | Optional fixed bottom strip text (e.g. "[Y] Confirm [N] Cancel"); user-input banners                                                       |

## Helper Wrappers (Utils.ahk)

These wrap `StandardLoadingBar_*` with preset styles:

| Function                                                    | Purpose                                                                                                                           |
| ----------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------- |
| `AiModelBanner_Show` / `AiModelBanner_Hide`                 | AI model selection (textWidth 450, fontSize 17)                                                                                   |
| `ClipAngelBanner_Show` / `ClipAngelBanner_Hide`             | Clip Angel (textWidth 200, fontSize 17)                                                                                           |
| `ShowSingleCharTabBanner_Utils(tabNumber)`                  | Tab number (1 or 2); auto-hides after 700 ms; **fontSize 72** (excluded from 17px standard)                                       |
| `ShowCenteredOverlay_Utils(text, duration, bgColor)`        | Short message with duration; Show + Hide(duration); fontSize 17; message should start with emoji                                  |
| `HotstringGeminiBanner_Show` / `HotstringGeminiBanner_Hide` | Gemini redirect (textWidth 280, fontSize 17); default text with emoji                                                             |
| `DictationGeminiConfirm_ShowAndWait()`                      | "❓ Send transcription to Gemini? (6s)" with Y/N keys, prompt strip `[Y] Confirm  [N] Cancel`, 6 s timeout; noBorder; fontSize 17 |

## Implementation Instances

### Utils.ahk

| Lines     | Context                                                                                                                                                                                                                   |
| --------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 1781–1786 | Global variables: `g_StandardLoadingBarGui`, `g_StandardLoadingBarValue`, `g_StandardLoadingBarIsKeysOverlay`, `g_StandardLoadingBarKeysHotkeys`, `g_StandardLoadingBarKeysTimeoutTimer`, `g_StandardLoadingBarBorderGui` |
| 1788–1849 | `GetWorkAreaForWindow_StandardBar`, `GetActiveMonitorWorkArea_StandardBar`                                                                                                                                                |
| 1851–1931 | `StandardLoadingBar_Show`, `StandardLoadingBar_Tick`                                                                                                                                                                      |
| 1949–1992 | `StandardLoadingBar_Update`, `StandardLoadingBar_Hide`, `StandardLoadingBar_CloseKeysOverlay`                                                                                                                             |
| 2029–2091 | `StandardLoadingBar_ShowWithKeys`, `StandardLoadingBar_RegisterKeyHandler`, `StandardLoadingBar_KeyWrapper`, `StandardLoadingBar_KeysTimeoutFired`                                                                        |
| 1408–1453 | `AiModelBanner_Show`/`Hide`, `ClipAngelBanner_Show`/`Hide`, `ShowSingleCharTabBanner_Utils`                                                                                                                               |
| 1767–1774 | `ShowCenteredOverlay_Utils`                                                                                                                                                                                               |
| 2097–2166 | `HotstringGeminiBanner_Show`/`Hide`, `DictationGeminiConfirm_ShowAndWait` (ShowWithKeys)                                                                                                                                  |
| 5493–5500 | Peek PDF flow                                                                                                                                                                                                             |

### Gemini.ahk

| Lines     | Context                                                            |
| --------- | ------------------------------------------------------------------ |
| 172–206   | `ShowSmallLoadingIndicator` / `HideSmallLoadingIndicator` wrappers |
| 235–264   | Async TTS state display                                            |
| 390–500   | Read aloud flow                                                    |
| 742–828   | First-time init (Opening Gemini, Sending prompt)                   |
| 946–1107  | Async lookup/TTS loading                                           |
| 1118      | `ShowWithKeys` for long-running state                              |
| 1211      | `ShowWithKeys` for "Copy? [N] [R] [E=close]" completion            |
| 1328–1449 | Additional loading states                                          |

### Shift keys.ahk

| Lines       | Context                       |
| ----------- | ----------------------------- |
| 3150–3196   | Wikipedia restore scroll      |
| 3248–3463   | Wikipedia save scroll         |
| 9936–10336  | Fold/Unfold Explorer          |
| 15034–15038 | `ShowCenteredOverlay` wrapper |

### AppLaunchers.ahk

| Lines     | Context                          |
| --------- | -------------------------------- |
| 522–563   | Restore scroll (short path)      |
| 842–1003  | Restore scroll (new window, UIA) |
| 1296–1392 | Restore scroll (existing window) |
| 1912–1941 | Save scroll position             |

### Act.ahk

| Lines | Context                                                                   |
| ----- | ------------------------------------------------------------------------- |
| 20–65 | Startup sequence (Updating scripts, Updating notes, Launching apps, Done) |

### Microsoft Teams.ahk

| Lines   | Context                                                   |
| ------- | --------------------------------------------------------- |
| 254–255 | `ShowCenteredOverlay` wrapper (Show + Hide with duration) |

## Lifecycle and Best Practices

1. **Show → Update → Hide** – Call `Show` at start, `Update` at milestones, `Hide` in all exit paths (including `try`/`finally` and error branches).
2. **Delayed hide** – Use `Hide(delayMs)` to show a final message briefly before hiding.
3. **Keys overlay** – `ShowWithKeys` registers hotkeys; `CloseKeysOverlay` or `Hide(0)` unregisters and destroys.
4. **Include Utils** – Scripts that use the bar must include `Utils.ahk` (`#Include %A_ScriptDir%\Utils.ahk`).
5. **No stuck bar** – Ensure every code path that calls `Show` eventually calls `Hide`.
6. **Font size 17** – Use default `fontSize` 17 for all new banners; only `ShowSingleCharTabBanner_Utils` keeps 72.
7. **Emoji** – Start every banner message with an appropriate emoji (e.g. ⏳ loading, ✅ done, ❌ error, ❓ user input).
8. **User-input banners** – When using `ShowWithKeys`, pass the 11th parameter `promptKeys` (e.g. `"[Y] Confirm  [N] Cancel"`) for a fixed bottom strip.

## Related Documentation

- [loading-bar-rollout-locations.md](../loading-bar-rollout-locations.md) – Rollout status and migration candidates
