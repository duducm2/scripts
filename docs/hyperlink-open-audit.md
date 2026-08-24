# Hyperlink Open Locations — Structured Report

Scope: **explicit URL launches only** (`Run`, `chrome.exe`, `StudyLink_OpenUrlInChrome`, and `FindAndActivateMiroWindow` fallback open).

Last updated: 2026-05-24 (audit implemented in code for flagged rows).

---

## Summary

| Category                            | Count                                     | Already uses `--new-window`              |
| ----------------------------------- | ----------------------------------------- | ---------------------------------------- |
| Study links (API)                   | 2 call sites (+ 1 shared helper)          | Article + YouTube (after implementation) |
| Utility / Miro                      | 1 open site (+ quick-open via helper)     | Miro fallback + article                  |
| Gemini                              | 2                                         | Yes                                      |
| App launchers (YouTube / Wikipedia) | 2                                         | Yes                                      |
| Legacy / helper                     | 2 (`StudyLink_Open`, `Run(url)` fallback) | `StudyLink_Open` updated                 |

Central helper: `StudyLinkHelpers.ahk` — `StudyLink_OpenUrlInChrome` (lines 242–249).

---

## Structured report

| #   | Code location                                                                 | Current behavior                                                              | Target destination                                                 | Selection flag                 |
| --- | ----------------------------------------------------------------------------- | ----------------------------------------------------------------------------- | ------------------------------------------------------------------ | ------------------------------ |
| 1   | `StudyArticleLink.ahk` **101** — `StudyTopicSelector_ManageArticleLinks_Open` | `StudyLink_OpenUrlInChrome(url, true)` → `chrome.exe --new-window "<url>"`    | Stored **article** URL from API (`STUDYLINK_KEY_ARTICLE`)          | `[ ]` Already isolated         |
| 2   | `Utils.ahk` **9036** — `StudyTopicSelector_ManageLinks_Open`                  | `StudyLink_OpenUrlInChrome(url, true)` (implemented)                          | Stored **YouTube subtopic** URL from API (`STUDYLINK_KEY_YOUTUBE`) | `[x]` Separate window          |
| 3   | `StudyLinkHelpers.ahk` **246** — `StudyLink_OpenUrlInChrome` (primary path)   | `Run(chromeCmd)` with optional `--new-window`                                 | Dynamic URL from callers                                           | `[ ]` Follows caller           |
| 4   | `StudyLinkHelpers.ahk` **248** — `StudyLink_OpenUrlInChrome` (catch fallback) | `Run(url)` — OS default handler                                               | Same as #3 on Chrome failure                                       | `[ ]` Fallback only            |
| 5   | `StudyLinkHelpers.ahk` **344** — `StudyLink_Open`                             | `StudyLink_OpenUrlInChrome(r["url"], true)` (implemented)                     | URL for arbitrary `studyKey` (no in-repo callers)                  | `[x]` Separate window          |
| 6   | `Utils.ahk` **11037** — `FindAndActivateMiroWindow` (fallback)                | `Run("chrome.exe --new-window " . url)`                                       | Dynamic Miro board URL                                             | `[ ]` Already isolated on open |
| 7   | `Utils.ahk` **11544** — `HandleHotstringChar` (Links, char `9`)               | `FindAndActivateMiroWindow` — activate if exists, else #6                     | `https://miro.com/app/board/uXjVJdbNFkA=/`                         | `[ ]` Activate-or-open         |
| 8   | `Utils.ahk` **11548** — `HandleHotstringChar` (Links, char `0`)               | Same as #7                                                                    | `https://miro.com/app/board/uXjVJVZSXvk=/`                         | `[ ]` Activate-or-open         |
| 9   | `Utils.ahk` **11599** — `TryRunFile`                                          | HTTP(S) → `StudyLink_OpenUrlInChrome(fp, true)`; else `Run(fp)` (implemented) | Quick-open URLs in `InitQuickOpenFiles` (**921–935**)              | `[x]` Separate window          |
| 10  | `Utils.ahk` **11269** — `GeminiNavigateFocusAndPasteFirstSnippet`             | `chrome.exe --new-window` when no Gemini window                               | `https://gemini.google.com/`                                       | `[ ]` Already isolated         |
| 11  | `Gemini.ahk` **1083** — `InitializeGeminiFirstTime`                           | `chrome.exe --new-window` (two Gemini tabs)                                   | `https://gemini.google.com/`                                       | `[ ]` Already isolated         |
| 12  | ~~`AppLaunchers` `#!+h` YouTube History~~                                     | _(removed)_ — `#!+h` now opens Utility Shortcuts → Prompts                    | —                                                                  | `[x]` Repurposed 2026-08-24    |
| 13  | `AppLaunchers.ahk` **1006** — `HandleWikipediaChar`                           | `chrome.exe --new-window` + `item.url`                                        | Hardcoded Wikipedia URLs (**482–485**)                             | `[ ]` Already isolated         |

---

## Implementation log

| #   | Change                                                                                          |
| --- | ----------------------------------------------------------------------------------------------- |
| 2   | `StudyTopicSelector_ManageLinks_Open` passes `true` to `StudyLink_OpenUrlInChrome`              |
| 5   | `StudyLink_Open` uses `StudyLink_OpenUrlInChrome` instead of bare `Run(url)`                    |
| 9   | `TryRunFile` routes `http://` / `https://` paths through `StudyLink_OpenUrlInChrome(..., true)` |

---

## Out of scope

| Excluded                                       | Reason                             |
| ---------------------------------------------- | ---------------------------------- |
| `Shift keys.ahk` 21447–21490                   | `UIA.Navigate` / address-bar paste |
| `AppLaunchers.ahk` 311, `Shift keys.ahk` 16743 | Chrome launch with no URL          |
| `Act.ahk` / WhatsApp `.lnk`                    | Shortcut launch, not raw URL       |
| `StudyLinkHelpers.ahk` API URLs                | HTTP to Apps Script                |
| `docs/study-link-lightweight-api-setup.md`     | Documentation only                 |

---

## Reference: shared helper

```ahk
StudyLink_OpenUrlInChrome(url, newWindow := false) {
    if (Trim(url) = "")
        return false
    chromeCmd := newWindow ? 'chrome.exe --new-window "' url '"' : 'chrome.exe "' url '"'
    try Run(chromeCmd)
    catch
        try Run(url)
    return true
}
```
