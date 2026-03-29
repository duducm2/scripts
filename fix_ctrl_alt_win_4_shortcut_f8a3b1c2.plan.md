---
name: Fix Ctrl Alt Win 4 Shortcut
overview: Resolve dictation trigger failures and tab synchronization issues in the Ctrl+Alt+Win+4 shortcut, introducing UIA validation.
todos:
  - id: fix_dictation_trigger
    content: Update the dictation trigger in ^!#4:: to use SendLevel 1 to properly invoke the ~#!+0:: hotkey.
    status: pending
  - id: implement_uia_tab_sync
    content: Replace blind g_GeminiToggleTab toggling with UIA-based active tab detection to reliably determine whether to switch to tab 1 or 2.
    status: pending
    dependencies: [fix_dictation_trigger]
  - id: add_automated_quality_checks
    content: Add validation logic after execution to verify that g_DictationActive is true and the UIA active tab index matches the intended target, displaying an error banner if either fails.
    status: pending
    dependencies: [implement_uia_tab_sync]
---

# Fix Ctrl Alt Win 4 Shortcut

## Analysis / Context
The `^!#4::` hotkey currently suffers from two distinct failure modes:
1. The `Send("#!+0")` command often fails to trigger the `~#!+0::` dictation hotkey because default `SendLevel` limits recursive hotkey invocation.
2. The script relies on a blind state variable (`g_GeminiToggleTab`) to switch between Gemini tabs (Ctrl+1 / Ctrl+2). If the user manually changes tabs, this variable desynchronizes, causing the shortcut to switch to the wrong tab.

## Proposed Changes
- **Dictation Fix:** Elevate the `SendLevel` to 1 right before sending `#!+0`, ensuring the AutoHotkey engine intercepts it as a hotkey trigger.
- **Tab Sync Fix:** Initialize a `UIA_Browser` object for the Gemini window and invoke the existing `GetChromeActiveTabIndex(uia)` helper. Use the returned index to accurately determine the alternate tab (if on tab 1, go to 2; if on 2, go to 1).
- **Quality Checks:** Introduce a verification block at the end of the hotkey execution. It will check if `g_DictationActive` is true and use `GetChromeActiveTabIndex` again to ensure the tab switched successfully.

## Files to Modify
- `c:\Users\eduev\Meu Drive\17 - Projects\scripts\Utils.ahk` (around line 1554: `^!#4::`)

## Implementation Strategy
1. Locate the `^!#4::` hotkey definition.
2. Replace `Send("#!+0")` with:
   ```ahk
   SendLevel 1
   Send "#!+0"
   SendLevel 0
   ```
3. After activating the `geminiHwnd` window, instantiate `uia := UIA_Browser("ahk_id " geminiHwnd)`.
4. Call `tabInfo := GetChromeActiveTabIndex(uia)`.
5. Determine `targetTab`: if `tabInfo` is valid and `tabInfo.index` is 1, set `targetTab` to 2; otherwise, set to 1. Update `g_GeminiToggleTab` with this value.
6. Send `^` + `targetTab` and call `ShowSingleCharTabBanner_Utils(targetTab)`.
7. Sleep briefly (e.g., 200ms) to allow the UI and dictation state to update.
8. Perform quality checks: 
   - Verify `g_DictationActive == true`. 
   - Call `tabInfoAfter := GetChromeActiveTabIndex(uia)` and verify `tabInfoAfter.index == targetTab`. 
   - If either fails, trigger `ShowCenteredOverlay_Utils("❌ Shortcut execution failed", 2000, BANNER_ACCENT_ERROR)`.
