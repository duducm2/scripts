---
name: Local Script Quick Update Refactor
overview: Refactor the Quick Update macro to rely exclusively on local file reloading, guarantee script termination using PowerShell, and implement standard progress bar feedback.
todos:
  - id: remove_non_local_logic
    content: In `Utils.ahk`, delete the `CheckScriptsNeedingUpdates()` and `QuickUpdate_ShutdownRunningScripts()` functions entirely. In `QuickUpdateScripts()` and `UpdateGeminiScript()`, remove all logic that executes `git fetch` and `git pull`.
    status: pending
  - id: refactor_update_gemini
    content: Rewrite `UpdateGeminiScript()` to only restart `Gemini.ahk`. Remove the dialog checking for other scripts. At the start, call `StandardLoadingBar_Show("⏳ Reloading Gemini...", BANNER_ACCENT_INTERMEDIATE, {passive: false})`. Use `Run(geminiPath)` to restart it, then call `StandardLoadingBar_Hide(0)` and `ShowCenteredOverlay_Utils("✅ Gemini script updated!", 1500, BANNER_ACCENT_SUCCESS)`.
    status: pending
    dependencies: [remove_non_local_logic]
  - id: refactor_quick_update
    content: Rewrite `QuickUpdateScripts()` to use a PowerShell handoff. Start by calling `StandardLoadingBar_Show("⏳ Restarting scripts...", BANNER_ACCENT_INTERMEDIATE, {passive: false})`. Construct a PowerShell command string that sleeps 500ms, force stops `AutoHotkey*` processes, sleeps another 500ms, and then uses `Start-Process` to launch each script from `GetScriptFiles()`. Append `/Updated` to the argument list for `Utils.ahk`. Execute the command asynchronously via `Run` and immediately call `ExitApp`.
    status: pending
    dependencies: [remove_non_local_logic]
  - id: implement_success_feedback
    content: In `Utils.ahk`, locate the end of the global auto-execute section (e.g., right before the hotkey definitions begin, around `#!+Q::`). Add a check for `if (A_Args.Length > 0 && A_Args[1] = "/Updated")`. Inside the block, call `ShowCenteredOverlay_Utils("✅ Scripts updated and relaunched", 6500, BANNER_ACCENT_SUCCESS)` and play the `quick-update-success.wav` sound.
    status: pending
    dependencies: [refactor_quick_update]
---

# Local Script Quick Update Refactor

## Analysis / Context

The current Quick Update macro in `Utils.ahk` suffers from state persistence issues and execution failures on subsequent runs when other macros are active. This is primarily because the AHK-based process termination (`QuickUpdate_ShutdownRunningScripts`) fails to detect and kill scripts without active/visible windows, leaving instances running in a locked state. Additionally, the macro relies on non-local Git commands (`git fetch`/`git pull`) which introduces unnecessary complexity, failure points, and delays for purely local updates.

To resolve these issues and adhere strictly to the `efficiency-canon.md` and `standard_information_display.md` guidelines, we will remove the non-local Git logic and shift the termination and restart responsibilities to an asynchronous PowerShell handoff. This guarantees a clean termination of all AutoHotkey processes and a fresh restart, while preserving the user experience with standard visual loading and success banners.

## Proposed Changes

1. **Strip Non-Local Logic:** Remove `CheckScriptsNeedingUpdates`, `QuickUpdate_ShutdownRunningScripts`, and all Git invocations.
2. **Standard Information Display:** Integrate `StandardLoadingBar_Show` with `{passive: false}` to display an animated progress bar indicating that scripts are being restarted.
3. **PowerShell Handoff:** Construct a deterministic PowerShell command within `QuickUpdateScripts()` to aggressively terminate all AutoHotkey processes (`Stop-Process -Name 'AutoHotkey*' -Force`), pause briefly to release file locks, and then relaunch all local scripts via `Start-Process`.
4. **Stateful Restart Feedback:** To display the success banner _after_ the scripts have fully restarted, `Utils.ahk` will be launched with a special `/Updated` CLI argument. The auto-execute section will intercept this argument and trigger the success banner.

## Files to Modify

- `Utils.ahk`

## Implementation Strategy

### Step 1: Clean Up Legacy Functions

Locate and delete the following functions entirely:

- `CheckScriptsNeedingUpdates()`
- `QuickUpdate_ShutdownRunningScripts()`

### Step 2: Refactor `UpdateGeminiScript()`

Remove the Git fetch/pull logic and the prompt that asks about other scripts needing updates. The simplified function should look like this:

```ahk
UpdateGeminiScript() {
    scriptsDir := GetScriptsDirectory()
    geminiPath := scriptsDir "\Gemini.ahk"

    if (!FileExist(geminiPath)) {
        MsgBox "Gemini.ahk not found at: " geminiPath, "Update Failed", "IconX"
        return
    }

    StandardLoadingBar_Show("⏳ Reloading Gemini...", BANNER_ACCENT_INTERMEDIATE, {passive: false})
    try {
        Run(geminiPath)
        StandardLoadingBar_Hide(0)
        ShowCenteredOverlay_Utils("✅ Gemini script updated!", 1500, BANNER_ACCENT_SUCCESS)
    } catch Error as e {
        StandardLoadingBar_Hide(0)
        ShowCenteredOverlay_Utils("❌ Failed to update Gemini", 2000, BANNER_ACCENT_ERROR)
    }
}
```
