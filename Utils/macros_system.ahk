; =============================================================================
; Utils module: macros_system.ahk
; Macros system RegisterMacro and assignments
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; Macros System
; =============================================================================

; Global variables for macros
global g_Macros := []
global g_MacroCharMap := Map()  ; Maps character to macro function
global g_ProgrammaticDictationStop := false  ; Skip ~#!+0 when script sends #!+0 programmatically
global g_GeminiToggleTab := 1  ; Last tab chosen by ^!#4 ToggleAICompanionChromeTab (UIA-synced when possible)

; Register a macro
RegisterMacro(func, title, char := "") {
    global g_Macros
    g_Macros.Push({ func: func, title: title, category: "Macros", char: char })
}

; Get scripts directory path based on environment
GetScriptsDirectory() {
    global IS_WORK_ENVIRONMENT
    if (IS_WORK_ENVIRONMENT) {
        return "C:\Users\fie7ca\Documents\scripts"
    } else {
        return "C:\Users\eduev\Meu Drive\17 - Projects\scripts"
    }
}

; Get list of script files to update.
; AppLaunchers.ahk must be last so QuickUpdateScripts can relaunch it with /Updated (success overlay; includes Utils.ahk).
GetScriptFiles() {
    scriptsDir := GetScriptsDirectory()
    return [
        scriptsDir "\WindowManagement.ahk",
        scriptsDir "\Spotify.ahk",
        scriptsDir "\Shift keys.ahk",
        scriptsDir "\Outlook.ahk",
        scriptsDir "\Microsoft Teams.ahk",
        scriptsDir "\Gemini.ahk",
        scriptsDir "\mousemaster.ahk",
        scriptsDir "\AppLaunchers.ahk"
    ]
}

; Quick Update Scripts macro: PowerShell handoff restart (local-only, no git).
QuickUpdateScripts() {
    static s_isQuickUpdateRunning := false
    if (s_isQuickUpdateRunning) {
        return
    }
    s_isQuickUpdateRunning := true

    try {
        ; Do not call ApplyScriptMasterVolumeTarget here - it only affects sessions that are about to be killed,
        ; then ExitApp cancels timers; new processes need volume applied after relaunch (see /Updated in Utils).
        scripts := GetScriptFiles()
        scriptsNames := ""
        for scriptPath in scripts {
            parts := StrSplit(scriptPath, "\")
            scriptsNames .= parts[parts.Length] . ";"
        }

        StandardLoadingBar_Show("⏳ Restarting scripts...", BANNER_ACCENT_INTERMEDIATE, { passive: false })

        anchorPath := ""
        psPaths := []

        for scriptPath in scripts {
            if (InStr(scriptPath, "\AppLaunchers.ahk"))
                anchorPath := scriptPath
            psPaths.Push("'" . StrReplace(scriptPath, "'", "''") . "'")
        }

        if (!anchorPath || anchorPath = "") {
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ AppLaunchers.ahk not found in script list", 2500, BANNER_ACCENT_ERROR)
            return
        }

        psAnchor := StrReplace(anchorPath, "'", "''")

        ; PowerShell contract (deterministic): sleep 500ms -> stop AutoHotkey* -> sleep 500ms -> Start-Process scripts.
        ps := ""
        ps .= "Start-Sleep -Milliseconds 500; "
        ps .= "Stop-Process -Name 'AutoHotkey*' -Force -ErrorAction SilentlyContinue; "
        ps .= "Start-Sleep -Milliseconds 500; "
        ps .= "$anchorPath = '" . psAnchor . "'; "
        ; AHK Array.Join() isn't available in this environment; build a CSV manually.
        psPathsJoined := ""
        for i, p in psPaths {
            if (i > 1)
                psPathsJoined .= ","
            psPathsJoined .= p
        }
        ps .= "$scripts = @(" . psPathsJoined . "); "
        ; Stagger launches so earlier scripts begin loading before the next; AppLaunchers stays last with /Updated.
        ; Short pause after all Start-Process so audio sessions can register before the relaunched AppLaunchers runs its volume schedule.
        ps .=
            "foreach ($s in $scripts) { if ($s -eq $anchorPath) { Start-Process -FilePath $s -ArgumentList '/Updated' | Out-Null } else { Start-Process -FilePath $s | Out-Null }; Start-Sleep -Milliseconds 450 }; "
        ps .= "Start-Sleep -Seconds 2; "

        ; Execute asynchronously, then terminate this AHK instance immediately.
        pid := 0
        try {
            pid := Run('powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "' . ps . '"', , "Hide")
        } catch as e {
            throw
        }
        ExitApp
    } catch as e {
        try StandardLoadingBar_Hide(0)
        ShowCenteredOverlay_Utils("❌ QuickUpdateScripts error: " . e.Message, 3500, BANNER_ACCENT_ERROR)
    } finally {
        ; If we didn't ExitApp (error path), reset the guard.
        s_isQuickUpdateRunning := false
    }
}

; Update Gemini script specifically (local-only): restart Gemini.ahk.
UpdateGeminiScript() {
    scriptsDir := GetScriptsDirectory()
    geminiPath := scriptsDir "\Gemini.ahk"

    if (!FileExist(geminiPath)) {
        ShowCenteredOverlay_Utils("❌ Gemini.ahk not found at: " geminiPath, 3000, BANNER_ACCENT_ERROR)
        return
    }

    StandardLoadingBar_Show("⏳ Reloading Gemini...", BANNER_ACCENT_INTERMEDIATE, { passive: false })
    try {
        Run(geminiPath)
        StandardLoadingBar_Hide(0)
        ShowCenteredOverlay_Utils("✅ Gemini script updated!", 1500, BANNER_ACCENT_SUCCESS)
    } catch {
        StandardLoadingBar_Hide(0)
        ShowCenteredOverlay_Utils("❌ Failed to reload Gemini", 2000, BANNER_ACCENT_ERROR)
    }
}

; Add specific word to Handy macro function
AddWordToHandy() {
    targetPath := GetHandyShortcutPath()

    try {
        if WinExist("Handy ahk_class Tauri Window") {
            WinActivate
        } else {
            if (targetPath = "" || !FileExist(targetPath)) {
                MsgBox "Failed to launch Handy.`n`nShortcut not found.", "Utils.ahk", "IconX"
                return
            }
            Run targetPath
            if !WinWait("Handy ahk_class Tauri Window", , 5) {
                MsgBox "Failed to launch Handy."
                return
            }
        }

        if (!WinWaitActive("Handy ahk_class Tauri Window", , 2)) {
            ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        hwnd := WinExist("Handy ahk_class Tauri Window")

        ; Initialize UIA
        el := UIA.ElementFromHandle(hwnd)
        if !el
            return

        ; Locate and click "Advanced" - find Group element directly by Type and ClassName pattern
        ; The clickable target is the Group (50026) with ClassName containing "cursor-pointer" and "flex gap-2 items-center"
        advancedBtn := ""
        try {
            allGroups := el.FindAll({ Type: 50026 })
            if allGroups {
                for group in allGroups {
                    try {
                        groupClassName := group.ClassName
                        ; Look for Group with cursor-pointer and flex gap-2 items-center (Advanced button pattern)
                        if (InStr(groupClassName, "cursor-pointer") && InStr(groupClassName, "flex gap-2 items-center")) {
                            ; Verify it contains "Advanced" text by checking children
                            try {
                                advancedText := group.FindFirst({ Type: 50020, Name: "Advanced" })
                                if advancedText {
                                    advancedBtn := group
                                    break
                                }
                            } catch {
                            }
                        }
                    } catch {
                    }
                }
            }
        } catch {
        }

        ; Fallback: Find by Name "Advanced" and verify parent has correct ClassName
        if !advancedBtn {
            advancedElement := el.FindFirst({ Name: "Advanced" })
            if advancedElement {
                try {
                    parentGroup := advancedElement.GetParentElement()
                    ; Verify parent is a Group with cursor-pointer class (clickable)
                    if (parentGroup && parentGroup.Type = 50026 && InStr(parentGroup.ClassName, "cursor-pointer")) {
                        advancedBtn := parentGroup
                    }
                } catch {
                }
            }
        }

        if advancedBtn {
            try {
                advancedBtn.Click()
            } catch Error as clickErr {
                try {
                    advancedBtn.Invoke()
                } catch {
                }
            }
        }

        Sleep 200

        ; Locate and focus "Add a word" text field
        addWordEdit := el.FindFirst({ Type: 50004, Name: "Add a word" })
        if addWordEdit {
            addWordEdit.SetFocus()
        }
    } catch Error as e {
        MsgBox "Error in AddWordToHandy macro: " e.Message
    }
}

; Resolve Handy shortcut/executable path (environment-aware: work vs home)
; Work: uses Documents\Handy\handy.exe; Home: uses Start Menu shortcuts.
GetHandyShortcutPath() {
    global IS_WORK_ENVIRONMENT

    if (IS_WORK_ENVIRONMENT) {
        ; Work: direct exe path first, then work shortcut fallback
        workExe := "C:\Users\fie7ca\Documents\Handy\handy.exe"
        if (FileExist(workExe))
            return workExe
        for , p in ["C:\Users\fie7ca\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Handy\Handy.lnk",
            "C:\Users\fie7ca\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Handy.lnk"] {
            if (p != "" && FileExist(p))
                return p
        }
        return ""
    }

    ; Home/personal: Start Menu shortcuts only
    candidates := [
        "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Handy\Handy.lnk",
        "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Handy\Handy.lnk",
        "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Handy.lnk"
    ]
    for , p in candidates {
        try {
            if (p != "" && FileExist(p))
                return p
        } catch {
        }
    }
    return ""
}

; Expected Handy process exe path for current environment (for targeting correct instance).
; Work: work exe path; Home: "" (any Handy window).
GetHandyProcessPath() {
    global IS_WORK_ENVIRONMENT
    if (IS_WORK_ENVIRONMENT) {
        workExe := "C:\Users\fie7ca\Documents\Handy\handy.exe"
        return FileExist(workExe) ? workExe : ""
    }
    return ""
}
