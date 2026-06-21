; =============================================================================
; Shift keys module: vscode_commit_message.ahk
; VSCode_TriggerGenerateCommitMessage helper
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

VSCode_TriggerGenerateCommitMessage(hwnd := 0) {
    if (!hwnd)
        hwnd := WinExist("ahk_exe Code.exe")
    if (!hwnd)
        return false

    ; Ensure Source Control view is focused so the Generate button is rendered.
    FocusSourceControlViewForCommitGeneration()

    try root := UIA.ElementFromHandle(hwnd)
    catch
        return false
    if (!root)
        return false

    genBtn := 0
    for name in ["Generate Commit Message (Ctrl+Alt+.)", "Generate Commit Message"] {
        try {
            genBtn := root.FindFirst({ Type: 50000, Name: name })
        } catch {
            genBtn := 0
        }
        if (genBtn)
            break
    }

    if (!genBtn) {
        try {
            allButtons := root.FindAll({ Type: 50000 })
            if (allButtons) {
                for btn in allButtons {
                    try {
                        nm := btn.Name
                        if (InStr(nm, "Generate Commit Message")) {
                            genBtn := btn
                            break
                        }
                    } catch {
                        continue
                    }
                }
            }
        } catch {
            genBtn := 0
        }
    }

    if (!genBtn)
        return false

    try {
        if (genBtn.GetPropertyValue(UIA.Property.IsInvokePatternAvailable)) {
            genBtn.InvokePattern.Invoke()
            return true
        }
    } catch {
    }

    try {
        genBtn.Click()
        return true
    } catch {
        return false
    }
}
