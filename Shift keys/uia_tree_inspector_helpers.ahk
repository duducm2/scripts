; =============================================================================
; Shift keys module: uia_tree_inspector_helpers.ahk
; UIATreeInspector UIA focus and jiggle helpers
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

;-------------------------------------------------------------------
; UIA Tree Inspector Shortcuts
;-------------------------------------------------------------------

; FindFirst/FindElement throw TargetError when no match; FindAll returns []. Use FindAll fallback for reliability.
UIATreeInspector_FindTreeByAutomationId(root, automationId) {
    if (!root)
        return 0
    aid := String(automationId)
    try
        return root.FindFirst({ Type: UIA.Type.Tree, AutomationId: aid })
    catch TargetError {
    }
    try {
        trees := root.FindAll({ Type: UIA.Type.Tree })
        for t in trees {
            if !t
                continue
            try {
                if (String(t.AutomationId) = aid)
                    return t
            } catch {
            }
        }
    } catch {
    }
    return 0
}

UIATreeInspector_FindRightDumpTree(root) {
    t := UIATreeInspector_FindTreeByAutomationId(root, "17")
    if (t)
        return t
    try {
        trees := root.FindAll({ Type: UIA.Type.Tree })
        for x in trees {
            if !x
                continue
            try {
                if (x.Name = "UIA Tree")
                    return x
            } catch {
            }
        }
    } catch {
    }
    bestL := -0x7FFFFFFF
    bestT := 0
    try {
        trees := root.FindAll({ Type: UIA.Type.Tree })
        for t in trees {
            if (!t)
                continue
            try {
                if (String(t.AutomationId) = "4")
                    continue
            } catch {
                continue
            }
            try {
                br := t.BoundingRectangle
                if (br.l > bestL) {
                    bestL := br.l
                    bestT := t
                }
            } catch {
                if (!bestT)
                    bestT := t
            }
        }
    } catch {
    }
    return bestT
}

; Win32 focus on left SysTreeView32 (AutomationId 4, TVWins) so UIA selection and arrow keys stay in sync.
UIATreeInspector_FocusLeftWindowsTree(treeContainer, winHwnd) {
    leftHwnd := 0
    if (!treeContainer || !winHwnd)
        return 0
    try
        leftHwnd := treeContainer.NativeWindowHandle
    catch {
        leftHwnd := 0
    }
    if (leftHwnd) {
        try
            ControlFocus "ahk_id " leftHwnd, "ahk_id " winHwnd
        catch {
            ; Invalid HWND pair or control not targetable; SetFocus below may still work.
        }
    }
    try
        treeContainer.SetFocus()
    catch {
        ; UIA SetFocus can surface COM/Win32 errors (e.g. "Target window not found"); non-fatal.
    }
    return leftHwnd
}

; Jiggle selection with keys guaranteed to go to the windows list TreeView (not filter / middle panels).
UIATreeInspector_JiggleLeftTree(leftHwnd, winHwnd, downDelayMs) {
    if (!WinExist("ahk_id " winHwnd))
        return
    if (!WinActive("ahk_id " winHwnd))
        return
    if (leftHwnd) {
        try {
            ControlSend "{Down}", "ahk_id " leftHwnd, "ahk_id " winHwnd
            Sleep downDelayMs
            ControlSend "{Up}", "ahk_id " leftHwnd, "ahk_id " winHwnd
        } catch {
            Send "{Down}"
            Sleep downDelayMs
            Send "{Up}"
        }
    } else {
        Send "{Down}"
        Sleep downDelayMs
        Send "{Up}"
    }
}
