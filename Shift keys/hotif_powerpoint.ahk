; =============================================================================
; Shift keys module: hotif_powerpoint.ahk
; PowerPoint hotkeys
; Loaded via #include into the Shift keys.ahk process.
; =============================================================================

PowerPoint_IsSaveAsDialogTitle(title) {
    return InStr(title, "Save As")
    || InStr(title, "Salvar como")
    || InStr(title, "Guardar como")
}

PowerPoint_WaitSaveAsDialog(timeoutMs := 5000) {
    deadline := A_TickCount + timeoutMs
    prevMatchMode := A_TitleMatchMode
    SetTitleMatchMode 2
    try {
        while (A_TickCount < deadline) {
            hwnd := WinExist("ahk_class #32770")
            if hwnd {
                title := ""
                try title := WinGetTitle("ahk_id " hwnd)
                catch {
                }
                if PowerPoint_IsSaveAsDialogTitle(title)
                    return hwnd
            }
            for str in ["Save As", "Salvar como", "Guardar como"] {
                hwnd := WinExist(str " ahk_class #32770")
                if hwnd
                    return hwnd
            }
            Sleep 100
        }
    } finally {
        SetTitleMatchMode prevMatchMode
    }
    return 0
}

PowerPoint_GetActivePresentationFolder() {
    try {
        pp := ComObjActive("PowerPoint.Application")
        path := pp.ActivePresentation.Path
        if (path != "")
            return path
    } catch {
    }
    return ""
}

PowerPoint_SaveAsPdf() {
    targetFolder := PowerPoint_GetActivePresentationFolder()
    if (targetFolder = "")
        targetFolder := A_Desktop

    Send "{F12}"
    hwnd := PowerPoint_WaitSaveAsDialog(5000)
    if !hwnd
        return

    try WinActivate("ahk_id " hwnd)
    catch {
    }
    Sleep 150

    try {
        FileDialog_NavigateToFolder(targetFolder)
        Sleep 100
        ; Folder change can rebuild dialog controls — refresh UIA root.
        root := UIA.ElementFromHandle(hwnd)
        if !FileDialog_SelectPdf(root)
            return
        Sleep 80
        FileDialog_ClickSaveButton(root)
        Sleep 200
        FileDialog_HandlePostSaveDialogs()
    } catch {
    }
}

;-------------------------------------------------------------------
; PowerPoint Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe POWERPNT.EXE") && WinGetClass("A") != "#32770"

; Shift + P : Save as PDF (same folder, or Desktop if unsaved)
+p:: PowerPoint_SaveAsPdf()

#HotIf