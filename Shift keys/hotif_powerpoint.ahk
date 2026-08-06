; =============================================================================
; Shift keys module: hotif_powerpoint.ahk
; PowerPoint hotkeys
; Loaded via #include into the Shift keys.ahk process.
; =============================================================================
; Shift+P: F12 Save As → Windows Desktop (A_Desktop) → PDF → Save.
; Does not read presentation Path/FullName (SharePoint COM export is unreliable).

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
                try {
                    title := WinGetTitle("ahk_id " hwnd)
                    if PowerPoint_IsSaveAsDialogTitle(title) {
                        WinActivate("ahk_id " hwnd)
                        return hwnd
                    }
                } catch {
                }
            }
            Sleep 50
        }
    } finally {
        SetTitleMatchMode prevMatchMode
    }
    return 0
}

PowerPoint_SaveAsPdf() {
    Send "{F12}"
    hwnd := PowerPoint_WaitSaveAsDialog()
    if !hwnd {
        ShowCenteredOverlay_Utils("❌ Save As dialog not found", 2200, BANNER_ACCENT_ERROR)
        return
    }

    FileDialog_NavigateToDesktop()
    Sleep 150

    try {
        root := UIA.ElementFromHandle(hwnd)
        if !root {
            ShowCenteredOverlay_Utils("❌ Save As UIA root not found", 2200, BANNER_ACCENT_ERROR)
            return
        }
        if !FileDialog_SelectPdf(root) {
            ShowCenteredOverlay_Utils("❌ Could not select PDF type", 2500, BANNER_ACCENT_ERROR)
            return
        }
        FileDialog_ClickSaveButton(root)
        FileDialog_HandlePostSaveDialogs()
        ShowCenteredOverlay_Utils("📄 Saved PDF on Desktop", 1800, BANNER_ACCENT_SUCCESS)
    } catch Error as e {
        ShowCenteredOverlay_Utils("❌ Save as PDF failed`n" e.Message, 3000, BANNER_ACCENT_ERROR)
    }
}

;-------------------------------------------------------------------
; PowerPoint Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe POWERPNT.EXE") && WinGetClass("A") != "#32770"

; Shift + P : Save as PDF on Desktop (F12 → Desktop → PDF → Save)
+p:: PowerPoint_SaveAsPdf()

#HotIf