; =============================================================================
; Shift keys module: hotif_powerpoint.ahk
; PowerPoint hotkeys
; Loaded via #include into the Shift keys.ahk process.
; =============================================================================

; ppFixedFormatTypePDF = 2; ppSaveAsPDF = 32
PowerPoint_GetActivePresentation() {
    try {
        pp := ComObjActive("PowerPoint.Application")
        return pp.ActivePresentation
    } catch {
    }
    return 0
}

PowerPoint_GetActivePresentationFolder(pres := 0) {
    if !pres
        pres := PowerPoint_GetActivePresentation()
    if !pres
        return ""
    try {
        path := pres.Path
        if (path != "")
            return path
    } catch {
    }
    return ""
}

PowerPoint_PdfOutputPath(pres, folder) {
    name := "Presentation"
    try {
        n := pres.Name
        if (n != "")
            name := n
    } catch {
    }
    name := RegExReplace(name, "i)\.(pptx|ppt|pptm|ppsx|pps)$", "")
    return folder "\" name ".pdf"
}

PowerPoint_SaveAsPdf() {
    pres := PowerPoint_GetActivePresentation()
    if !pres {
        ShowCenteredOverlay_Utils("❌ PowerPoint COM unavailable", 2200, BANNER_ACCENT_ERROR)
        return
    }

    targetFolder := PowerPoint_GetActivePresentationFolder(pres)
    if (targetFolder = "")
        targetFolder := A_Desktop

    outPath := PowerPoint_PdfOutputPath(pres, targetFolder)
    if (outPath = "" || targetFolder = "") {
        ShowCenteredOverlay_Utils("❌ Could not resolve PDF path", 2200, BANNER_ACCENT_ERROR)
        return
    }

    try {
        ; ExportAsFixedFormat keeps the .pptx open; does not switch active format.
        ; FixedFormatType 2 = ppFixedFormatTypePDF
        pres.ExportAsFixedFormat(outPath, 2)
        ShowCenteredOverlay_Utils("📄 Saved PDF`n" outPath, 2200, BANNER_ACCENT_SUCCESS)
        return
    } catch {
    }

    try {
        ; Fallback: SaveAs with ppSaveAsPDF = 32
        pres.SaveAs(outPath, 32)
        ShowCenteredOverlay_Utils("📄 Saved PDF`n" outPath, 2200, BANNER_ACCENT_SUCCESS)
        return
    } catch Error as err {
        ShowCenteredOverlay_Utils("❌ PDF export failed`n" err.Message, 2800, BANNER_ACCENT_ERROR)
    }
}

;-------------------------------------------------------------------
; PowerPoint Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe POWERPNT.EXE") && WinGetClass("A") != "#32770"

; Shift + P : Save as PDF via COM (same folder, or Desktop if unsaved)
+p:: PowerPoint_SaveAsPdf()

#HotIf