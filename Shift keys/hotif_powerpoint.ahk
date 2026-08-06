; =============================================================================
; Shift keys module: hotif_powerpoint.ahk
; PowerPoint hotkeys
; Loaded via #include into the Shift keys.ahk process.
; =============================================================================
;
; Shift+P — Save as PDF on Desktop (working approach)
; ----------------------------------------------------
; No Save As UI. PowerPoint COM only.
;
; Flow:
;   1. ComObjActive("PowerPoint.Application").ActivePresentation
;   2. outPath := A_Desktop "\" <presentation Name without .pptx> ".pdf"
;      (A_Desktop = Windows Desktop; ignore pres.Path / SharePoint URLs)
;   3. pres.ExportAsFixedFormat(outPath, 2)   ; ppFixedFormatTypePDF
;   4. On throw → pres.SaveAs(outPath, 32)    ; ppSaveAsPDF fallback
;
; Why this shape:
;   - F12 / FileTypeControlHost type-ahead often picked BMP instead of PDF.
;   - Writing beside SharePoint/OneDrive paths (https://…) failed on work PCs.
;   - Desktop-only COM export is fast and keeps the .pptx open (ExportAsFixedFormat).
;
; Cheat sheet: cheatSheets["POWERPNT.EXE"] in cheat_sheet_registry.ahk
;
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

; Shift + P : Save as PDF on Desktop via COM
+p:: PowerPoint_SaveAsPdf()

#HotIf