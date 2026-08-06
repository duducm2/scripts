; =============================================================================
; Shift keys module: hotif_powerpoint.ahk
; PowerPoint hotkeys
; Loaded via #include into the Shift keys.ahk process.
; =============================================================================

; ppFixedFormatTypePDF = 2; ppFixedFormatIntentScreen = 1; ppSaveAsPDF = 32

PowerPoint_GetActivePresentation() {
    try {
        pp := ComObjActive("PowerPoint.Application")
        return pp.ActivePresentation
    } catch {
    }
    return 0
}

PowerPoint_PdfDesktopPath(pres) {
    name := "Presentation"
    try {
        n := pres.Name
        if (n != "")
            name := n
    } catch {
    }
    name := RegExReplace(name, "i)\.(pptx|ppt|pptm|ppsx|pps)$", "")
    name := RegExReplace(name, '[<>:"/\\|?*]', "_")
    return A_Desktop "\" name ".pdf"
}

PowerPoint_TryExportPdf(pres, outPath) {
    try {
        pres.ExportAsFixedFormat(outPath, 2, 1)
        return { ok: true, err: "" }
    } catch Error as e1 {
        try {
            pres.ExportAsFixedFormat(outPath, 2)
            return { ok: true, err: "" }
        } catch Error as e2 {
            try {
                pres.SaveAs(outPath, 32)
                return { ok: true, err: "" }
            } catch Error as e3 {
                return { ok: false, err: e1.Message "`n" e2.Message "`n" e3.Message }
            }
        }
    }
}

PowerPoint_SaveAsPdf() {
    pres := PowerPoint_GetActivePresentation()
    if !pres {
        ShowCenteredOverlay_Utils("❌ PowerPoint COM unavailable", 2200, BANNER_ACCENT_ERROR)
        return
    }

    outPath := PowerPoint_PdfDesktopPath(pres)
    result := PowerPoint_TryExportPdf(pres, outPath)
    if !result.ok {
        ShowCenteredOverlay_Utils("❌ PDF export failed`n" outPath "`n" result.err, 3500, BANNER_ACCENT_ERROR)
        return
    }
    ShowCenteredOverlay_Utils("📄 Saved PDF`n" outPath, 2200, BANNER_ACCENT_SUCCESS)
}

;-------------------------------------------------------------------
; PowerPoint Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe POWERPNT.EXE") && WinGetClass("A") != "#32770"

; Shift + P : Save as PDF on Desktop
+p:: PowerPoint_SaveAsPdf()

#HotIf