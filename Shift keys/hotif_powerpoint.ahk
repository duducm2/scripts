; =============================================================================
; Shift keys module: hotif_powerpoint.ahk
; PowerPoint hotkeys
; Loaded via #include into the Shift keys.ahk process.
; =============================================================================
; Desktop PDF via COM: export to local TEMP (reliable), verify, then copy to Desktop.
; Direct Desktop ExportAsFixedFormat often returns OK with no file on OneDrive/work PCs.

; ppFixedFormatTypePDF = 2; ppFixedFormatIntentScreen = 1; ppSaveAsPDF = 32

PowerPoint_GetActivePresentation() {
    try {
        pp := ComObjActive("PowerPoint.Application")
        return pp.ActivePresentation
    } catch {
    }
    return 0
}

PowerPoint_PdfBaseName(pres) {
    name := "Presentation"
    try {
        n := pres.Name
        if (n != "")
            name := n
    } catch {
    }
    name := RegExReplace(name, "i)\.(pptx|ppt|pptm|ppsx|pps)$", "")
    name := RegExReplace(name, '[<>:"/\\|?*]', "_")
    if (name = "")
        name := "Presentation"
    return name
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

PowerPoint_FileLooksValid(path) {
    if !FileExist(path)
        return false
    try {
        return FileGetSize(path) > 0
    } catch {
        return false
    }
}

PowerPoint_SaveAsPdf() {
    pres := PowerPoint_GetActivePresentation()
    if !pres {
        ShowCenteredOverlay_Utils("❌ PowerPoint COM unavailable", 2200, BANNER_ACCENT_ERROR)
        return
    }

    baseName := PowerPoint_PdfBaseName(pres)
    destPath := A_Desktop "\" baseName ".pdf"
    stageDir := A_Temp "\ShiftKeysPptPdf"
    stagePath := stageDir "\" baseName ".pdf"

    try DirCreate(stageDir)
    catch {
    }
    try FileDelete(stagePath)
    catch {
    }

    result := PowerPoint_TryExportPdf(pres, stagePath)
    if !result.ok {
        ShowCenteredOverlay_Utils("❌ PDF export failed`n" stagePath "`n" result.err, 3500, BANNER_ACCENT_ERROR)
        return
    }
    if !PowerPoint_FileLooksValid(stagePath) {
        ShowCenteredOverlay_Utils("❌ COM said OK but no PDF in TEMP`n" stagePath "`nDesktop target:`n" destPath, 4000,
            BANNER_ACCENT_ERROR)
        return
    }

    try FileDelete(destPath)
    catch {
    }
    try {
        FileCopy(stagePath, destPath, true)
    } catch Error as eCopy {
        ShowCenteredOverlay_Utils("❌ PDF ready in TEMP but Desktop copy failed`n" destPath "`n" eCopy.Message, 4000,
            BANNER_ACCENT_ERROR)
        return
    }
    if !PowerPoint_FileLooksValid(destPath) {
        ShowCenteredOverlay_Utils("❌ TEMP PDF ok, but missing on Desktop`n" destPath "`nTEMP:`n" stagePath, 4000,
            BANNER_ACCENT_ERROR)
        return
    }

    try FileDelete(stagePath)
    catch {
    }

    ShowCenteredOverlay_Utils("📄 Saved PDF`n" destPath, 2200, BANNER_ACCENT_SUCCESS)
}

;-------------------------------------------------------------------
; PowerPoint Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe POWERPNT.EXE") && WinGetClass("A") != "#32770"

; Shift + P : Save as PDF on Desktop (COM → TEMP → Desktop)
+p:: PowerPoint_SaveAsPdf()

#HotIf