; =============================================================================
; Shift keys module: hotif_powerpoint.ahk
; PowerPoint hotkeys
; Loaded via #include into the Shift keys.ahk process.
; =============================================================================
; Shift+P: COM ExportAsFixedFormat → Windows Desktop (A_Desktop). No Save As UI.
; Work PPT/AHK often throws Type mismatch on Export; PowerShell COM fallback fixes that.
; Never treat SaveAs as success without verifying the Desktop file exists.

; ppFixedFormatTypePDF = 2; ppFixedFormatIntentScreen = 1

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
        n := String(pres.Name)
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

PowerPoint_FileLooksValid(path) {
    if !FileExist(path)
        return false
    try {
        return FileGetSize(path) > 0
    } catch {
        return false
    }
}

PowerPoint_TryExportAhk(pres, outPath) {
    try {
        pres.ExportAsFixedFormat(outPath, 2, 1)
        return { ok: true, err: "" }
    } catch Error as e1 {
        try {
            pres.ExportAsFixedFormat(outPath, 2)
            return { ok: true, err: "" }
        } catch Error as e2 {
            return { ok: false, err: e1.Message " | " e2.Message }
        }
    }
}

; Same ExportAsFixedFormat via PowerShell — correct COM enums on work Office builds.
PowerPoint_TryExportPowerShell(outPath) {
    outPs := StrReplace(outPath, "'", "''")
    ps := "$ErrorActionPreference = 'Stop'`n"
        . "$pp = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')`n"
        . "$pres = $pp.ActivePresentation`n"
        . "$dest = '" outPs "'`n"
        . "$pres.ExportAsFixedFormat($dest, 2, 1)`n"
        . "if (-not (Test-Path -LiteralPath $dest)) { throw ('Export returned but file missing: ' + $dest) }`n"
    ps1 := A_Temp "\ShiftKeysPptPdfExport.ps1"
    try FileDelete(ps1)
    catch {
    }
    try {
        FileAppend(ps, ps1, "UTF-8")
    } catch Error as eWrite {
        return { ok: false, err: "Could not write PS1: " eWrite.Message }
    }
    exitCode := 0
    try {
        RunWait('powershell.exe -NoProfile -ExecutionPolicy Bypass -File "' ps1 '"', , "Hide", &exitCode)
    } catch Error as eRun {
        return { ok: false, err: eRun.Message }
    }
    if (exitCode != 0)
        return { ok: false, err: "PowerShell exit " exitCode }
    if !PowerPoint_FileLooksValid(outPath)
        return { ok: false, err: "PowerShell finished but no Desktop PDF" }
    return { ok: true, err: "" }
}

PowerPoint_SaveAsPdf() {
    pres := PowerPoint_GetActivePresentation()
    if !pres {
        ShowCenteredOverlay_Utils("❌ PowerPoint COM unavailable", 2200, BANNER_ACCENT_ERROR)
        return
    }

    outPath := A_Desktop "\" PowerPoint_PdfBaseName(pres) ".pdf"
    try FileDelete(outPath)
    catch {
    }

    result := PowerPoint_TryExportAhk(pres, outPath)
    method := "AHK"
    if result.ok && !PowerPoint_FileLooksValid(outPath) {
        result := { ok: false, err: "AHK COM OK but file missing" }
    }
    if !result.ok {
        psResult := PowerPoint_TryExportPowerShell(outPath)
        if psResult.ok {
            result := psResult
            method := "PS"
        } else {
            ShowCenteredOverlay_Utils("❌ PDF export failed`n" outPath "`nAHK: " result.err "`nPS: " psResult.err, 4500,
                BANNER_ACCENT_ERROR)
            return
        }
    }

    if !PowerPoint_FileLooksValid(outPath) {
        ShowCenteredOverlay_Utils("❌ Export claimed OK (" method ") but no Desktop PDF`n" outPath, 4000,
            BANNER_ACCENT_ERROR)
        return
    }
    ShowCenteredOverlay_Utils("📄 Saved PDF`n" outPath, 1800, BANNER_ACCENT_SUCCESS)
}

;-------------------------------------------------------------------
; PowerPoint Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe POWERPNT.EXE") && WinGetClass("A") != "#32770"

; Shift + P : Save as PDF on Desktop via COM (no UI)
+p:: PowerPoint_SaveAsPdf()

#HotIf