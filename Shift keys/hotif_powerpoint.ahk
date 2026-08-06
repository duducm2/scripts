; =============================================================================
; Shift keys module: hotif_powerpoint.ahk
; PowerPoint hotkeys
; Loaded via #include into the Shift keys.ahk process.
; =============================================================================
; Desktop PDF via COM: ExportAsFixedFormat → A_Desktop.
; Evidence (debug-a610fd): plain ints → Type mismatch (0x80020005);
; SaveAs(...,32) returns OK on SharePoint decks but writes no Desktop file.

; ppFixedFormatTypePDF = 2; ppFixedFormatIntentScreen = 1
; msoFalse = 0; ppPrintAll = 1; ppPrintOutputSlides = 1

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

; #region agent log
PowerPoint_DebugLog(hypothesisId, location, message, dataMap := "") {
    logPath := A_ScriptDir "\..\debug-a610fd.log"
    deskLog := A_Desktop "\debug-a610fd.log"
    ts := A_Now
    tick := A_TickCount
    dataJson := "{}"
    if (dataMap != "") {
        parts := ""
        for k, v in dataMap {
            vs := RegExReplace(String(v), '\\', '\\\\')
            vs := RegExReplace(vs, '"', '\"')
            vs := RegExReplace(vs, "`r|`n", " ")
            parts .= (parts = "" ? "" : ",") '"' k '":"' vs '"'
        }
        dataJson := "{" parts "}"
    }
    msg := RegExReplace(String(message), '\\', '\\\\')
    msg := RegExReplace(msg, '"', '\"')
    line := '{"sessionId":"a610fd","runId":"post-fix","hypothesisId":"' hypothesisId '","location":"' location '","message":"' msg '","data":' dataJson ',"timestamp":' tick ',"now":"' ts '"}`n'
    try FileAppend(line, logPath)
    catch {
    }
    try FileAppend(line, deskLog)
    catch {
    }
}
; #endregion

PowerPoint_FileLooksValid(path) {
    if !FileExist(path)
        return false
    try {
        return FileGetSize(path) > 0
    } catch {
        return false
    }
}

PowerPoint_I4(n) {
    return ComValue(0x3, n) ; VT_I4 — plain AHK ints cause DISP_E_TYPEMISMATCH on ExportAsFixedFormat
}

PowerPoint_TryExportPdf(pres, outPath) {
    pdf := PowerPoint_I4(2)
    intent := PowerPoint_I4(1)
    msoFalse := PowerPoint_I4(0)
    printAll := PowerPoint_I4(1)
    slidesOut := PowerPoint_I4(1)

    ; Named minimal call with typed enums
    try {
        pres.ExportAsFixedFormat(outPath, pdf, intent)
        return { ok: true, method: "Export(path,pdf,intent) ComValue", err: "" }
    } catch Error as e1 {
        ; Fuller signature (FrameSlides / HandoutOrder / OutputType / PrintHidden / RangeType)
        try {
            pres.ExportAsFixedFormat(outPath, pdf, intent, msoFalse, printAll, slidesOut, msoFalse, , printAll)
            return { ok: true, method: "Export(full) ComValue", err: "" }
        } catch Error as e2 {
            try {
                pres.ExportAsFixedFormat(outPath, pdf)
                return { ok: true, method: "Export(path,pdf) ComValue", err: "" }
            } catch Error as e3 {
                return { ok: false, method: "", err: e1.Message " | " e2.Message " | " e3.Message }
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

    baseName := PowerPoint_PdfBaseName(pres)
    destPath := A_Desktop "\" baseName ".pdf"

    ; #region agent log
    presPath := ""
    presFull := ""
    try presPath := String(pres.Path)
    catch {
    }
    try presFull := String(pres.FullName)
    catch {
    }
    PowerPoint_DebugLog("H-fix", "hotif_powerpoint.ahk:entry", "Desktop PDF export start", Map(
        "destPath", destPath,
        "presPath", presPath,
        "presFull", presFull
    ))
    ; #endregion

    try FileDelete(destPath)
    catch {
    }

    result := PowerPoint_TryExportPdf(pres, destPath)

    exists := PowerPoint_FileLooksValid(destPath)
    size := 0
    try size := FileGetSize(destPath)
    catch {
    }

    ; #region agent log
    PowerPoint_DebugLog("H-fix", "hotif_powerpoint.ahk:after", "Export result", Map(
        "ok", result.ok,
        "method", result.method,
        "err", result.err,
        "exists", exists,
        "size", size,
        "destPath", destPath
    ))
    ; #endregion

    if !result.ok {
        ShowCenteredOverlay_Utils("❌ PDF export failed`n" destPath "`n" result.err, 4000, BANNER_ACCENT_ERROR)
        return
    }
    if !exists {
        ShowCenteredOverlay_Utils("❌ COM OK (" result.method ") but no Desktop PDF`n" destPath, 4000,
            BANNER_ACCENT_ERROR)
        return
    }
    ShowCenteredOverlay_Utils("📄 Saved PDF`n" destPath, 2200, BANNER_ACCENT_SUCCESS)
}

;-------------------------------------------------------------------
; PowerPoint Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe POWERPNT.EXE") && WinGetClass("A") != "#32770"

; Shift + P : Save as PDF on Desktop via COM
+p:: PowerPoint_SaveAsPdf()

#HotIf