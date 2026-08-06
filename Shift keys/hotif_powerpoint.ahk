; =============================================================================
; Shift keys module: hotif_powerpoint.ahk
; PowerPoint hotkeys
; Loaded via #include into the Shift keys.ahk process.
; =============================================================================
; Desktop PDF via COM (simple): ExportAsFixedFormat straight to A_Desktop.

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
    line := '{"sessionId":"a610fd","runId":"pre-fix","hypothesisId":"' hypothesisId '","location":"' location '","message":"' msg '","data":' dataJson ',"timestamp":' tick ',"now":"' ts '"}`n'
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

PowerPoint_ListDesktopPdfs(limit := 12) {
    names := ""
    count := 0
    try {
        loop files A_Desktop "\*.pdf", "F" {
            names .= (names = "" ? "" : " | ") A_LoopFileName
            count += 1
            if (count >= limit)
                break
        }
    } catch {
    }
    return names
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
    presName := ""
    try presPath := String(pres.Path)
    catch {
    }
    try presFull := String(pres.FullName)
    catch {
    }
    try presName := String(pres.Name)
    catch {
    }
    PowerPoint_DebugLog("H-D", "hotif_powerpoint.ahk:SaveAsPdf:entry", "COM Desktop PDF start", Map(
        "destPath", destPath,
        "A_Desktop", A_Desktop,
        "presName", presName,
        "presPath", presPath,
        "presFull", presFull,
        "scriptDir", A_ScriptDir
    ))
    ; #endregion

    try FileDelete(destPath)
    catch {
    }

    methodOk := ""
    methodErr := ""

    ; Method 1: ExportAsFixedFormat path, PDF, Screen
    try {
        pres.ExportAsFixedFormat(destPath, 2, 1)
        methodOk := "ExportAsFixedFormat(path,2,1)"
    } catch Error as e1 {
        methodErr .= "E1:" e1.Message " | "
        ; #region agent log
        PowerPoint_DebugLog("H-B", "hotif_powerpoint.ahk:m1", "ExportAsFixedFormat(2,1) threw", Map("err", e1.Message,
            "destPath", destPath))
        ; #endregion
        try {
            pres.ExportAsFixedFormat(destPath, 2)
            methodOk := "ExportAsFixedFormat(path,2)"
        } catch Error as e2 {
            methodErr .= "E2:" e2.Message " | "
            ; #region agent log
            PowerPoint_DebugLog("H-B", "hotif_powerpoint.ahk:m2", "ExportAsFixedFormat(2) threw", Map("err", e2.Message,
                "destPath", destPath))
            ; #endregion
            try {
                pres.SaveAs(destPath, 32)
                methodOk := "SaveAs(path,32)"
            } catch Error as e3 {
                methodErr .= "E3:" e3.Message
                ; #region agent log
                PowerPoint_DebugLog("H-A", "hotif_powerpoint.ahk:m3", "SaveAs(32) threw", Map("err", e3.Message,
                    "destPath", destPath))
                ; #endregion
                ShowCenteredOverlay_Utils("❌ PDF export failed`n" destPath "`n" methodErr, 3500, BANNER_ACCENT_ERROR)
                return
            }
        }
    }

    exists := PowerPoint_FileLooksValid(destPath)
    size := 0
    try size := FileGetSize(destPath)
    catch {
    }
    deskPdfs := PowerPoint_ListDesktopPdfs()

    ; #region agent log
    PowerPoint_DebugLog("H-A", "hotif_powerpoint.ahk:after", "COM returned without throw", Map(
        "methodOk", methodOk,
        "exists", exists,
        "size", size,
        "destPath", destPath,
        "desktopPdfs", deskPdfs
    ))
    PowerPoint_DebugLog("H-C", "hotif_powerpoint.ahk:scan", "Desktop PDF scan after COM", Map(
        "desktopPdfs", deskPdfs,
        "targetBase", baseName ".pdf"
    ))
    ; #endregion

    if !exists {
        ShowCenteredOverlay_Utils("❌ COM OK (" methodOk ") but no Desktop PDF`n" destPath "`nCheck Desktop\debug-a610fd.log",
            4500, BANNER_ACCENT_ERROR)
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