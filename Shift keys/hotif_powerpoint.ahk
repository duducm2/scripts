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
; Shift+C — Center selected shape(s) on slide
; ---------------------------------------------
; COM ShapeRange.Align: msoAlignCenters (1) + msoAlignMiddles (4),
; RelativeTo = msoTrue (-1) → align to slide edges.
;
; Shape pack (L/R/T/B/H/V/D/Y/G/U/F/K/E/W/Q)
; ----------------------------------------------
; All COM-only via ShapeRange.Align / Distribute / ZOrder / Group.
; RelativeTo = msoTrue (-1) = relative to slide edges.
;
; msoAlignLefts=0  Centers=1  Rights=2  Tops=3  Middles=4  Bottoms=5
; msoDistributeHorizontally=0  Vertically=1
; msoBringToFront=0  msoSendToBack=1  msoBringForward=2  msoSendBackward=3
; msoTrue=-1
;
; =============================================================================

; ppFixedFormatTypePDF = 2; ppSaveAsPDF = 32

;-------------------------------------------------------------------
; Shared COM helpers
;-------------------------------------------------------------------

PowerPoint_GetActivePresentation() {
    try {
        pp := ComObjActive("PowerPoint.Application")
        return pp.ActivePresentation
    } catch {
    }
    return 0
}

PowerPoint_GetShapeRange() {
    try {
        pp := ComObjActive("PowerPoint.Application")
        return pp.ActiveWindow.Selection.ShapeRange
    } catch {
    }
    return 0
}

PowerPoint_Align(cmd) {
    sr := PowerPoint_GetShapeRange()
    if !sr {
        ShowCenteredOverlay_Utils("❌ Select a shape first", 1800, BANNER_ACCENT_ERROR)
        return
    }
    try {
        sr.Align(cmd, -1)
    } catch Error as e {
        ShowCenteredOverlay_Utils("❌ Align failed`n" e.Message, 2200, BANNER_ACCENT_ERROR)
    }
}

PowerPoint_Distribute(cmd) {
    sr := PowerPoint_GetShapeRange()
    if !sr {
        ShowCenteredOverlay_Utils("❌ Select shapes first", 1800, BANNER_ACCENT_ERROR)
        return
    }
    try {
        sr.Distribute(cmd, -1)
    } catch Error as e {
        ShowCenteredOverlay_Utils("❌ Distribute failed`n" e.Message, 2200, BANNER_ACCENT_ERROR)
    }
}

PowerPoint_ZOrder(cmd) {
    sr := PowerPoint_GetShapeRange()
    if !sr {
        ShowCenteredOverlay_Utils("❌ Select a shape first", 1800, BANNER_ACCENT_ERROR)
        return
    }
    try {
        sr.ZOrder(cmd)
    } catch Error as e {
        ShowCenteredOverlay_Utils("❌ Z-order failed`n" e.Message, 2200, BANNER_ACCENT_ERROR)
    }
}

PowerPoint_Group() {
    sr := PowerPoint_GetShapeRange()
    if !sr {
        ShowCenteredOverlay_Utils("❌ Select shapes first", 1800, BANNER_ACCENT_ERROR)
        return
    }
    try {
        sr.Group()
    } catch Error as e {
        ShowCenteredOverlay_Utils("❌ Group failed`n" e.Message, 2200, BANNER_ACCENT_ERROR)
    }
}

PowerPoint_Ungroup() {
    sr := PowerPoint_GetShapeRange()
    if !sr {
        ShowCenteredOverlay_Utils("❌ Select a group first", 1800, BANNER_ACCENT_ERROR)
        return
    }
    try {
        sr.Ungroup()
    } catch Error as e {
        ShowCenteredOverlay_Utils("❌ Ungroup failed`n" e.Message, 2200, BANNER_ACCENT_ERROR)
    }
}

PowerPoint_Duplicate() {
    sr := PowerPoint_GetShapeRange()
    if !sr {
        ShowCenteredOverlay_Utils("❌ Select a shape first", 1800, BANNER_ACCENT_ERROR)
        return
    }
    try {
        sr.Duplicate()
    } catch Error as e {
        ShowCenteredOverlay_Utils("❌ Duplicate failed`n" e.Message, 2200, BANNER_ACCENT_ERROR)
    }
}

;-------------------------------------------------------------------
; PDF export
;-------------------------------------------------------------------

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
        pres.ExportAsFixedFormat(outPath, 2)
        ShowCenteredOverlay_Utils("📄 Saved PDF`n" outPath, 2200, BANNER_ACCENT_SUCCESS)
        return
    } catch {
    }

    try {
        pres.SaveAs(outPath, 32)
        ShowCenteredOverlay_Utils("📄 Saved PDF`n" outPath, 2200, BANNER_ACCENT_SUCCESS)
        return
    } catch Error as err {
        ShowCenteredOverlay_Utils("❌ PDF export failed`n" err.Message, 2800, BANNER_ACCENT_ERROR)
    }
}

;-------------------------------------------------------------------
; Center on slide (Shift+C)
;-------------------------------------------------------------------

PowerPoint_CenterOnSlide() {
    sr := PowerPoint_GetShapeRange()
    if !sr {
        ShowCenteredOverlay_Utils("❌ Center failed — select a shape first", 2200, BANNER_ACCENT_ERROR)
        return
    }
    try {
        sr.Align(1, -1)
        sr.Align(4, -1)
    } catch Error as e {
        ShowCenteredOverlay_Utils("❌ Center failed`n" e.Message, 2500, BANNER_ACCENT_ERROR)
    }
}

;-------------------------------------------------------------------
; PowerPoint Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe POWERPNT.EXE") && WinGetClass("A") != "#32770"

; --- PDF ---
+p:: PowerPoint_SaveAsPdf()

; --- Align (relative to slide) ---
+c:: PowerPoint_CenterOnSlide()          ; Center (H+V)
+l:: PowerPoint_Align(0)                 ; Left
+r:: PowerPoint_Align(2)                 ; Right
+t:: PowerPoint_Align(3)                 ; Top
+b:: PowerPoint_Align(5)                 ; Bottom
+h:: PowerPoint_Align(1)                 ; Horizontal center only
+v:: PowerPoint_Align(4)                 ; Vertical middle only

; --- Distribute (relative to slide) ---
+d:: PowerPoint_Distribute(0)            ; Distribute Horizontally
+y:: PowerPoint_Distribute(1)            ; Distribute Vertically

; --- Group / Ungroup ---
+g:: PowerPoint_Group()                  ; Group
+u:: PowerPoint_Ungroup()                ; Ungroup

; --- Z-order ---
+f:: PowerPoint_ZOrder(0)                ; bring to Front
+k:: PowerPoint_ZOrder(1)                ; send to bacK
+e:: PowerPoint_ZOrder(2)                ; bring forward (onE step)
+w:: PowerPoint_ZOrder(3)                ; send backWard

; --- Duplicate ---
+q:: PowerPoint_Duplicate()              ; clo-Q-ne / duplicate

#HotIf