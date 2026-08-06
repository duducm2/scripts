# -*- coding: utf-8 -*-
from pathlib import Path

p = Path(
    r"C:\Users\eduev\Meu Drive\17 - Projects\scripts\Shift keys\hotif_powerpoint.ahk"
)
text = p.read_text(encoding="utf-8")
start = text.find("PowerPoint_TryExportPdf")
if start < 0:
    raise SystemExit("marker not found")
head = text[:start]

tail = r"""
PowerPoint_TryExportPdf(pres, outPath) {
    try {
        XXX.ExportAsFixedFormat(outPath, 2, 1)
        return { ok: true, err: "" }
    } catch Error as e1 {
        try {
            XXX.ExportAsFixedFormat(outPath, 2)
            return { ok: true, err: "" }
        } catch Error as e2 {
            try {
                XXX.SaveAs(outPath, 32)
                return { ok: true, err: "" }
            } catch Error as e3 {
                return { ok: false, err: e1.Message "`n" e2.Message "`n" e3.Message }
            }
        }
    }
}

; Local non-OneDrive temp folder (work Desktop is often under OneDrive too).
PowerPoint_StagingDir() {
    dir := A_Temp "\ShiftKeysPptPdf"
    try {
        if !DirExist(dir)
            DirCreate(dir)
    } catch {
    }
    return dir
}

PowerPoint_VerifyPdfFile(path, minBytes := 256) {
    if (path = "" || !FileExist(path))
        return false
    try {
        return FileGetSize(path) >= minBytes
    } catch {
        return false
    }
}

PowerPoint_WaitForPdfFile(path, timeoutMs := 8000) {
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        if PowerPoint_VerifyPdfFile(path)
            return true
        Sleep 200
    }
    return PowerPoint_VerifyPdfFile(path)
}

; Copy verified file into dest folder; only succeed if dest exists with real size.
PowerPoint_CopyPdfToFolder(srcPdf, destFolder) {
    if (srcPdf = "" || destFolder = "" || !PowerPoint_VerifyPdfFile(srcPdf) || !DirExist(destFolder))
        return ""
    SplitPath(srcPdf, &fileName)
    destPdf := destFolder "\" fileName
    if (PowerPoint_NormalizeFolder(srcPdf) = PowerPoint_NormalizeFolder(destPdf))
        return PowerPoint_VerifyPdfFile(destPdf) ? destPdf : ""

    try {
        if FileExist(destPdf)
            FileDelete(destPdf)
    } catch {
    }

    try FileCopy(srcPdf, destPdf, 1)
    catch {
    }
    if PowerPoint_WaitForPdfFile(destPdf)
        return destPdf
    return ""
}

PowerPoint_ClearDesktopLeftoverPdf(pres, keptPath := "") {
    deskPdf := PowerPoint_PdfOutputPath(pres, A_Desktop)
    if (deskPdf = "" || !FileExist(deskPdf))
        return
    if (keptPath != "" && (PowerPoint_NormalizeFolder(deskPdf) = PowerPoint_NormalizeFolder(keptPath)))
        return
    try FileDelete(deskPdf)
    catch {
        Sleep 200
        try FileDelete(deskPdf)
        catch {
        }
    }
}

; Always export to local TEMP first, verify size, then copy beside the presentation.
PowerPoint_SaveAsPdf() {
    XXX := PowerPoint_GetActivePresentation()
    if !XXX {
        ShowCenteredOverlay_Utils("❌ PowerPoint COM unavailable", 2200, BANNER_ACCENT_ERROR)
        return
    }

    resolved := PowerPoint_ResolveExportFolder(XXX)
    targetFolder := resolved.folder
    syncFolder := resolved.syncFolder
    deskNorm := PowerPoint_NormalizeFolder(A_Desktop)

    if ((targetFolder = "" || PowerPoint_NormalizeFolder(targetFolder) = deskNorm) && syncFolder != "")
        targetFolder := syncFolder
    if (targetFolder = "" || PowerPoint_NormalizeFolder(targetFolder) = deskNorm) {
        hint := ""
        try hint := String(XXX.FullName)
        catch {
            try hint := String(XXX.Path)
            catch {
            }
        }
        byName := PowerPoint_FindLocalFolderByFileName(XXX, hint)
        if (byName != "" && DirExist(byName))
            targetFolder := byName
    }
    if (targetFolder = "" || !DirExist(targetFolder))
        targetFolder := A_Desktop

    stageDir := PowerPoint_StagingDir()
    stagePath := PowerPoint_PdfOutputPath(XXX, stageDir)
    try {
        if FileExist(stagePath)
            FileDelete(stagePath)
    } catch {
    }

    result := PowerPoint_TryExportPdf(XXX, stagePath)
    if !result.ok || !PowerPoint_WaitForPdfFile(stagePath) {
        ShowCenteredOverlay_Utils("❌ PDF export failed in temp`n" stagePath "`n" result.err, 3500, BANNER_ACCENT_ERROR)
        return
    }

    if (PowerPoint_NormalizeFolder(targetFolder) = PowerPoint_NormalizeFolder(stageDir)) {
        ShowCenteredOverlay_Utils("📄 Saved PDF (temp)`n" stagePath, 2800, BANNER_ACCENT_INTERMEDIATE)
        return
    }

    placed := PowerPoint_CopyPdfToFolder(stagePath, targetFolder)
    if (placed = "") {
        ShowCenteredOverlay_Utils("❌ Could not place PDF in deck folder`n" targetFolder "`nKept here:`n" stagePath, 4500, BANNER_ACCENT_ERROR)
        return
    }

    if !PowerPoint_WaitForPdfFile(placed) {
        ShowCenteredOverlay_Utils("❌ PDF vanished after copy`n" placed "`nKept temp:`n" stagePath, 4500, BANNER_ACCENT_ERROR)
        return
    }

    try FileDelete(stagePath)
    catch {
    }
    PowerPoint_ClearDesktopLeftoverPdf(XXX, placed)
    ShowCenteredOverlay_Utils("📄 Saved PDF`n" placed, 2800, BANNER_ACCENT_SUCCESS)
}

;-------------------------------------------------------------------
; PowerPoint Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe POWERPNT.EXE") && WinGetClass("A") != "#32770"

; Shift + P : Save as PDF (temp verify, then copy beside deck)
+p:: PowerPoint_SaveAsPdf()

#HotIf
""".lstrip(
    "\n"
)

tail = tail.replace("XXX", "pres")
p.write_text(head + tail, encoding="utf-8", newline="\n")
print("wrote", p, p.stat().st_size)
