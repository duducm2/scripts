# -*- coding: utf-8 -*-
from pathlib import Path

path = Path(
    r"C:\Users\eduev\Meu Drive\17 - Projects\scripts\Shift keys\hotif_powerpoint.ahk"
)

content = r"""`; =============================================================================
`; Shift keys module: hotif_powerpoint.ahk
`; PowerPoint hotkeys
`; Loaded via #include into the Shift keys.ahk process.
`; =============================================================================

`; ppFixedFormatTypePDF = 2; ppFixedFormatIntentScreen = 1; ppSaveAsPDF = 32

PowerPoint_GetActivePresentation() {
    try {
        pp := ComObjActive("PowerPoint.Application")
        return pp.ActivePresentation
    } catch {
    }
    return 0
}

PowerPoint_IsCloudOrNonLocalPath(path) {
    if (path = "")
        return true
    if RegExMatch(path, "i)^https?://")
        return true
    if InStr(path, "sharepoint.com", false) || InStr(path, ".sharepoint.", false)
        return true
    `; WebDAV-style SharePoint / OneDrive UNC often fails with ExportAsFixedFormat too.
    if RegExMatch(path, "i)^\\\\[^\\]+\.sharepoint\.com\\")
        return true
    return false
}

PowerPoint_NormalizeFolder(path) {
    if (path = "")
        return ""
    return RTrim(path, "\/")
}

`; Prefer a real local folder. SharePoint/OneDrive http(s) Path breaks COM export on work PCs.
PowerPoint_ResolveExportFolder(pres) {
    candidates := []

    try {
        p := PowerPoint_NormalizeFolder(pres.Path)
        if (p != "")
            candidates.Push(p)
    } catch {
    }

    try {
        full := pres.FullName
        if (full != "" && !PowerPoint_IsCloudOrNonLocalPath(full)) {
            SplitPath(full, , &dir)
            dir := PowerPoint_NormalizeFolder(dir)
            if (dir != "")
                candidates.Push(dir)
        }
    } catch {
    }

    for c in candidates {
        if PowerPoint_IsCloudOrNonLocalPath(c)
            continue
        if DirExist(c)
            return { folder: c, cloudFallback: false }
    }

    `; Unsaved, or cloud-only Path/FullName (typical work OneDrive/SharePoint open).
    return { folder: A_Desktop, cloudFallback: true }
}

PowerPoint_PdfOutputPath(pres, folder) {
    name := "Presentation"
    try {
        n := __PRES__.Name
        if (n != "")
            name := n
    } catch {
    }
    name := RegExReplace(name, "i)\.(pptx|ppt|pptm|ppsx|pps)$", "")
    `; Strip characters illegal in Windows file names (SharePoint titles can be odd).
    name := RegExReplace(name, '[<>:"/\\|?*]', "_")
    return folder "\" name ".pdf"
}

PowerPoint_TryExportPdf(pres, outPath) {
    `; Intent 1 = ppFixedFormatIntentScreen (some tenants reject bare 2-arg calls).
    try {
        __PRES__.ExportAsFixedFormat(outPath, 2, 1)
        return { ok: true, err: "" }
    } catch Error as e1 {
        try {
            __PRES__.ExportAsFixedFormat(outPath, 2)
            return { ok: true, err: "" }
        } catch Error as e2 {
            try {
                `; Last COM attempt: SaveAs PDF (may open PDF as active doc on some builds).
                __PRES__.SaveAs(outPath, 32)
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

    resolved := PowerPoint_ResolveExportFolder(pres)
    targetFolder := resolved.folder
    cloudFallback := resolved.cloudFallback

    outPath := PowerPoint_PdfOutputPath(pres, targetFolder)
    if (outPath = "" || targetFolder = "") {
        ShowCenteredOverlay_Utils("❌ Could not resolve PDF path", 2200, BANNER_ACCENT_ERROR)
        return
    }

    result := PowerPoint_TryExportPdf(pres, outPath)
    if result.ok {
        if cloudFallback
            ShowCenteredOverlay_Utils("📄 Cloud deck — PDF on Desktop`n" outPath, 2800, BANNER_ACCENT_SUCCESS)
        else
            ShowCenteredOverlay_Utils("📄 Saved PDF`n" outPath, 2200, BANNER_ACCENT_SUCCESS)
        return
    }

    `; Primary folder failed (permissions / sync): retry Desktop once if we hadn't already.
    if (targetFolder != A_Desktop) {
        deskPath := PowerPoint_PdfOutputPath(pres, A_Desktop)
        deskResult := PowerPoint_TryExportPdf(pres, deskPath)
        if deskResult.ok {
            ShowCenteredOverlay_Utils("📄 Folder blocked — PDF on Desktop`n" deskPath, 2800, BANNER_ACCENT_SUCCESS)
            return
        }
        result.err := result.err "`n" deskResult.err
        outPath := deskPath
    }

    ShowCenteredOverlay_Utils("❌ PDF export failed`n" outPath "`n" result.err, 3500, BANNER_ACCENT_ERROR)
}

`;-------------------------------------------------------------------
`; PowerPoint Shortcuts
`;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe POWERPNT.EXE") && WinGetClass("A") != "#32770"

`; Shift + P : Save as PDF via COM (local folder; Desktop if cloud/unsaved)
+p:: PowerPoint_SaveAsPdf()

#HotIf
"""

# Use placeholder to avoid write-tool corrupting "pres." member access in earlier attempts
content = content.replace("__PRES__", "pres")
# Turn escaped AHK comments back: we used `; so Python raw string keeps backtick — strip leading backticks on comments
lines = []
for line in content.splitlines(True):
    if line.startswith("`;"):
        line = ";" + line[2:]
    elif "`;" in line:
        line = line.replace("`;", ";")
    lines.append(line)
text = "".join(lines)

path.write_text(text, encoding="utf-8", newline="\n")
print(f"Wrote {path} ({len(text)} bytes)")
