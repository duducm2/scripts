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

PowerPoint_IsCloudOrNonLocalPath(path) {
    if (path = "")
        return true
    if RegExMatch(path, "i)^https?://")
        return true
    if InStr(path, "sharepoint.com", false) || InStr(path, ".sharepoint.", false)
        return true
    if RegExMatch(path, "i)^\\\\[^\\]+\.sharepoint\.com\\")
        return true
    return false
}

PowerPoint_NormalizeFolder(path) {
    if (path = "")
        return ""
    return RTrim(path, "\/")
}

PowerPoint_UrlDecode(s) {
    if (s = "")
        return ""
    s := StrReplace(s, "+", " ")
    out := ""
    i := 1
    len := StrLen(s)
    while (i <= len) {
        ch := SubStr(s, i, 1)
        if (ch = "%" && i + 2 <= len) {
            hex := SubStr(s, i + 1, 2)
            if RegExMatch(hex, "i)^[0-9A-F]{2}$") {
                out .= Chr(Integer("0x" hex))
                i += 3
                continue
            }
        }
        out .= ch
        i += 1
    }
    return out
}

; Strip Office sharing wrappers (/:p:/r/, query strings) so SyncEngines prefixes can match.
PowerPoint_NormalizeCloudUrl(url) {
    if (url = "")
        return ""
    url := RegExReplace(url, "\?.*$", "")
    url := PowerPoint_UrlDecode(url)
    if RegExMatch(url, "i)^(https?://[^/]+)/:[a-z]:/r/(.+)$", &m)
        url := m[1] "/" m[2]
    return PowerPoint_NormalizeFolder(url)
}

PowerPoint_PresentationFileName(pres) {
    name := ""
    try name := pres.Name
    catch {
    }
    if (name = "")
        return ""
    if !RegExMatch(name, "i)\.(pptx|ppt|pptm|ppsx|pps)$")
        name .= ".pptx"
    return name
}

; Returns array of { url, mount } from HKCU SyncEngines (AHK first, PowerShell fallback).
PowerPoint_GetOneDriveMounts() {
    mounts := []
    baseKey := "HKEY_CURRENT_USER\Software\SyncEngines\Providers\OneDrive"
    try {
        loop reg, baseKey, "K" {
            url := ""
            mount := ""
            try url := RegRead(baseKey "\" A_LoopRegName, "UrlNamespace")
            catch {
            }
            try mount := RegRead(baseKey "\" A_LoopRegName, "MountPoint")
            catch {
            }
            url := PowerPoint_NormalizeFolder(url)
            mount := PowerPoint_NormalizeFolder(StrReplace(mount, "\\", "\"))
            if (url != "" && mount != "" && DirExist(mount))
                mounts.Push({ url: url, mount: mount })
        }
    } catch {
    }

    if (mounts.Length = 0) {
        tmp := A_Temp "\ppt_onedrive_mounts.txt"
        psFile := A_Temp "\ppt_onedrive_mounts.ps1"
        try FileDelete(tmp)
        catch {
        }
        try FileDelete(psFile)
        catch {
        }
        script := (
            "$base = 'HKCU:\Software\SyncEngines\Providers\OneDrive'`n"
            "if (-not (Test-Path $base)) { exit 0 }`n"
            "Get-ChildItem $base -ErrorAction SilentlyContinue | ForEach-Object {`n"
            "  $p = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue`n"
            "  if ($p.UrlNamespace -and $p.MountPoint) {`n"
            "    $u = ([string]$p.UrlNamespace).TrimEnd('/')`n"
            "    $m = ([string]$p.MountPoint).TrimEnd('\','/')`n"
            "    Add-Content -LiteralPath $env:TEMP\ppt_onedrive_mounts.txt -Value ($u + [char]9 + $m)`n"
            "  }`n"
            "}`n"
        )
        try {
            FileAppend(script, psFile)
            RunWait('powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "' psFile '"', , "Hide")
            if FileExist(tmp) {
                for line in StrSplit(FileRead(tmp), "`n", "`r") {
                    line := Trim(line)
                    if (line = "" || !InStr(line, "`t"))
                        continue
                    parts := StrSplit(line, "`t")
                    if (parts.Length < 2)
                        continue
                    url := PowerPoint_NormalizeFolder(parts[1])
                    mount := PowerPoint_NormalizeFolder(parts[2])
                    if (url != "" && mount != "" && DirExist(mount))
                        mounts.Push({ url: url, mount: mount })
                }
            }
        } catch {
        }
        try FileDelete(tmp)
        catch {
        }
        try FileDelete(psFile)
        catch {
        }
    }
    return mounts
}

; Local filesystem roots to search for the .pptx (mounts + OneDrive env dirs).
PowerPoint_GetOneDriveSearchRoots() {
    roots := []
    seen := Map()
    addRoot(r) {
        r := PowerPoint_NormalizeFolder(r)
        if (r = "" || !DirExist(r))
            return
        key := StrLower(r)
        if seen.Has(key)
            return
        seen[key] := true
        roots.Push(r)
    }

    for m in PowerPoint_GetOneDriveMounts()
        addRoot(m.mount)

    for envName in ["OneDriveCommercial", "OneDriveConsumer", "OneDrive"] {
        try addRoot(EnvGet(envName))
        catch {
        }
    }

    ; A_UserProfile is not built-in on this AHK build — use USERPROFILE / Desktop parent.
    profile := ""
    try profile := EnvGet("USERPROFILE")
    catch {
    }
    if (profile = "" || !DirExist(profile)) {
        try SplitPath(A_Desktop, , &profile)
        catch {
            profile := ""
        }
    }
    if (profile != "" && DirExist(profile)) {
        try {
            loop files, profile "\OneDrive*", "D" {
                addRoot(A_LoopFileFullPath)
            }
        } catch {
        }
    }

    ; If Desktop itself lives under OneDrive (common at work), its parent is a sync root.
    try {
        deskParent := ""
        SplitPath(A_Desktop, , &deskParent)
        deskParent := PowerPoint_NormalizeFolder(deskParent)
        if InStr(deskParent, "OneDrive", false)
            addRoot(deskParent)
    } catch {
    }
    return roots
}

PowerPoint_CloudHintSegments(cloudHint) {
    segs := []
    if (cloudHint = "")
        return segs
    hint := PowerPoint_NormalizeCloudUrl(cloudHint)
    hint := RegExReplace(hint, "i)^https?://", "")
    for part in StrSplit(hint, ["/", "\"]) {
        part := Trim(PowerPoint_UrlDecode(part))
        if (part = "" || StrLen(part) < 3)
            continue
        if RegExMatch(part, "i)^(sites|teams|personal|documents|shared documents|forms|_layouts|r)$")
            continue
        if RegExMatch(part, "i)\.(pptx|ppt|pptm|ppsx|pps)$")
            continue
        segs.Push(part)
    }
    return segs
}

PowerPoint_PresentationBaseName(pres) {
    fileName := PowerPoint_PresentationFileName(pres)
    if (fileName = "")
        return ""
    return RegExReplace(fileName, "i)\.(pptx|ppt|pptm|ppsx|pps)$", "")
}

; Prefer the open deck when several same-named .pptx exist (e.g. SNB vs Piloto).
PowerPoint_ScoreLocalHit(fullPath, segments, baseName) {
    score := 0
    pathLower := StrLower(fullPath)
    parentDir := ""
    parentName := ""
    SplitPath(fullPath, , &parentDir)
    parentDir := PowerPoint_NormalizeFolder(parentDir)
    if (parentDir != "")
        SplitPath(parentDir, &parentName)

    ; Strong signal: ...\august-presentation-to-rafael\august-presentation-to-rafael.pptx
    if (baseName != "" && parentName != "" && (StrLower(parentName) = StrLower(baseName)))
        score += 1000

    ; Weight longer URL path segments higher (unique folder names beat generic ones).
    for seg in segments {
        if InStr(pathLower, StrLower(seg))
            score += 10 + StrLen(seg)
    }
    return score
}

; Find local folder of the deck by AHK recursive search (no PowerShell).
PowerPoint_FindLocalFolderByFileName(pres, cloudHint := "") {
    fileName := PowerPoint_PresentationFileName(pres)
    if (fileName = "")
        return ""
    baseName := PowerPoint_PresentationBaseName(pres)

    roots := PowerPoint_GetOneDriveSearchRoots()
    if (roots.Length = 0)
        return ""

    if (cloudHint = "") {
        try cloudHint := String(pres.FullName)
        catch {
            try cloudHint := String(pres.Path)
            catch {
            }
        }
    }
    segments := PowerPoint_CloudHintSegments(cloudHint)

    ; Pass 1: require parent folder name == presentation base name
    ; (…\august-presentation-to-rafael\august-presentation-to-rafael.pptx)
    bestPath := PowerPoint_PickBestPptxHit(roots, fileName, segments, baseName, true)
    if (bestPath = "")
        bestPath := PowerPoint_PickBestPptxHit(roots, fileName, segments, baseName, false)

    if (bestPath = "" || !FileExist(bestPath))
        return ""
    SplitPath(bestPath, , &dir)
    dir := PowerPoint_NormalizeFolder(dir)
    if (dir != "" && DirExist(dir))
        return dir
    return ""
}

PowerPoint_PickBestPptxHit(roots, fileName, segments, baseName, requireParentMatch) {
    bestPath := ""
    bestScore := -1
    bestTime := 0
    hitCount := 0
    for root in roots {
        try {
            loop files, root "\" fileName, "FR" {
                hitCount += 1
                full := A_LoopFileFullPath
                parentDir := ""
                parentName := ""
                SplitPath(full, , &parentDir)
                parentDir := PowerPoint_NormalizeFolder(parentDir)
                if (parentDir != "")
                    SplitPath(parentDir, &parentName)
                if (requireParentMatch) {
                    if (baseName = "" || parentName = "" || (StrLower(parentName) != StrLower(baseName))) {
                        if (hitCount >= 80)
                            break 2
                        continue
                    }
                }
                score := PowerPoint_ScoreLocalHit(full, segments, baseName)
                mtime := 0
                try mtime := FileGetTime(full, "M")
                catch {
                }
                if ((score > bestScore) || (score = bestScore && mtime > bestTime)) {
                    bestScore := score
                    bestTime := mtime
                    bestPath := full
                }
                if (hitCount >= 80)
                    break 2
            }
        } catch {
        }
    }
    return bestPath
}

; Map https SharePoint/OneDrive file/folder URL to a local directory that exists.
PowerPoint_CloudUrlToLocalFolder(cloudUrl) {
    if (cloudUrl = "" || !PowerPoint_IsCloudOrNonLocalPath(cloudUrl))
        return ""

    url := PowerPoint_NormalizeCloudUrl(cloudUrl)
    if RegExMatch(url, "i)\.(pptx|ppt|pptm|ppsx|pps)(/)?$") {
        url := RegExReplace(url, "i)/[^/]+\.(pptx|ppt|pptm|ppsx|pps)$", "")
        url := PowerPoint_NormalizeFolder(url)
    }

    mounts := PowerPoint_GetOneDriveMounts()
    if (mounts.Length = 0)
        return ""

    bestUrl := ""
    bestMount := ""
    urlLower := StrLower(url)
    for m in mounts {
        ns := PowerPoint_NormalizeFolder(m.url)
        if (ns = "")
            continue
        nsLower := StrLower(ns)
        ; Prefix match OR namespace contained in URL (handles CID / path quirks).
        matched := (InStr(urlLower, nsLower) = 1) || InStr(urlLower, nsLower)
        if matched && (StrLen(ns) > StrLen(bestUrl)) {
            bestUrl := ns
            bestMount := m.mount
        }
    }
    if (bestMount = "")
        return ""

    ; Prefer remainder after namespace when it was a true prefix; else use mount root and let
    ; the filename search refine further.
    localFolder := bestMount
    if (InStr(urlLower, StrLower(bestUrl)) = 1) {
        rest := SubStr(url, StrLen(bestUrl) + 1)
        rest := RegExReplace(rest, "^/+", "")
        rest := PowerPoint_UrlDecode(rest)
        rest := StrReplace(rest, "/", "\")
        rest := PowerPoint_NormalizeFolder(rest)
        if (rest != "")
            localFolder := bestMount "\" rest
    }

    probe := PowerPoint_NormalizeFolder(localFolder)
    while (probe != "" && probe != bestMount) {
        if DirExist(probe)
            return probe
        SplitPath(probe, , &parent)
        parent := PowerPoint_NormalizeFolder(parent)
        if (parent = "" || parent = probe)
            break
        probe := parent
    }
    if DirExist(bestMount)
        return bestMount
    return ""
}

PowerPoint_ResolveExportFolder(pres) {
    candidates := []
    cloudUrls := []

    try {
        p := PowerPoint_NormalizeFolder(pres.Path)
        if (p != "") {
            candidates.Push(p)
            if PowerPoint_IsCloudOrNonLocalPath(p)
                cloudUrls.Push(p)
        }
    } catch {
    }

    try {
        full := pres.FullName
        if (full != "") {
            if !PowerPoint_IsCloudOrNonLocalPath(full) {
                SplitPath(full, , &dir)
                dir := PowerPoint_NormalizeFolder(dir)
                if (dir != "")
                    candidates.Push(dir)
            } else {
                cloudUrls.Push(full)
            }
        }
    } catch {
    }

    for c in candidates {
        if PowerPoint_IsCloudOrNonLocalPath(c)
            continue
        if DirExist(c)
            return { folder: c, syncFolder: c, cloudFallback: false }
    }

    cloudHint := (cloudUrls.Length > 0) ? cloudUrls[1] : ""

    for cu in cloudUrls {
        syncDir := PowerPoint_CloudUrlToLocalFolder(cu)
        if (syncDir != "" && DirExist(syncDir)) {
            ; Refine with filename search when possible (exact deck folder).
            byName := PowerPoint_FindLocalFolderByFileName(pres, cu)
            if (byName != "" && DirExist(byName))
                return { folder: byName, syncFolder: byName, cloudFallback: false }
            return { folder: syncDir, syncFolder: syncDir, cloudFallback: false }
        }
    }

    ; Primary fix for work PCs: locate the local .pptx under OneDrive sync roots.
    byName := PowerPoint_FindLocalFolderByFileName(pres, cloudHint)
    if (byName != "" && DirExist(byName))
        return { folder: byName, syncFolder: byName, cloudFallback: false }

    return { folder: A_Desktop, syncFolder: "", cloudFallback: true }
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
    name := RegExReplace(name, '[<>:"/\\|?*]', "_")
    return folder "\" name ".pdf"
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

PowerPoint_MovePdfToFolder(srcPdf, destFolder) {
    if (srcPdf = "" || destFolder = "" || !FileExist(srcPdf) || !DirExist(destFolder))
        return ""
    SplitPath(srcPdf, &fileName)
    destPdf := destFolder "\" fileName
    if (PowerPoint_NormalizeFolder(srcPdf) = PowerPoint_NormalizeFolder(destPdf))
        return destPdf
    try {
        if FileExist(destPdf)
            FileDelete(destPdf)
    } catch {
    }
    try {
        FileMove(srcPdf, destPdf, 1)
        if FileExist(destPdf) {
            ; OneDrive sometimes leaves a ghost/source copy — force-clear source.
            if FileExist(srcPdf) {
                try FileDelete(srcPdf)
                catch {
                }
            }
            return destPdf
        }
    } catch {
    }
    try {
        FileCopy(srcPdf, destPdf, 1)
        if FileExist(destPdf) {
            try FileDelete(srcPdf)
            catch {
            }
            return destPdf
        }
    } catch {
    }
    return ""
}

; Delete same-named Desktop leftover after a successful sync-folder save.
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

PowerPoint_SaveAsPdf() {
    pres := PowerPoint_GetActivePresentation()
    if !pres {
        ShowCenteredOverlay_Utils("❌ PowerPoint COM unavailable", 2200, BANNER_ACCENT_ERROR)
        return
    }

    resolved := PowerPoint_ResolveExportFolder(pres)
    targetFolder := resolved.folder
    syncFolder := resolved.syncFolder
    cloudFallback := resolved.cloudFallback

    outPath := PowerPoint_PdfOutputPath(pres, targetFolder)
    if (outPath = "" || targetFolder = "") {
        ShowCenteredOverlay_Utils("❌ Could not resolve PDF path", 2200, BANNER_ACCENT_ERROR)
        return
    }

    deskNorm := PowerPoint_NormalizeFolder(A_Desktop)
    result := PowerPoint_TryExportPdf(pres, outPath)

    ; If we already wrote next to the deck (not Desktop), clear any Desktop leftover and done.
    if (result.ok && (PowerPoint_NormalizeFolder(targetFolder) != deskNorm)) {
        PowerPoint_ClearDesktopLeftoverPdf(pres, outPath)
        ShowCenteredOverlay_Utils("📄 Saved PDF`n" outPath, 2200, BANNER_ACCENT_SUCCESS)
        return
    }

    ; Ensure a Desktop-staged PDF exists, then move it into the sync folder.
    deskPath := PowerPoint_PdfOutputPath(pres, A_Desktop)
    if (result.ok && (PowerPoint_NormalizeFolder(targetFolder) = deskNorm)) {
        deskPath := outPath
    } else {
        deskResult := PowerPoint_TryExportPdf(pres, deskPath)
        if !deskResult.ok {
            ShowCenteredOverlay_Utils("❌ PDF export failed`n" deskPath "`n" result.err "`n" deskResult.err, 3500,
                BANNER_ACCENT_ERROR)
            return
        }
    }

    moveTarget := syncFolder
    if (moveTarget = "" || (PowerPoint_NormalizeFolder(moveTarget) = deskNorm)) {
        hint := ""
        try hint := String(pres.FullName)
        catch {
            try hint := String(pres.Path)
            catch {
            }
        }
        moveTarget := PowerPoint_FindLocalFolderByFileName(pres, hint)
    }
    if (PowerPoint_NormalizeFolder(moveTarget) = deskNorm)
        moveTarget := ""

    if (moveTarget != "" && DirExist(moveTarget)) {
        moved := PowerPoint_MovePdfToFolder(deskPath, moveTarget)
        if (moved != "") {
            PowerPoint_ClearDesktopLeftoverPdf(pres, moved)
            ShowCenteredOverlay_Utils("📄 Moved PDF to sync folder`n" moved, 2800, BANNER_ACCENT_SUCCESS)
            return
        }
        ShowCenteredOverlay_Utils("📄 PDF on Desktop (move failed)`n" deskPath "`n→ " moveTarget, 3200,
            BANNER_ACCENT_INTERMEDIATE)
        return
    }

    ShowCenteredOverlay_Utils("📄 Left on Desktop — could not resolve sync folder`n" deskPath, 3200,
        BANNER_ACCENT_INTERMEDIATE)
}

;-------------------------------------------------------------------
; PowerPoint Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe POWERPNT.EXE") && WinGetClass("A") != "#32770"

; Shift + P : Save as PDF via COM (sync folder when possible; else Desktop then move)
+p:: PowerPoint_SaveAsPdf()

#HotIf