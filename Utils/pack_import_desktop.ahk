; =============================================================================
; Utils module: pack_import_desktop.ahk
; Canonical Desktop pack paths for AI pack imports (overwrite variants on import)
; Agent docs: docs/prompt-data-output-and-finance-packs.md
; =============================================================================

PackImport_CanonicalDesktopPath(fileName) {
    return A_Desktop . "\" . fileName
}

; Move discovered Desktop pack to canonical name (overwrite). Returns canonical path or "".
PackImport_NormalizeDesktopSource(discoveredPath, canonicalFileName) {
    discoveredPath := Trim(discoveredPath)
    if (discoveredPath = "" || !FileExist(discoveredPath))
        return ""
    canonical := PackImport_CanonicalDesktopPath(canonicalFileName)
    SplitPath(discoveredPath, &discoveredName)
    if (StrLower(discoveredName) = StrLower(canonicalFileName))
        return discoveredPath
    try {
        FileMove(discoveredPath, canonical, 1)
        return canonical
    } catch {
        try {
            FileCopy(discoveredPath, canonical, 1)
            try FileDelete(discoveredPath)
            catch {
            }
            return canonical
        } catch {
            return discoveredPath
        }
    }
}
