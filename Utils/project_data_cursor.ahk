; =============================================================================
; Utils module: project_data_cursor.ahk
; Shared project registry for WindowManagement / Utils / Shift keys.
; Persistent store: assets/data/projects.ini
; =============================================================================

; Character sequence for assignment: 1 2 3 4 5 q w e r t a s d f g z x c v b 6 7 8 9 0 y u i o p h j k l n m , .
global g_ProjectCharSequence := ["1", "2", "3", "4", "5", "q", "w", "e", "r", "t", "a", "s", "d", "f", "g", "z", "x",
    "c", "v", "b", "6", "7", "8", "9", "0", "y", "u", "i", "o", "p", "h", "j", "k", "l", "n", "m", ",", "."]

global g_Projects := []
global g_ProjectDataCacheReady := false
global g_ProjectDataCacheMtime := ""

ProjectData_IniPath() {
    return A_ScriptDir "\assets\data\projects.ini"
}

ProjectData_IsValidChar(char) {
    global g_ProjectCharSequence
    if (char = "" || !IsObject(g_ProjectCharSequence))
        return false
    for c in g_ProjectCharSequence {
        if (c = char)
            return true
    }
    return false
}

ProjectData_NormalizeIniValue(val) {
    if (val = "" || val = "ERROR")
        return ""
    return val
}

ProjectData_Invalidate() {
    global g_Projects, g_ProjectDataCacheReady, g_ProjectDataCacheMtime
    g_Projects := []
    g_ProjectDataCacheReady := false
    g_ProjectDataCacheMtime := ""
}

ProjectData_FileMtime() {
    path := ProjectData_IniPath()
    if (!FileExist(path))
        return ""
    mtime := ""
    try mtime := FileGetTime(path, "M")
    catch {
        mtime := ""
    }
    return mtime
}

; Load projects from INI. Reloads when the file mtime changes (or force=true).
; skipMtime: return in-memory cache without FileGetTime (hotkey path on Google Drive).
ProjectData_Load(force := false, skipMtime := false) {
    global g_Projects, g_ProjectDataCacheReady, g_ProjectDataCacheMtime
    if (!force && skipMtime && g_ProjectDataCacheReady)
        return g_Projects
    mtime := ProjectData_FileMtime()
    if (!force && g_ProjectDataCacheReady && mtime = g_ProjectDataCacheMtime)
        return g_Projects

    list := []
    taken := Map()
    path := ProjectData_IniPath()
    if (FileExist(path)) {
        idx := 1
        loop 200 {
            section := "Project_" . idx
            name := ""
            try name := IniRead(path, section, "Name", "")
            catch {
                break
            }
            if (name = "ERROR")
                break
            charVal := ""
            pathVal := ""
            workPath := ""
            try charVal := IniRead(path, section, "Char", "")
            try pathVal := IniRead(path, section, "Path", "")
            try workPath := IniRead(path, section, "WorkPath", "")
            name := ProjectData_NormalizeIniValue(name)
            charVal := StrLower(ProjectData_NormalizeIniValue(charVal))
            pathVal := ProjectData_NormalizeIniValue(pathVal)
            workPath := ProjectData_NormalizeIniValue(workPath)
            if (name = "" && pathVal = "" && workPath = "") {
                idx += 1
                continue
            }
            if (!ProjectData_IsValidChar(charVal) || taken.Has(charVal))
                charVal := ""
            if (charVal != "")
                taken[charVal] := true
            list.Push({ name: name, char: charVal, path: pathVal, workPath: workPath })
            idx += 1
        }
    }

    g_Projects := list
    g_ProjectDataCacheReady := true
    g_ProjectDataCacheMtime := mtime
    return g_Projects
}

; Rewrite INI as contiguous [Project_1]…[Project_N] and refresh the in-process cache.
ProjectData_Save(list) {
    global g_Projects, g_ProjectDataCacheReady, g_ProjectDataCacheMtime
    path := ProjectData_IniPath()
    try DirCreate(A_ScriptDir "\assets\data")
    catch {
    }
    try FileDelete(path)
    catch {
    }
    if (!IsObject(list) || list.Length = 0) {
        g_Projects := []
        g_ProjectDataCacheReady := true
        g_ProjectDataCacheMtime := ProjectData_FileMtime()
        return true
    }
    try {
        idx := 1
        for project in list {
            section := "Project_" . idx
            IniWrite(project.HasProp("name") ? project.name : "", path, section, "Name")
            IniWrite(project.HasProp("char") ? project.char : "", path, section, "Char")
            IniWrite(project.HasProp("path") ? project.path : "", path, section, "Path")
            IniWrite(project.HasProp("workPath") ? project.workPath : "", path, section, "WorkPath")
            idx += 1
        }
    } catch {
        ProjectData_Invalidate()
        return false
    }
    g_Projects := list
    g_ProjectDataCacheReady := true
    g_ProjectDataCacheMtime := ProjectData_FileMtime()
    return true
}

; Extract matching segments from project path for window title matching
; Cursor window titles have format: "filename - folder-name - Cursor" or "filename - path-segment - Cursor"
ExtractProjectMatchSegments(projectPath) {
    ; Normalize the project path (remove trailing backslashes)
    normalizedPath := RTrim(projectPath, "\")

    ; Split path into segments
    pathSegments := StrSplit(normalizedPath, "\")

    ; Extract the last folder name (e.g., "zmk-sofle", "26-ai-experiment", "scripts")
    lastSegment := pathSegments[pathSegments.Length]

    ; Build list of potential match strings
    matchSegments := [lastSegment]

    ; If we have at least 2 segments, also try the combination
    if (pathSegments.Length >= 2) {
        ; Try last two segments joined with " - " (for cases like "17 - Projects")
        lastTwoJoined := pathSegments[pathSegments.Length - 1] . " - " . pathSegments[pathSegments.Length]
        if (lastTwoJoined != lastSegment) {  ; Only add if different
            matchSegments.Push(lastTwoJoined)
        }
    }

    return matchSegments
}

ProjectData_Load()