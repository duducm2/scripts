; =============================================================================
; Utils module: project_data_cursor.ahk
; Project data for Cursor window focus selector
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; Project Data (for Cursor Window Focus Selector)
; Ported from WindowManagement.ahk to ensure consistent key mapping
; =============================================================================

; Character sequence for assignment: 1 2 3 4 5 q w e r t a s d f g z x c v b 6 7 8 9 0 y u i o p h j k l n m , .
global g_ProjectCharSequence := ["1", "2", "3", "4", "5", "q", "w", "e", "r", "t", "a", "s", "d", "f", "g", "z", "x",
    "c", "v", "b", "6", "7", "8", "9", "0", "y", "u", "i", "o", "p", "h", "j", "k", "l", "n", "m", ",", "."]

; Category display order (General first, Personal second, Work last)
global g_ProjectCategories := ["General", "Personal", "Work"]

; Global project list - must match WindowManagement.ahk for consistent key mapping
; Each project should have: name, path, workPath, and category ("General", "Personal", or "Work")
global g_Projects := [
    ; General category
    { name: "Scripts", path: "C:\Users\eduev\Meu Drive\17 - Projects\scripts", workPath: "C:\Users\fie7ca\Documents\scripts",
        category: "General" }, { name: "14-my-notes", path: "C:\Users\eduev\Meu Drive\17 - Projects\notes", workPath: "C:\Users\fie7ca\OneDrive - Bosch Group\14-my-notes",
            category: "General" }, { name: "", path: "", workPath: "", category: "General" }, { name: "", path: "",
                workPath: "", category: "General" }, { name: "", path: "", workPath: "", category: "General" },
                ; Personal category
                { name: "ZMK Sofle", path: "C:\Users\eduev\Documents\ZMK\zmk-sofle", workPath: "", category: "Personal" }, { name: "AI Experiment",
                    path: "C:\Users\eduev\Meu Drive\04 - Pós-graduação\01 - Mestrado\26-ai-experiment", workPath: "",
                    category: "Personal" }, { name: "my-personal-repo", path: "C:\Users\eduev\Meu Drive\17 - Projects\my-personal-repo",
                        workPath: "C:\Users\fie7ca\OneDrive - Bosch Group\13 - General workspace\my-personal-repo",
                        category: "Personal" }, { name: "",
                            path: "", workPath: "", category: "Personal" }, { name: "", path: "", workPath: "",
                                category: "Personal" },
                            ; Work category
                            { name: "dashboard-model-research", path: "", workPath: "C:\Users\fie7ca\OneDrive - Bosch Group\13 - General workspace\GS_E&S_CIP Dashboard research and design workspace folder\dashboard-model-research",
                                category: "Work" }, { name: "GS_UX core team_UX and CIP Integration", path: "",
                                    workPath: "C:\Users\fie7ca\OneDrive - Bosch Group\13 - General workspace\GS_UX core team_UX and CIP Integration",
                                    category: "Work" }, { name: "🪂 Avante", path: "", workPath: "C:\Users\fie7ca\OneDrive - Bosch Group\General - GS_BDU_Team\00_UX_GS_Team\AM_Planning\Avante",
                                        category: "Work" }, { name: "Piloto PT B2B", path: "", workPath: "C:\Users\fie7ca\OneDrive - Bosch Group\SO UX - LA (Internal) - Data Insights SO - Piloto PT B2B",
                                            category: "Work" }, { name: "Python ScripTs", path: "C:\Users\eduev\Meu Drive\17 - Projects\My-Python-Scripts",
                                                workPath: "C:\Users\fie7ca\OneDrive - Bosch Group\17 - Python Scripts",
                                                category: "Work", char: "t" }, { name: "",
                                                    path: "", workPath: "", category: "Work" }
]

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
