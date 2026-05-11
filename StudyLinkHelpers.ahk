; StudyLink INI helpers for per-study subtopic links
StudyLink_IniPath() {
    return A_ScriptDir "\data\study_links.ini"
}

; Get the stored link for a study (by name or ID)
StudyLink_Get(studyKey) {
    iniPath := StudyLink_IniPath()
    try {
        url := IniRead(iniPath, "Links", studyKey, "")
    } catch {
        url := ""
    }
    return url
}

; Set the link for a study
StudyLink_Set(studyKey, url) {
    iniPath := StudyLink_IniPath()
    try IniWrite(url, iniPath, "Links", studyKey)
}

; Open the stored link for a study
StudyLink_Open(studyKey) {
    url := StudyLink_Get(studyKey)
    if (url != "") {
        try Run(url)
    } else {
        MsgBox "No link stored for this study."
    }
}
