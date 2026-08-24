; =============================================================================
; Utils module: prompt_usage.ahk
; Append-only prompt usage log
; =============================================================================

PromptUsage_LogPath() {
    return A_ScriptDir "\assets\data\prompt_usage.log"
}

PromptUsage_Log(prompt, extra := "", pickedCount := 0) {
    if (!IsObject(prompt))
        return
    path := PromptUsage_LogPath()
    try DirCreate(A_ScriptDir "\assets\data")
    catch {
    }
    ts := FormatTime(, "yyyy-MM-dd HH:mm:ss")
    ch := prompt.HasProp("char") ? prompt.char : ""
    name := prompt.HasProp("name") ? prompt.name : ""
    mode := prompt.HasProp("pasteMode") ? prompt.pasteMode : ""
    companion := ""
    try companion := ResolveGlobalAICompanion()
    catch {
    }
    ctxCount := PromptData_ContextEntriesForCurrentEnv(prompt).Length
    ctxLabel := "ctx=" . ctxCount
    if (pickedCount > 0)
        ctxLabel .= "+" . pickedCount
    line := ts . "|" . ch . "|" . name . "|" . companion . "|" . mode . "|" . ctxLabel
    if (extra != "")
        line .= "|" . extra
    line .= "`n"
    try FileAppend(line, path, "UTF-8")
    catch {
    }
}
