; =============================================================================
; Utils module: clip_angel_db.ahk
; Read-only ClipAngel SQLite helper (ClipAngelDb.exe). Failure => "" / false so
; callers fall back to UIA. Loaded via #include into Utils.ahk before favorite/
; merge modules that call these helpers.
; =============================================================================

CLIPANGEL_DB_EXE_REL := "Utils\bin\ClipAngelDb.exe"

ClipAngelDb_ExePath() {
    global CLIPANGEL_DB_EXE_REL
    ; Prefer path relative to Utils.ahk host (A_ScriptDir = scripts root).
    p := A_ScriptDir "\" CLIPANGEL_DB_EXE_REL
    if FileExist(p)
        return p
    ; Fallback when this file is run standalone from Utils\.
    p2 := A_ScriptDir "\bin\ClipAngelDb.exe"
    if FileExist(p2)
        return p2
    return ""
}

; Run helper; returns stdout (trimmed) on exit 0, else "".
; waitnew may exit 2 on timeout — still returns stdout id when present; caller checks.
ClipAngelDb_Run(args, timeoutMs := 3000, allowExit2 := false) {
    exe := ClipAngelDb_ExePath()
    if (exe = "" || !FileExist(exe))
        return ""
    stamp := A_TickCount "_" Random(1000, 9999)
    outFile := A_Temp "\clipangeldb_" stamp ".txt"
    errFile := A_Temp "\clipangeldb_err_" stamp ".txt"
    try {
        ; cmd.exe /c with nested quotes for paths that contain spaces.
        full := A_ComSpec ' /c ""' exe '" ' args ' >"' outFile '" 2>"' errFile '""'
        exitCode := 1
        try exitCode := RunWait(full, A_ScriptDir, "Hide")
        catch
            return ""
        if (exitCode != 0 && !(allowExit2 && exitCode = 2))
            return ""
        if !FileExist(outFile)
            return ""
        text := ""
        try text := FileRead(outFile, "UTF-8")
        catch {
            try text := FileRead(outFile)
            catch
                text := ""
        }
        return RTrim(text, "`r`n")
    } finally {
        try FileDelete(outFile)
        catch {
        }
        try FileDelete(errFile)
        catch {
        }
    }
}

ClipAngelDb_MaxId() {
    out := ClipAngelDb_Run("maxid", 2000)
    if (out = "" || !IsInteger(out))
        return ""
    return out
}

; Wait until MAX(Id) > prevId. Returns new id string, or "" on failure/timeout.
ClipAngelDb_WaitNew(prevId, timeoutMs := 600) {
    if !(prevId is Integer) && !IsInteger(prevId)
        return ""
    if !(timeoutMs is Integer)
        timeoutMs := 600
    ; allowExit2: timeout still prints current max id — treat empty/unchanged as fail.
    out := ClipAngelDb_Run("waitnew " Integer(prevId) " " Integer(timeoutMs), timeoutMs + 1500, true)
    if (out = "" || !IsInteger(out))
        return ""
    if (Integer(out) <= Integer(prevId))
        return ""
    return out
}

; Returns Map("id","type","favorite","title") or 0 on failure.
ClipAngelDb_Newest() {
    out := ClipAngelDb_Run("newest", 2000)
    if (out = "")
        return 0
    parts := StrSplit(out, "`t")
    if (parts.Length < 3)
        return 0
    id := parts[1]
    if !IsInteger(id)
        return 0
    fav := (parts.Length >= 3 && (parts[3] = "1" || parts[3] = "true"))
    title := parts.Length >= 4 ? parts[4] : ""
    return Map("id", Integer(id), "type", parts[2], "favorite", fav, "title", title)
}

; Returns 1 / 0 / "" (unknown/failure).
ClipAngelDb_IsFav(id) {
    if !(id is Integer) && !IsInteger(id)
        return ""
    out := ClipAngelDb_Run("isfav " Integer(id), 2000)
    if (out = "1" || out = "0")
        return Integer(out)
    return ""
}

; Returns Map("boundaryId","total","skipped","text") or 0 on failure.
ClipAngelDb_MergePayload() {
    out := ClipAngelDb_Run("mergepayload", 5000)
    if (out = "")
        return 0
    nl := InStr(out, "`n")
    if (!nl) {
        header := Trim(out, "`r")
        body := ""
    } else {
        header := Trim(SubStr(out, 1, nl - 1), "`r")
        body := SubStr(out, nl + 1)
        if (SubStr(body, 1, 1) = "`r")
            body := SubStr(body, 2)
    }
    hp := StrSplit(header, "`t")
    if (hp.Length < 3 || !IsInteger(hp[1]) || !IsInteger(hp[2]) || !IsInteger(hp[3]))
        return 0
    return Map(
        "boundaryId", Integer(hp[1]),
        "total", Integer(hp[2]),
        "skipped", Integer(hp[3]),
        "text", body
    )
}
