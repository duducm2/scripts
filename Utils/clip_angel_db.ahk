; =============================================================================
; Utils module: clip_angel_db.ahk
; SQLite read/write helper (ClipAngelDb.exe). Failure => "" / false / 0 so
; callers can fall back to UIA when needed. Loaded via #include into Utils.ahk
; before favorite/merge modules that call these helpers.
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

; Direct CreateProcessW + anonymous stdout pipe (no cmd.exe, no temp files).
ClipAngelDb_ExecCapture(exe, args, workDir, timeoutMs, &exitCode) {
    exitCode := 1
    sa := Buffer(A_PtrSize = 8 ? 24 : 12, 0)
    NumPut("UInt", sa.Size, sa, 0)
    NumPut("Int", 1, sa, A_PtrSize = 8 ? 16 : 8) ; bInheritHandle = TRUE
    hRead := 0, hWrite := 0
    if !DllCall("CreatePipe", "Ptr*", &hRead, "Ptr*", &hWrite, "Ptr", sa, "UInt", 0)
        return ""
    DllCall("SetHandleInformation", "Ptr", hRead, "UInt", 1, "UInt", 0) ; clear HANDLE_FLAG_INHERIT

    hNull := DllCall("CreateFileW", "WStr", "NUL", "UInt", 0x40000000, "UInt", 3, "Ptr", sa, "UInt", 3, "UInt", 0, "Ptr", 0, "Ptr")
    if (hNull = -1)
        hNull := 0

    siSize := (A_PtrSize = 8) ? 104 : 68
    si := Buffer(siSize, 0)
    NumPut("UInt", siSize, si, 0)
    NumPut("UInt", 0x101, si, (A_PtrSize = 8) ? 60 : 44) ; STARTF_USESTDHANDLES | STARTF_USESHOWWINDOW
    NumPut("UShort", 0, si, (A_PtrSize = 8) ? 64 : 48)    ; SW_HIDE
    NumPut("Ptr", 0, si, (A_PtrSize = 8) ? 80 : 56)       ; hStdInput
    NumPut("Ptr", hWrite, si, (A_PtrSize = 8) ? 88 : 60)  ; hStdOutput
    NumPut("Ptr", hNull ? hNull : hWrite, si, (A_PtrSize = 8) ? 96 : 64) ; hStdError

    pi := Buffer((A_PtrSize = 8) ? 24 : 16, 0)
    cmdLine := '"' exe '"' (args != "" ? " " args : "")
    ok := DllCall("CreateProcessW", "Ptr", 0, "WStr", cmdLine, "Ptr", 0, "Ptr", 0,
        "Int", 1, "UInt", 0x08000000, "Ptr", 0, "WStr", workDir, "Ptr", si, "Ptr", pi)

    DllCall("CloseHandle", "Ptr", hWrite)
    if (hNull)
        DllCall("CloseHandle", "Ptr", hNull)
    if (!ok) {
        DllCall("CloseHandle", "Ptr", hRead)
        return ""
    }

    hProc := NumGet(pi, 0, "Ptr")
    hThread := NumGet(pi, A_PtrSize, "Ptr")
    DllCall("CloseHandle", "Ptr", hThread)

    raw := Buffer(0)
    rawLen := 0
    chunk := Buffer(4096)
    deadline := A_TickCount + timeoutMs
    timedOut := false
    loop {
        avail := 0
        while DllCall("PeekNamedPipe", "Ptr", hRead, "Ptr", 0, "UInt", 0, "Ptr", 0, "UInt*", &avail, "Ptr", 0) && avail > 0 {
            toRead := Min(avail, chunk.Size)
            bytesRead := 0
            if !DllCall("ReadFile", "Ptr", hRead, "Ptr", chunk, "UInt", toRead, "UInt*", &bytesRead, "Ptr", 0) || bytesRead = 0
                break
            oldLen := rawLen
            rawLen += bytesRead
            raw.Size := rawLen
            DllCall("RtlMoveMemory", "Ptr", raw.Ptr + oldLen, "Ptr", chunk.Ptr, "UPtr", bytesRead)
        }
        waitRes := DllCall("WaitForSingleObject", "Ptr", hProc, "UInt", 10, "UInt")
        if (waitRes = 0)
            break
        if (A_TickCount >= deadline) {
            timedOut := true
            DllCall("TerminateProcess", "Ptr", hProc, "UInt", 1)
            break
        }
    }
    if (!timedOut) {
        avail := 0
        while DllCall("PeekNamedPipe", "Ptr", hRead, "Ptr", 0, "UInt", 0, "Ptr", 0, "UInt*", &avail, "Ptr", 0) && avail > 0 {
            toRead := Min(avail, chunk.Size)
            bytesRead := 0
            if !DllCall("ReadFile", "Ptr", hRead, "Ptr", chunk, "UInt", toRead, "UInt*", &bytesRead, "Ptr", 0) || bytesRead = 0
                break
            oldLen := rawLen
            rawLen += bytesRead
            raw.Size := rawLen
            DllCall("RtlMoveMemory", "Ptr", raw.Ptr + oldLen, "Ptr", chunk.Ptr, "UPtr", bytesRead)
        }
        ec := 1
        DllCall("GetExitCodeProcess", "Ptr", hProc, "UInt*", &ec)
        exitCode := ec
    }
    DllCall("CloseHandle", "Ptr", hRead)
    DllCall("CloseHandle", "Ptr", hProc)
    if (timedOut || rawLen = 0)
        return ""
    return StrGet(raw, rawLen, "UTF-8")
}

; Run helper; returns stdout (trimmed) on exit 0, else "".
; waitnew may exit 2 on timeout — still returns stdout id when present; caller checks.
ClipAngelDb_Run(args, timeoutMs := 3000, allowExit2 := false) {
    exe := ClipAngelDb_ExePath()
    if (exe = "" || !FileExist(exe))
        return ""
    exitCode := 1
    text := ClipAngelDb_ExecCapture(exe, args, A_ScriptDir, timeoutMs, &exitCode)
    if (exitCode != 0 && !(allowExit2 && exitCode = 2))
        return ""
    return RTrim(text, "`r`n")
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

; Sets Favorite = 1/0 on a specific clip Id. Returns true on success.
ClipAngelDb_SetFav(id, isFav := 1) {
    if !(id is Integer) && !IsInteger(id)
        return false
    fav := isFav ? 1 : 0
    out := ClipAngelDb_Run("setfav " Integer(id) " " fav, 3500)
    if (out = "")
        return false
    parts := StrSplit(out, "`t")
    return parts.Length >= 1 && IsInteger(parts[1]) && Integer(parts[1]) > 0
}

; Marks the newest clip Favorite = 1/0 atomically. Returns true on success.
ClipAngelDb_FavNewest(isFav := 1, &clipId := 0, &title := "") {
    fav := isFav ? 1 : 0
    out := ClipAngelDb_Run("favnewest " fav, 3500)
    if (out = "")
        return false
    parts := StrSplit(out, "`t")
    if (parts.Length < 2 || !IsInteger(parts[1]) || Integer(parts[1]) <= 0)
        return false
    clipId := Integer(parts[1])
    title := parts.Length >= 3 ? parts[3] : ""
    return true
}

; Clears Favorite = 0 on all favorited clips. Returns true on success.
ClipAngelDb_UnfavAll(&count := 0) {
    out := ClipAngelDb_Run("unfavall", 4000)
    if (out = "" || !IsInteger(out))
        return false
    count := Integer(out)
    return true
}

; Deletes all non-favorite clips from ClipAngel's SQLite DB. Returns true on success.
ClipAngelDb_DeleteNonFav(&count := 0) {
    out := ClipAngelDb_Run("deletenonfav", 15000)
    if (out = "" || !IsInteger(out))
        return false
    count := Integer(out)
    return true
}
