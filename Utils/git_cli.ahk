; =============================================================================
; Shared git CLI helpers (RunWaitWithTimeout). Used by Act.ahk and editor Alt+S.
; =============================================================================

; Run a command with a timeout.
; Returns the process exit code, or 124 on timeout.
RunWaitWithTimeout(cmd, workingDir := "", options := "", timeoutMs := 120000) {
    safeWorkDir := StrReplace(workingDir, "'", "''")
    safeCmd := StrReplace(cmd, "'", "''")

    ps := ""
        . "$ErrorActionPreference='Stop';"
        . "$cmd='" . safeCmd . "';"
        . "$wd='" . safeWorkDir . "';"
        . "$t=[int]" . timeoutMs . ";"
        .
        "$p=Start-Process -FilePath 'cmd.exe' -ArgumentList @('/v:on','/c',$cmd) -WorkingDirectory $wd -PassThru -WindowStyle Hidden;"
        . "if(-not $p.WaitForExit($t)){try{$p.Kill()}catch{}; exit 124};"
        . "exit $p.ExitCode"

    try return RunWait("powershell.exe -NoProfile -ExecutionPolicy Bypass -Command " . Chr(34) . ps . Chr(34),
    workingDir, options)
    catch {
        return 1
    }
}

; Returns git repo top-level path or "" if not a repo.
GitCli_RevParseTopLevel(repoDir, timeoutMs := 15000) {
    if !repoDir || !DirExist(repoDir)
        return ""
    cmd := 'set GIT_TERMINAL_PROMPT=0& set GCM_INTERACTIVE=Never& git rev-parse --show-toplevel'
    outFile := A_Temp "\git-rev-parse-" A_TickCount ".txt"
    errFile := A_Temp "\git-rev-parse-err-" A_TickCount ".txt"
    fullCmd := cmd . ' 1>"' outFile '" 2>"' errFile '"'
    exitCode := RunWaitWithTimeout(fullCmd, repoDir, "Hide", timeoutMs)
    top := ""
    try {
        if FileExist(outFile)
            top := Trim(FileRead(outFile, "UTF-8"), "`r`n `t")
    } catch {
    }
    try FileDelete(outFile)
    try FileDelete(errFile)
    if (exitCode != 0 || top = "")
        return ""
    return top
}

; Capture git stdout (trimmed). Returns "" on failure/timeout.
; Uses git -C and a temp working dir so repo paths with spaces do not break cmd.exe.
GitCli_CaptureStdout(repoDir, gitArgs, timeoutMs := 15000) {
    if !repoDir || !DirExist(repoDir) || gitArgs = ""
        return ""
    quotedRepo := StrReplace(repoDir, '"', '')
    cmd := Format('set GIT_TERMINAL_PROMPT=0& set GCM_INTERACTIVE=Never& git -C "{1}" {2}', quotedRepo, gitArgs)
    outFile := A_Temp "\git-cli-out-" A_TickCount ".txt"
    errFile := A_Temp "\git-cli-err-" A_TickCount ".txt"
    fullCmd := cmd . ' 1>"' outFile '" 2>"' errFile '"'
    exitCode := RunWaitWithTimeout(fullCmd, A_Temp, "Hide", timeoutMs)
    out := ""
    try {
        if FileExist(outFile)
            out := Trim(FileRead(outFile, "UTF-8"), "`r`n `t")
    } catch {
    }
    try FileDelete(outFile)
    try FileDelete(errFile)
    if (exitCode != 0)
        return ""
    return out
}

; Commits local HEAD is behind @{upstream}, or -1 if the check failed.
GitCli_BehindUpstreamCount(repoDir, timeoutMs := 15000) {
    out := GitCli_CaptureStdout(repoDir, 'rev-list --count "HEAD..@{upstream}"', timeoutMs)
    if (out = "" || !RegExMatch(out, "^\d+$"))
        return -1
    return Integer(out)
}
