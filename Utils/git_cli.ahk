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
    r := GitCli_Run(repoDir, gitArgs, timeoutMs)
    return (r.exitCode = 0) ? r.stdout : ""
}

; Run git from repoDir as working directory (avoids git -C paths with spaces breaking via PowerShell).
GitCli_Run(repoDir, gitArgs, timeoutMs := 15000) {
    if !repoDir || !DirExist(repoDir) || gitArgs = ""
        return { exitCode: 1, stdout: "", stderr: "" }
    cmd := Format("set GIT_TERMINAL_PROMPT=0& set GCM_INTERACTIVE=Never& git {1}", gitArgs)
    outFile := A_Temp "\git-cli-out-" A_TickCount ".txt"
    errFile := A_Temp "\git-cli-err-" A_TickCount ".txt"
    fullCmd := cmd . ' 1>"' outFile '" 2>"' errFile '"'
    exitCode := RunWaitWithTimeout(fullCmd, repoDir, "Hide", timeoutMs)
    out := ""
    err := ""
    try {
        if FileExist(outFile)
            out := Trim(FileRead(outFile, "UTF-8"), "`r`n `t")
    } catch {
    }
    try {
        if FileExist(errFile)
            err := Trim(FileRead(errFile, "UTF-8"), "`r`n `t")
    } catch {
    }
    try FileDelete(outFile)
    try FileDelete(errFile)
    return { exitCode: exitCode, stdout: out, stderr: err }
}

; Commits local HEAD is behind @{upstream}, or -1 if the check failed.
GitCli_BehindUpstreamCount(repoDir, timeoutMs := 15000) {
    out := GitCli_CaptureStdout(repoDir, 'rev-list --count "HEAD..@{upstream}"', timeoutMs)
    if (out = "" || !RegExMatch(out, "^\d+$"))
        return -1
    return Integer(out)
}

; my-personal-repo clone for the current machine. Built-in paths so a work PC
; picks up the Bosch clone after a scripts pull (env.ahk is gitignored).
; PERSONAL_REPO_PATH_WORK / PERSONAL_REPO_PATH_PERSONAL in env.ahk override when set.
; Returns the path when the folder exists; else "" (Act aborts, Utility [G] soft-skips).
; Do not define this in env.ahk.
GetPersonalRepoPath() {
    global IS_WORK_ENVIRONMENT, PERSONAL_REPO_PATH_WORK, PERSONAL_REPO_PATH_PERSONAL
    work := "C:\Users\fie7ca\OneDrive - Bosch Group\13 - General workspace\my-personal-repo"
    personal := "C:\Users\eduev\Meu Drive\17 - Projects\my-personal-repo"
    if (IsSet(PERSONAL_REPO_PATH_WORK)) {
        override := RTrim(Trim(PERSONAL_REPO_PATH_WORK), "\")
        if (override != "")
            work := override
    }
    if (IsSet(PERSONAL_REPO_PATH_PERSONAL)) {
        override := RTrim(Trim(PERSONAL_REPO_PATH_PERSONAL), "\")
        if (override != "")
            personal := override
    }
    isWork := IsSet(IS_WORK_ENVIRONMENT) && IS_WORK_ENVIRONMENT
    path := isWork ? work : personal
    if (path != "" && DirExist(path))
        return path
    other := isWork ? personal : work
    if (other != "" && DirExist(other))
        return other
    return ""
}
