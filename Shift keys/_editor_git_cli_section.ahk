
; =============================================================================
; Alt+S — Git stash / fetch / pull via native CLI (Editor-GitStashFetchPull.ps1)
; =============================================================================
global EDITOR_GIT_CLI_TIMEOUT_MS := 120000

Editor_GitFlowFail(step, reason := "") {
    msg := "❌ " step " failed"
    if (reason != "")
        msg .= ": " reason
    try StandardLoadingBar_Update(msg, BANNER_ACCENT_ERROR)
    try StandardLoadingBar_Hide(1200)
}

Editor_GetEditorProcessCommandLine(editorHwnd := 0) {
    if !editorHwnd
        editorHwnd := WinExist("A")
    if !editorHwnd
        return ""
    pid := 0
    try pid := WinGetPID("ahk_id " editorHwnd)
    if !pid
        return ""
    cmd := ""
    try {
        locator := ComObject("WbemScripting.SWbemLocator")
        svc := locator.ConnectServer(".", "root\cimv2")
        for proc in svc.ExecQuery("SELECT CommandLine FROM Win32_Process WHERE ProcessId = " pid) {
            try cmd := proc.CommandLine
            break
        }
    } catch {
    }
    return cmd ? cmd : ""
}

Editor_DecodeFolderUri(uri) {
    if (uri = "" || !InStr(uri, "file:"))
        return ""
    path := RegExReplace(uri, "i)^file://+", "")
    path := StrReplace(path, "/", "\")
    if RegExMatch(path, "^[A-Za-z]/")
        path := SubStr(path, 1, 1) ":" SubStr(path, 2)
    loop {
        if !RegExMatch(path, "%([0-9A-Fa-f]{2})", &m)
            break
        path := StrReplace(path, m[0], Chr("0x" m[1]))
    }
    return path
}

Editor_ExtractFolderPathsFromCmdLine(cmdLine) {
    paths := []
    seen := Map()
    if (cmdLine = "")
        return paths
    pos := 1
    while RegExMatch(cmdLine, "i)--folder-uri\s+(\S+)", &m, pos) {
        decoded := Editor_DecodeFolderUri(m[1])
        if decoded && DirExist(decoded) && !seen.Has(decoded) {
            seen[decoded] := true
            paths.Push(decoded)
        }
        pos := m.Pos(0) + m.Len(0)
    }
    pos := 1
    while RegExMatch(cmdLine, '"([A-Za-z]:[^"]+)"', &m, pos) {
        p := m[1]
        if DirExist(p) && !seen.Has(p) {
            seen[p] := true
            paths.Push(p)
        }
        pos := m.Pos(0) + m.Len(0)
    }
    return paths
}

Editor_GetScmRepoBasenameFromStatusBar(editorHwnd := 0) {
    try {
        if !editorHwnd
            editorHwnd := WinExist("A")
        if !editorHwnd
            return ""
        root := UIA.ElementFromHandle(editorHwnd)
        if !root
            return ""
        el := root.FindFirst({ AutomationId: "status.scm.0" })
        if !el
            return ""
        if RegExMatch(el.Name, "^([^(]+)\s*\(Git\)", &m)
            return Trim(m[1])
    } catch {
    }
    return ""
}

Editor_DeduplicatePaths(paths) {
    out := []
    seen := Map()
    for p in paths {
        if !p || seen.Has(p)
            continue
        seen[p] := true
        out.Push(p)
    }
    return out
}

Editor_AppendProjectRegistryPaths(&paths) {
    if !IsSet(g_Projects) || !IsObject(g_Projects)
        return
    try {
        loop g_Projects.Length {
            project := g_Projects[A_Index]
            for key in ["path", "workPath"] {
                p := ""
                try {
                    if IsObject(project) && project.HasProp(key)
                        p := project.%key%
                } catch {
                }
                if p && DirExist(p)
                    paths.Push(p)
            }
        }
    } catch {
    }
}

Editor_PickGitRootByBasename(candidates, basename) {
    if !basename
        return ""
    baseLow := StrLower(basename)
    for p in candidates {
        SplitPath p, , , &dirName
        if (StrLower(dirName) = baseLow) {
            top := GitCli_RevParseTopLevel(p)
            if top
                return top
        }
    }
    return ""
}

Editor_ResolveGitRepoDir(editorHwnd := 0) {
    if !editorHwnd
        editorHwnd := WinExist("A")
    if !editorHwnd
        return ""
    cmdLine := Editor_GetEditorProcessCommandLine(editorHwnd)
    candidates := Editor_ExtractFolderPathsFromCmdLine(cmdLine)
    Editor_AppendProjectRegistryPaths(&candidates)
    candidates := Editor_DeduplicatePaths(candidates)
    if !candidates.Length
        return ""
    if (candidates.Length = 1)
        return GitCli_RevParseTopLevel(candidates[1])
    basename := Editor_GetScmRepoBasenameFromStatusBar(editorHwnd)
    if basename {
        picked := Editor_PickGitRootByBasename(candidates, basename)
        if picked
            return picked
    }
    for p in candidates {
        top := GitCli_RevParseTopLevel(p)
        if top
            return top
    }
    return ""
}

Editor_UnescapeJsonString(s) {
    s := StrReplace(s, '\n', "`n")
    s := StrReplace(s, '\r', "`r")
    s := StrReplace(s, '\t', "`t")
    s := StrReplace(s, '\"', '"')
    s := StrReplace(s, '\\', '\')
    return s
}

Editor_ParseGitScriptJson(raw) {
    result := Map(
        "ok", false,
        "failedStep", "",
        "error", "",
        "stashPopWarning", false
    )
    if RegExMatch(raw, '"ok"\s*:\s*(true|false)', &m)
        result["ok"] := (m[1] = "true")
    if RegExMatch(raw, '"failedStep"\s*:\s*"([^"]*)"', &m)
        result["failedStep"] := m[1]
    if RegExMatch(raw, '"error"\s*:\s*"((?:\\.|[^"\\])*)"', &m)
        result["error"] := Editor_UnescapeJsonString(m[1])
    if RegExMatch(raw, '"stashPopWarning"\s*:\s*(true|false)', &m)
        result["stashPopWarning"] := (m[1] = "true")
    return result
}

Editor_TruncateGitError(msg, maxLen := 180) {
    msg := Trim(msg, "`r`n `t")
    if (StrLen(msg) <= maxLen)
        return msg
    return SubStr(msg, 1, maxLen - 3) "..."
}

Editor_RunGitStashFetchPullScript(repoDir, timeoutMs := 0) {
    global EDITOR_GIT_CLI_TIMEOUT_MS
    if !timeoutMs
        timeoutMs := EDITOR_GIT_CLI_TIMEOUT_MS
    ps1 := A_ScriptDir "\infra\tools\Editor-GitStashFetchPull.ps1"
    result := Map("ok", false, "failedStep", "Git", "error", "", "stashPopWarning", false)
    if !FileExist(ps1) {
        result["failedStep"] := "Script"
        result["error"] := "Editor-GitStashFetchPull.ps1 not found"
        return result
    }
    resultPath := A_Temp "\editor-git-" A_TickCount ".json"
    timeoutSec := Max(30, Round(timeoutMs / 1000))
    cmd := 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "' ps1 '" -RepoDir "' repoDir
        . '" -ResultPath "' resultPath '" -TimeoutSec ' timeoutSec
    exitCode := RunWaitWithTimeout(cmd, repoDir, "Hide", timeoutMs)
    try {
        if FileExist(resultPath) {
            raw := FileRead(resultPath, "UTF-8")
            FileDelete(resultPath)
            if (raw != "") {
                parsed := Editor_ParseGitScriptJson(raw)
                if parsed
                    return parsed
            }
        }
    } catch {
    }
    if (exitCode = 124)
        result["error"] := "timed out"
    else if (exitCode != 0)
        result["error"] := "git script exit " exitCode
    return result
}

Editor_GitStashAndPull() {
    global EDITOR_GIT_CLI_TIMEOUT_MS
    hwnd := WinExist("A")
    if !hwnd
        return
    if !(WinActive("ahk_exe Cursor.exe") || WinActive("ahk_exe Code.exe")) {
        Editor_GitFlowFail("Git", "active window is not Cursor or VS Code")
        return
    }
    repoDir := Editor_ResolveGitRepoDir(hwnd)
    if !repoDir {
        Editor_GitFlowFail("Repo", "could not resolve git root for active editor")
        return
    }
    StandardLoadingBar_Show("⏳ Git stash, fetch, and pull…", BANNER_ACCENT_INTERMEDIATE, { passive: false,
        centerOnHwnd: hwnd })
    try StandardLoadingBar_Update("⏳ Running git in " repoDir, BANNER_ACCENT_INTERMEDIATE)
    result := Editor_RunGitStashFetchPullScript(repoDir, EDITOR_GIT_CLI_TIMEOUT_MS)
    if !result["ok"] {
        step := result.Has("failedStep") && result["failedStep"] != "" ? result["failedStep"] : "Git"
        err := result.Has("error") ? Editor_TruncateGitError(result["error"]) : "unknown error"
        Editor_GitFlowFail(step, err)
        return
    }
    if result.Has("stashPopWarning") && result["stashPopWarning"]
        try ShowCenteredOverlay_Utils("ℹ Stashed changes remain on stack", 2200, BANNER_ACCENT_INFO)
    StandardLoadingBar_Update("✅ Pull complete", BANNER_ACCENT_SUCCESS)
    try {
        soundPath := A_ScriptDir "\assets\sounds\pull-successful.wav"
        if FileExist(soundPath)
            ScriptSoundPlay(soundPath, true)
    } catch {
    }
    StandardLoadingBar_Hide(600)
}
