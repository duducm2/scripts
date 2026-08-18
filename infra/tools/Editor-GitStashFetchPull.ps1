# Alt+S editor git flow: stash (single -u), fetch, ff-only pull, stash pop.
# Success (ok=true / exit 0) only if every quality gate passes — never on git exit 0 alone.
# Writes JSON result to -ResultPath.
# -SkipPull is test-only: fetch then skip pull so the behind-count gate must fail.
param(
    [Parameter(Mandatory = $true)]
    [string]$RepoDir,
    [Parameter(Mandatory = $true)]
    [string]$ResultPath,
    [int]$TimeoutSec = 120,
    [switch]$SkipPull
)

$ErrorActionPreference = 'Continue'
$env:GIT_TERMINAL_PROMPT = '0'
$env:GCM_INTERACTIVE = 'Never'

$stashMsg = "alt+s-auto-$(Get-Date -Format 'yyyyMMddHHmmss')-$PID"

$result = [ordered]@{
    ok              = $false
    didStash        = $false
    stashPopWarning = $false
    failedStep      = ''
    error           = ''
    behindCount     = -1
    aheadCount      = -1
    upstream        = ''
    stashRef        = ''
    steps           = @()
}

function Write-Result {
    param([int]$ExitCode = 1)
    try {
        $json = $result | ConvertTo-Json -Compress -Depth 6
        Set-Content -LiteralPath $ResultPath -Value $json -Encoding UTF8
    } catch {
        exit $ExitCode
    }
    exit $ExitCode
}

function Add-Step {
    param([string]$Name, [bool]$Ok, [string]$Detail = '')
    $result.steps += [ordered]@{ name = $Name; ok = $Ok; detail = $Detail }
}

function Fail-Step {
    param([string]$Step, [string]$ErrorText)
    $result.failedStep = $Step
    $result.error = $ErrorText
    if ($Step -eq 'StashPop') {
        $result.stashPopWarning = $true
    }
    Add-Step $Step $false $ErrorText
    Write-Result 1
}

function Invoke-GitRaw {
    param([string]$GitArgs)
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = 'git'
    $psi.Arguments = "-C `"$RepoDir`" $GitArgs"
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.EnvironmentVariables['GIT_TERMINAL_PROMPT'] = '0'
    $psi.EnvironmentVariables['GCM_INTERACTIVE'] = 'Never'
    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $psi
    $null = $p.Start()
    $ms = [Math]::Max(1000, $TimeoutSec * 1000)
    if (-not $p.WaitForExit($ms)) {
        try { $p.Kill() } catch {}
        return [pscustomobject]@{ Ok = $false; ExitCode = 124; Output = "Timed out after ${TimeoutSec}s" }
    }
    $stdout = $p.StandardOutput.ReadToEnd()
    $stderr = $p.StandardError.ReadToEnd()
    $combined = ($stdout + "`n" + $stderr).Trim()
    $code = $p.ExitCode
    return [pscustomobject]@{
        Ok       = ($code -eq 0)
        ExitCode = $code
        Output   = $combined
        StdOut   = $stdout.Trim()
    }
}

function Invoke-GitStep {
    param([string]$StepName, [string]$GitArgs)
    $r = Invoke-GitRaw $GitArgs
    $detail = if ($r.Output) { $r.Output.Substring(0, [Math]::Min(500, $r.Output.Length)) } else { '' }
    Add-Step $StepName $r.Ok $detail
    if (-not $r.Ok -and $r.ExitCode -eq 124) {
        Fail-Step $StepName $r.Output
    }
    $r | Add-Member -NotePropertyName StepOk -NotePropertyValue $r.Ok -Force
    return $r
}

function Get-Porcelain {
    $r = Invoke-GitRaw 'status --porcelain'
    if (-not $r.Ok) {
        Fail-Step 'Status' $(if ($r.Output) { $r.Output } else { 'git status failed' })
    }
    return $r.StdOut
}

function Test-HasStashableChanges {
    Invoke-GitRaw 'update-index --refresh' | Out-Null
    $wt = Invoke-GitRaw 'diff --quiet --ignore-submodules=dirty HEAD'
    if ($wt.ExitCode -eq 124) {
        Fail-Step 'Status' $wt.Output
    }
    if ($wt.ExitCode -ne 0) { return $true }
    $st = Invoke-GitRaw 'diff --quiet --cached --ignore-submodules=dirty'
    if ($st.ExitCode -eq 124) {
        Fail-Step 'Status' $st.Output
    }
    if ($st.ExitCode -ne 0) { return $true }
    $others = Invoke-GitRaw 'ls-files -o --exclude-standard'
    if (-not $others.Ok) {
        Fail-Step 'Status' $(if ($others.Output) { $others.Output } else { 'git ls-files failed' })
    }
    if ($others.StdOut -and ($others.StdOut -match '\S')) { return $true }
    return $false
}

function Get-GitIndexLockPath {
    $r = Invoke-GitRaw 'rev-parse --absolute-git-dir'
    if (-not $r.Ok -or -not $r.StdOut) {
        Fail-Step 'Repo' $(if ($r.Output) { $r.Output } else { 'cannot resolve .git directory' })
    }
    return (Join-Path $r.StdOut.Trim() 'index.lock')
}

function Wait-IndexLock {
    param([int]$WaitSec = 60)
    $lock = Get-GitIndexLockPath
    $deadline = [datetime]::UtcNow.AddSeconds($WaitSec)
    while (Test-Path -LiteralPath $lock) {
        if ([datetime]::UtcNow -gt $deadline) {
            Fail-Step 'Stash' "index.lock still present after ${WaitSec}s: $lock"
        }
        Start-Sleep -Milliseconds 250
    }
}

function Wait-WorkingTreePullable {
    param([int]$WaitSec = 30)
    $deadline = [datetime]::UtcNow.AddSeconds($WaitSec)
    $last = ''
    while ($true) {
        Wait-IndexLock
        $last = Get-Porcelain
        if (-not (Test-HasStashableChanges)) {
            return $last
        }
        if ([datetime]::UtcNow -ge $deadline) {
            Fail-Step 'Stash' "working tree still not pullable after stash: $last"
        }
        Start-Sleep -Milliseconds 500
    }
}

function Get-StashListText {
    $r = Invoke-GitRaw 'stash list'
    if (-not $r.Ok) {
        Fail-Step 'Stash' $(if ($r.Output) { $r.Output } else { 'git stash list failed' })
    }
    return $r.StdOut
}

function Get-StashLineCount {
    param([string]$Text)
    if (-not $Text) { return 0 }
    return @($Text -split "`r?`n" | Where-Object { $_ -match '\S' }).Count
}

function Update-AheadBehind {
    $behind = Invoke-GitRaw 'rev-list --count HEAD..@{upstream}'
    $ahead = Invoke-GitRaw 'rev-list --count @{upstream}..HEAD'
    if ($behind.Ok -and $behind.StdOut -match '^\d+$') {
        $result.behindCount = [int]$behind.StdOut
    } else {
        $result.behindCount = -1
    }
    if ($ahead.Ok -and $ahead.StdOut -match '^\d+$') {
        $result.aheadCount = [int]$ahead.StdOut
    } else {
        $result.aheadCount = -1
    }
}

function Test-RefExists {
    param([string]$Ref)
    $r = Invoke-GitRaw "rev-parse -q --verify $Ref"
    return $r.Ok
}

if (-not (Test-Path -LiteralPath $RepoDir)) {
    Fail-Step 'Repo' "Directory not found: $RepoDir"
}

$inside = Invoke-GitStep 'rev-parse' 'rev-parse --is-inside-work-tree'
if (-not $inside.Ok -or ($inside.StdOut -ne 'true')) {
    Fail-Step 'Repo' $(if ($inside.Output) { $inside.Output } else { 'Not a git repository' })
}

$sym = Invoke-GitRaw 'symbolic-ref -q HEAD'
if (-not $sym.Ok) {
    Fail-Step 'Detached' 'HEAD is detached; checkout a branch before Alt+S'
}
Add-Step 'preflight-head' $true $sym.StdOut

if (Test-RefExists 'MERGE_HEAD') {
    Fail-Step 'Merge' 'merge in progress'
}
if ((Test-RefExists 'REBASE_HEAD') -or (Test-RefExists 'REBASE_MERGE') -or (Test-RefExists 'REBASE_APPLY')) {
    Fail-Step 'Rebase' 'rebase in progress'
}
if (Test-RefExists 'CHERRY_PICK_HEAD') {
    Fail-Step 'CherryPick' 'cherry-pick in progress'
}
Add-Step 'preflight-state' $true 'no merge/rebase/cherry-pick'

$up = Invoke-GitRaw 'rev-parse --abbrev-ref --symbolic-full-name @{upstream}'
if (-not $up.Ok -or -not $up.StdOut) {
    Fail-Step 'Upstream' $(if ($up.Output) { $up.Output } else { 'no upstream branch' })
}
$result.upstream = $up.StdOut
Add-Step 'preflight-upstream' $true $result.upstream
Update-AheadBehind

$stashListBefore = Get-StashListText
$dirty = Test-HasStashableChanges

if ($dirty) {
    $push = Invoke-GitStep 'stash' "stash push -u -m `"$stashMsg`""
    if ($push.Output -match '(?i)No local changes to save') {
        Wait-WorkingTreePullable | Out-Null
        Add-Step 'stash-skip' $true 'stash reported no local changes'
    } else {
    if ($push.Output -notmatch '(?i)Saved working directory') {
        Fail-Step 'Stash' $(if ($push.Output) { $push.Output } else { 'stash did not save working directory' })
    }
    Wait-IndexLock
    $stashListAfter = Get-StashListText
    if ((Get-StashLineCount $stashListAfter) -le (Get-StashLineCount $stashListBefore)) {
        Fail-Step 'Stash' 'stash list did not grow'
    }
    $top = Invoke-GitRaw 'stash list -1'
    if (-not $top.Ok -or ($top.StdOut -notmatch [regex]::Escape($stashMsg))) {
        Fail-Step 'Stash' "new stash@{0} is not ours: $($top.StdOut)"
    }
    $result.didStash = $true
    $result.stashRef = 'stash@{0}'
    Wait-WorkingTreePullable | Out-Null
    Add-Step 'stash-gate' $true $stashMsg
    }
} else {
    Add-Step 'stash-skip' $true 'clean working tree'
}

$fetch = Invoke-GitStep 'fetch' 'fetch'
if (-not $fetch.Ok) {
    Fail-Step 'Fetch' $(if ($fetch.Output) { $fetch.Output } else { 'git fetch failed' })
}
Add-Step 'fetch-gate' $true 'fetch exit 0'
Update-AheadBehind

function Invoke-PullOnce {
    param([string]$StepName)
    if ($SkipPull) {
        Add-Step $StepName $false 'skipped (test hook)'
        return [pscustomobject]@{ Ok = $false; Output = 'SkipPull'; StepOk = $false }
    }
    return Invoke-GitStep $StepName 'pull --ff-only'
}

function Test-BehindGatePass {
    Update-AheadBehind
    return ($result.behindCount -eq 0)
}

$pull = Invoke-PullOnce 'pull'
if (-not (Test-BehindGatePass)) {
    $retry = Invoke-GitStep 'fetch-retry' 'fetch'
    if (-not $retry.Ok) {
        Fail-Step 'Fetch' $(if ($retry.Output) { $retry.Output } else { 'git fetch retry failed' })
    }
    $pull = Invoke-PullOnce 'pull-retry'
    if (-not (Test-BehindGatePass)) {
        $why = if ($SkipPull) {
            "still behind $($result.behindCount) (pull skipped)"
        } elseif ($pull.Output) {
            "still behind $($result.behindCount): $($pull.Output)"
        } else {
            "still behind $($result.behindCount) after pull --ff-only"
        }
        Fail-Step 'Pull' $why
    }
}
Add-Step 'pull-gate' $true "behind=$($result.behindCount) ahead=$($result.aheadCount)"

if ($result.didStash) {
    $top = Invoke-GitRaw 'stash list -1'
    if (-not $top.Ok -or ($top.StdOut -notmatch [regex]::Escape($stashMsg))) {
        Fail-Step 'StashPop' "expected stash@{0} to be ours before pop: $($top.StdOut)"
    }
    $pop = Invoke-GitStep 'stash-pop' 'stash pop stash@{0}'
    if (-not $pop.Ok) {
        Fail-Step 'StashPop' $(if ($pop.Output) { $pop.Output } else { 'git stash pop failed' })
    }
    $stashListAfterPop = Get-StashListText
    if ($stashListAfterPop -match [regex]::Escape($stashMsg)) {
        Fail-Step 'StashPop' "stash message still present after pop: $stashMsg"
    }
    Add-Step 'stash-pop-gate' $true 'stash restored'
}

$result.ok = $true
Write-Result 0
