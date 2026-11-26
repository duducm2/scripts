# Alt+S editor git flow: stash (single -u), fetch, ff-only pull, stash pop.
# Success (ok=true / exit 0) only if every quality gate passes — never on git exit 0 alone.
# Writes JSON result to -ResultPath (tests) or plain ok/fail with -Robot (Alt+S terminal).
# -SkipPull is test-only: fetch then skip pull so the behind-count gate must fail.
# -Robot prints === ROBOT === banners and git output in the editor terminal.
param(
    [Parameter(Mandatory = $true)]
    [string]$RepoDir,
    [Parameter(Mandatory = $true)]
    [string]$ResultPath,
    [int]$TimeoutSec = 45,
    [switch]$SkipPull,
    [switch]$Robot
)

$ErrorActionPreference = 'Continue'
$env:GIT_TERMINAL_PROMPT = '0'
$env:GCM_INTERACTIVE = 'Never'
$env:GIT_PAGER = 'cat'

$stashMsg = "alt+s-auto-$(Get-Date -Format 'yyyyMMddHHmmss')-$PID"

$result = [ordered]@{
    ok                   = $false
    didStash             = $false
    pulledUpdated        = $false
    stashRestoredOnError = $false
    stashPopWarning      = $false
    failedStep           = ''
    error                = ''
    behindCount          = -1
    aheadCount           = -1
    upstream             = ''
    stashRef             = ''
    steps                = @()
}

$script:didStashNeedsRestore = $false

function Write-Robot {
    param([string]$Msg, [string]$Color = 'Cyan')
    if ($Robot) {
        Write-Host "=== ROBOT $Msg ===" -ForegroundColor $Color
    }
}

function Write-Result {
    param([int]$ExitCode = 1)
    try {
        if ($Robot) {
            $plain = if (($ExitCode -eq 0) -and $result.ok) { 'ok' } else { 'fail' }
            [IO.File]::WriteAllText($ResultPath, $plain)
            if ($plain -eq 'ok') {
                Write-Host '=== ROBOT done ===' -ForegroundColor Green
            } else {
                Write-Host ('=== ROBOT failed === ' + $result.failedStep + ': ' + $result.error) -ForegroundColor Red
            }
        } else {
            $json = $result | ConvertTo-Json -Compress -Depth 6
            Set-Content -LiteralPath $ResultPath -Value $json -Encoding UTF8
        }
    } catch {
        exit $ExitCode
    }
    exit $ExitCode
}

function Add-Step {
    param([string]$Name, [bool]$Ok, [string]$Detail = '')
    $result.steps += [ordered]@{ name = $Name; ok = $Ok; detail = $Detail }
}

function Find-OurStashRef {
    $r = Invoke-GitRaw 'stash list'
    if (-not $r.Ok -or -not $r.StdOut) { return '' }
    foreach ($line in ($r.StdOut -split "`r?`n")) {
        if ($line -match [regex]::Escape($stashMsg) -and $line -match '^(stash@\{\d+\}):') {
            return $Matches[1]
        }
    }
    return ''
}

function Restore-StashOnFailure {
    if (-not $script:didStashNeedsRestore) { return }
    $script:didStashNeedsRestore = $false
    Write-Robot 'restore stash after error' 'Yellow'
    try {
        $dirRes = Invoke-GitRaw 'rev-parse --absolute-git-dir'
        if ($dirRes.Ok -and $dirRes.StdOut) {
            $lock = Join-Path $dirRes.StdOut.Trim() 'index.lock'
            $dl = [datetime]::UtcNow.AddSeconds(15)
            while ((Test-Path -LiteralPath $lock) -and ([datetime]::UtcNow -lt $dl)) {
                Start-Sleep -Milliseconds 200
            }
        }
        $ref = Find-OurStashRef
        if (-not $ref) {
            Add-Step 'stash-restore-on-error' $false "stash not found for $stashMsg"
            $result.stashPopWarning = $true
            return
        }
        $pop = Invoke-GitRaw "stash pop $ref"
        if (-not $pop.Ok) {
            # If Google Drive left a tracked file dirty during stash push, reset tracked files to HEAD and retry pop.
            Invoke-GitRaw 'checkout -f HEAD' | Out-Null
            $pop = Invoke-GitRaw "stash pop $ref"
        }
        if ($pop.Ok) {
            $result.stashRestoredOnError = $true
            Add-Step 'stash-restore-on-error' $true "restored $ref"
        } else {
            $result.stashPopWarning = $true
            Add-Step 'stash-restore-on-error' $false $(if ($pop.Output) { $pop.Output } else { "git stash pop $ref failed" })
        }
    } catch {
        $result.stashPopWarning = $true
        Add-Step 'stash-restore-on-error' $false $_.Exception.Message
    }
}

function Fail-Step {
    param([string]$Step, [string]$ErrorText)
    if ($Step -ne 'StashPop') {
        Restore-StashOnFailure
    }
    $result.failedStep = $Step
    $result.error = $ErrorText
    if ($Step -eq 'StashPop') {
        $result.stashPopWarning = $true
    }
    Add-Step $Step $false $ErrorText
    Write-Result 1
}

trap {
    Fail-Step 'Unexpected' $_.Exception.Message
}

function Invoke-GitRaw {
    param([string]$GitArgs)
    if ($Robot) {
        Write-Host ("git " + $GitArgs) -ForegroundColor DarkGray
        $output = cmd.exe /c "git -c core.fsmonitor=false -C `"$RepoDir`" $GitArgs 2>&1"
        $code = $LASTEXITCODE
        $text = if ($null -eq $output) { '' } else { ($output | Out-String).Trim() }
        return [pscustomobject]@{
            Ok       = ($code -eq 0)
            ExitCode = $code
            Output   = $text
            StdOut   = $text
        }
    }
    $id = [guid]::NewGuid().ToString('N')
    $outFile = Join-Path $env:TEMP "git-raw-out-$PID-$id.txt"
    $errFile = Join-Path $env:TEMP "git-raw-err-$PID-$id.txt"
    try {
        $cmdLine = "/c `"git -c core.fsmonitor=false -C `"$RepoDir`" $GitArgs 1>`"$outFile`" 2>`"$errFile`"`""
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = 'cmd.exe'
        $psi.Arguments = $cmdLine
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        $psi.RedirectStandardOutput = $false
        $psi.RedirectStandardError = $false
        $psi.RedirectStandardInput = $false
        $p = New-Object System.Diagnostics.Process
        $p.StartInfo = $psi
        $null = $p.Start()
        $ms = [Math]::Max(1000, $TimeoutSec * 1000)
        if (-not $p.WaitForExit($ms)) {
            try { $p.Kill() } catch {}
            return [pscustomobject]@{ Ok = $false; ExitCode = 124; Output = "Timed out after ${TimeoutSec}s"; StdOut = '' }
        }
        $p.WaitForExit()
        $code = $p.ExitCode
        $stdout = if (Test-Path -LiteralPath $outFile) { [IO.File]::ReadAllText($outFile, [Text.Encoding]::UTF8) } else { '' }
        $stderr = if (Test-Path -LiteralPath $errFile) { [IO.File]::ReadAllText($errFile, [Text.Encoding]::UTF8) } else { '' }
        $combined = ($stdout + "`n" + $stderr).Trim()
        return [pscustomobject]@{
            Ok       = ($code -eq 0)
            ExitCode = $code
            Output   = $combined
            StdOut   = $stdout.Trim()
        }
    } finally {
        try { if (Test-Path -LiteralPath $outFile) { Remove-Item -LiteralPath $outFile -Force -ErrorAction SilentlyContinue } } catch {}
        try { if (Test-Path -LiteralPath $errFile) { Remove-Item -LiteralPath $errFile -Force -ErrorAction SilentlyContinue } } catch {}
    }
}

function Invoke-GitStep {
    param([string]$StepName, [string]$GitArgs)
    $r = Invoke-GitRaw $GitArgs
    $show = $Robot -and $r.Output -and ($StepName -match '^(stash|fetch|pull)')
    if ($show) {
        Write-Host $r.Output
    }
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
    $ref = Invoke-GitRaw 'update-index --refresh'
    if ($ref.ExitCode -eq 124) {
        Fail-Step 'Status' $ref.Output
    }
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
    param([int]$WaitSec = 15)
    $deadline = [datetime]::UtcNow.AddSeconds($WaitSec)
    $last = ''
    $triedCheckoutReset = $false
    while ($true) {
        Wait-IndexLock
        $last = Get-Porcelain
        if (-not (Test-HasStashableChanges)) {
            return $last
        }
        if (-not $triedCheckoutReset -and $script:didStashNeedsRestore) {
            $triedCheckoutReset = $true
            Invoke-GitRaw 'checkout -f HEAD' | Out-Null
            Start-Sleep -Milliseconds 150
            continue
        }
        if ([datetime]::UtcNow -ge $deadline) {
            Fail-Step 'Stash' "working tree still not pullable after stash: $last"
        }
        Start-Sleep -Milliseconds 300
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
    if ($behind.ExitCode -eq 124) {
        Fail-Step 'Status' $behind.Output
    }
    $ahead = Invoke-GitRaw 'rev-list --count @{upstream}..HEAD'
    if ($ahead.ExitCode -eq 124) {
        Fail-Step 'Status' $ahead.Output
    }
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

Write-Robot 'preflight'

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
    Write-Robot 'stash'
    $push = Invoke-GitStep 'stash' "stash push -u -m `"$stashMsg`""
    if ($push.Output -match '(?i)No local changes to save') {
        Write-Robot 'stash skip' 'Yellow'
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
    $script:didStashNeedsRestore = $true
    $result.stashRef = 'stash@{0}'
    Wait-WorkingTreePullable | Out-Null
    Add-Step 'stash-gate' $true $stashMsg
    }
} else {
    Write-Robot 'stash skip' 'Yellow'
    Add-Step 'stash-skip' $true 'clean working tree'
}

Write-Robot 'fetch'
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

$headBeforePull = (Invoke-GitRaw 'rev-parse HEAD').StdOut

Write-Robot 'pull'
Wait-IndexLock
$pull = Invoke-PullOnce 'pull'
if ((-not $pull.Ok) -or (-not (Test-BehindGatePass))) {
    $retry = Invoke-GitStep 'fetch-retry' 'fetch'
    if (-not $retry.Ok) {
        Fail-Step 'Fetch' $(if ($retry.Output) { $retry.Output } else { 'git fetch retry failed' })
    }
    Wait-IndexLock
    $pull = Invoke-PullOnce 'pull-retry'
    if (((-not $pull.Ok) -or (-not (Test-BehindGatePass))) -and (-not $SkipPull) -and ($result.behindCount -gt 0) -and ($result.aheadCount -gt 0)) {
        Write-Robot 'pull --rebase (diverged)' 'Yellow'
        $pull = Invoke-GitStep 'pull-rebase' 'pull --rebase'
        if (-not $pull.Ok) {
            Invoke-GitRaw 'rebase --abort' | Out-Null
        }
    }
    if ((-not $pull.Ok) -or (-not (Test-BehindGatePass))) {
        $why = if ($SkipPull) {
            "still behind $($result.behindCount) (pull skipped)"
        } elseif ($pull.Output) {
            "pull failed (behind=$($result.behindCount)): $($pull.Output)"
        } else {
            "pull --ff-only failed (behind=$($result.behindCount))"
        }
        Fail-Step 'Pull' $why
    }
}
$headAfterPull = (Invoke-GitRaw 'rev-parse HEAD').StdOut
$result.pulledUpdated = [bool]($headBeforePull -and $headAfterPull -and ($headBeforePull -ne $headAfterPull))
Add-Step 'pull-gate' $true "behind=$($result.behindCount) ahead=$($result.aheadCount) updated=$($result.pulledUpdated)"

if ($result.didStash) {
    Write-Robot 'stash pop'
    $script:didStashNeedsRestore = $false
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
