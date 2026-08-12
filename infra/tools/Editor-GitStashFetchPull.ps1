# Alt+S editor git flow: stash (with untracked fallback), fetch, pull, stash pop.
# Writes JSON result to -ResultPath; exit 0 on success, 1 on failure.
param(
    [Parameter(Mandatory = $true)]
    [string]$RepoDir,
    [Parameter(Mandatory = $true)]
    [string]$ResultPath,
    [int]$TimeoutSec = 120
)

$ErrorActionPreference = 'Continue'
$env:GIT_TERMINAL_PROMPT = '0'
$env:GCM_INTERACTIVE = 'Never'

$result = [ordered]@{
    ok              = $false
    didStash        = $false
    stashPopWarning = $false
    failedStep      = ''
    error           = ''
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

function Test-DirtyTreeMessage {
    param([string]$Text)
    return ($Text -match '(?i)clean your repository working tree|commit your changes or stash|would be overwritten by merge')
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
    }
}

function Test-GitStashOutputOk {
    param([string]$Output, [int]$ExitCode)
    if ($ExitCode -eq 0) { return $true }
    return ($Output -match '(?i)Saved working directory|No local changes to save')
}

function Invoke-GitStep {
    param(
        [string]$StepName,
        [string]$GitArgs,
        [switch]$AllowStashPartialOk
    )
    $r = Invoke-GitRaw $GitArgs
    $ok = $r.Ok
    if ($AllowStashPartialOk -and (Test-GitStashOutputOk $r.Output $r.ExitCode)) {
        $ok = $true
    }
    $detail = if ($r.Output) { $r.Output.Substring(0, [Math]::Min(500, $r.Output.Length)) } else { '' }
    Add-Step $StepName $ok $detail
    if (-not $ok -and $r.ExitCode -eq 124) {
        $result.failedStep = $StepName
        $result.error = $r.Output
        Write-Result 1
    }
    $r | Add-Member -NotePropertyName StepOk -NotePropertyValue $ok -Force
    return $r
}

function Test-WorkingTreeDirty {
    $r = Invoke-GitRaw 'status --porcelain'
    if (-not $r.Ok) {
        $result.failedStep = 'Status'
        $result.error = $r.Output
        Write-Result 1
    }
    return ($r.Output -match '\S')
}

function Invoke-StashIfNeeded {
    param([switch]$IncludeUntracked)
    if (-not (Test-WorkingTreeDirty)) {
        Add-Step $(if ($IncludeUntracked) { 'stash-untracked-skip' } else { 'stash-skip' }) $true 'clean working tree'
        return $false
    }
    $msg = if ($IncludeUntracked) { 'alt+s untracked' } else { 'alt+s auto' }
    $args = if ($IncludeUntracked) { "stash push -u -m `"$msg`"" } else { "stash push -m `"$msg`"" }
    $r = Invoke-GitStep $(if ($IncludeUntracked) { 'stash-untracked' } else { 'stash' }) $args -AllowStashPartialOk
    if (-not $r.StepOk) {
        $result.failedStep = $(if ($IncludeUntracked) { 'Stash (untracked)' } else { 'Stash' })
        $result.error = $r.Output
        Write-Result 1
    }
    $result.didStash = $true
    return $true
}

if (-not (Test-Path -LiteralPath $RepoDir)) {
    $result.failedStep = 'Repo'
    $result.error = "Directory not found: $RepoDir"
    Write-Result 1
}

$top = Invoke-GitStep 'rev-parse' 'rev-parse --show-toplevel'
if (-not $top.Ok) {
    $result.failedStep = 'Repo'
    $result.error = if ($top.Output) { $top.Output } else { 'Not a git repository' }
    Write-Result 1
}

Invoke-StashIfNeeded | Out-Null
if (Test-WorkingTreeDirty) {
    Invoke-StashIfNeeded -IncludeUntracked | Out-Null
}

$fetch = Invoke-GitStep 'fetch' 'fetch'
if (-not $fetch.Ok) {
    $result.failedStep = 'Fetch'
    $result.error = $fetch.Output
    Write-Result 1
}

function Invoke-PullWithRecovery {
    $pull = Invoke-GitStep 'pull' 'pull'
    if ($pull.Ok) { return $pull }
    if (Test-DirtyTreeMessage $pull.Output) {
        Invoke-StashIfNeeded -IncludeUntracked | Out-Null
        $retry = Invoke-GitStep 'pull-retry' 'pull'
        if ($retry.Ok) { return $retry }
        $result.failedStep = 'Pull'
        $result.error = $retry.Output
        Write-Result 1
    }
    $result.failedStep = 'Pull'
    $result.error = $pull.Output
    Write-Result 1
}

Invoke-PullWithRecovery | Out-Null

if ($result.didStash) {
    $pop = Invoke-GitStep 'stash-pop' 'stash pop' -AllowStashPartialOk
    if (-not $pop.StepOk) {
        $result.stashPopWarning = $true
        Add-Step 'stash-pop-warning' $false $pop.Output
    }
}

$result.ok = $true
Write-Result 0
