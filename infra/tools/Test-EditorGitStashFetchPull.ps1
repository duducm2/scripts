# Regression harness for Editor-GitStashFetchPull.ps1 quality gates.
# Uses throwaway repos under %TEMP% — never the user's real working tree.
param(
    [string]$ScriptPath = (Join-Path $PSScriptRoot 'Editor-GitStashFetchPull.ps1')
)

$ErrorActionPreference = 'Stop'
$failed = 0
$passed = 0
$testRoot = Join-Path $env:TEMP ("editor-git-qg-" + [guid]::NewGuid().ToString('N'))

function Write-Case {
    param([string]$Name, [bool]$Ok, [string]$Detail = '')
    if ($Ok) {
        $script:passed++
        Write-Host "PASS  $Name"
    } else {
        $script:failed++
        Write-Host "FAIL  $Name  $Detail"
    }
}

function Set-GitIdentity {
    param([string]$Dir)
    git -C $Dir config user.email 'qg@test.local' | Out-Null
    git -C $Dir config user.name 'Quality Gate' | Out-Null
    git -C $Dir config commit.gpgsign false | Out-Null
    git -C $Dir config core.autocrlf false | Out-Null
    git -C $Dir config advice.detachedHead false | Out-Null
}

function New-TestRepos {
    param([string]$Name)
    $root = Join-Path $script:testRoot $Name
    New-Item -ItemType Directory -Path $root | Out-Null
    $bare = Join-Path $root 'upstream.git'
    $seed = Join-Path $root 'seed'
    $clone = Join-Path $root 'clone'
    $other = Join-Path $root 'other'
    git init --bare --initial-branch=main --quiet $bare
    git init --initial-branch=main --quiet $seed
    Set-GitIdentity $seed
    Set-Content -LiteralPath (Join-Path $seed 'README.md') -Value 'seed' -Encoding utf8
    git -C $seed add README.md
    git -C $seed commit -m 'init' --quiet
    git -C $seed remote add origin $bare
    git -C $seed push -u origin main --quiet
    git clone --quiet $bare $clone
    Set-GitIdentity $clone
    git -C $clone checkout main --quiet
    git clone --quiet $bare $other
    Set-GitIdentity $other
    return [pscustomobject]@{ Root = $root; Bare = $bare; Clone = $clone; Other = $other }
}

function Invoke-Flow {
    param(
        [string]$RepoDir,
        [string]$ResultPath,
        [switch]$SkipPull
    )
    $argLine = "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptPath`" -RepoDir `"$RepoDir`" -ResultPath `"$ResultPath`" -TimeoutSec 30"
    if ($SkipPull) { $argLine += ' -SkipPull' }
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = 'powershell.exe'
    $psi.Arguments = $argLine
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $psi
    $null = $p.Start()
    $p.WaitForExit()
    $json = $null
    if (Test-Path -LiteralPath $ResultPath) {
        $json = Get-Content -LiteralPath $ResultPath -Raw -Encoding UTF8 | ConvertFrom-Json
    }
    return [pscustomobject]@{ ExitCode = $p.ExitCode; Json = $json }
}

function Get-Head {
    param([string]$Dir)
    (git -C $Dir rev-parse HEAD).Trim()
}

function Get-Behind {
    param([string]$Dir)
    [int](git -C $Dir rev-list --count 'HEAD..@{upstream}').Trim()
}

try {
    New-Item -ItemType Directory -Path $testRoot | Out-Null
    if (-not (Test-Path -LiteralPath $ScriptPath)) {
        throw "Script not found: $ScriptPath"
    }

    # --- Clean clone, already in sync ---
    $r = New-TestRepos 'sync'
    $out = Invoke-Flow $r.Clone (Join-Path $r.Root 'result.json')
    $j = $out.Json
    Write-Case 'clean-sync-ok' ($out.ExitCode -eq 0 -and $j.ok -eq $true -and $j.behindCount -eq 0 -and $j.didStash -eq $false) (
        "exit=$($out.ExitCode) ok=$($j.ok) behind=$($j.behindCount) didStash=$($j.didStash) err=$($j.error)")

    # --- Dirty tracked file + new upstream commit ---
    $r = New-TestRepos 'dirty-pull-pop'
    Set-Content -LiteralPath (Join-Path $r.Clone 'README.md') -Value 'local-edit' -Encoding utf8
    Set-Content -LiteralPath (Join-Path $r.Other 'upstream.txt') -Value 'from-upstream' -Encoding utf8
    git -C $r.Other add upstream.txt
    git -C $r.Other commit -m 'upstream commit' --quiet
    git -C $r.Other push origin main --quiet
    $originBefore = (git -C $r.Clone rev-parse origin/main).Trim()
    $out = Invoke-Flow $r.Clone (Join-Path $r.Root 'result.json')
    $j = $out.Json
    $head = Get-Head $r.Clone
    $origin = (git -C $r.Clone rev-parse origin/main).Trim()
    $readme = Get-Content -LiteralPath (Join-Path $r.Clone 'README.md') -Raw
    $hasUp = Test-Path -LiteralPath (Join-Path $r.Clone 'upstream.txt')
    $behind = Get-Behind $r.Clone
    $stashCount = @(git -C $r.Clone stash list).Count
    Write-Case 'dirty-tracked-pull-pop' (
        $out.ExitCode -eq 0 -and $j.ok -eq $true -and $j.didStash -eq $true -and $behind -eq 0 `
        -and $head -eq $origin -and $origin -ne $originBefore -and $hasUp `
        -and ($readme -match 'local-edit') -and $stashCount -eq 0
    ) ("exit=$($out.ExitCode) ok=$($j.ok) didStash=$($j.didStash) behind=$behind headMatch=$($head -eq $origin) stash=$stashCount err=$($j.error)")

    # --- Tracked + untracked dirty: single stash, both restored ---
    $r = New-TestRepos 'double-dirty'
    Set-Content -LiteralPath (Join-Path $r.Clone 'README.md') -Value 'tracked-local' -Encoding utf8
    Set-Content -LiteralPath (Join-Path $r.Clone 'scratch.untracked') -Value 'untracked-local' -Encoding utf8
    Set-Content -LiteralPath (Join-Path $r.Other 'other.txt') -Value 'other' -Encoding utf8
    git -C $r.Other add other.txt
    git -C $r.Other commit -m 'other' --quiet
    git -C $r.Other push origin main --quiet
    $out = Invoke-Flow $r.Clone (Join-Path $r.Root 'result.json')
    $j = $out.Json
    $readme = Get-Content -LiteralPath (Join-Path $r.Clone 'README.md') -Raw
    $scratch = if (Test-Path -LiteralPath (Join-Path $r.Clone 'scratch.untracked')) {
        Get-Content -LiteralPath (Join-Path $r.Clone 'scratch.untracked') -Raw
    } else { '' }
    $stashCount = @(git -C $r.Clone stash list).Count
    $stashSteps = @($j.steps | Where-Object { $_.name -eq 'stash' })
    Write-Case 'tracked-and-untracked-one-stash' (
        $out.ExitCode -eq 0 -and $j.ok -eq $true -and $j.didStash -eq $true `
        -and ($readme -match 'tracked-local') -and ($scratch -match 'untracked-local') `
        -and $stashCount -eq 0 -and $stashSteps.Count -eq 1 -and (Get-Behind $r.Clone) -eq 0
    ) ("ok=$($j.ok) stashSteps=$($stashSteps.Count) stashLeft=$stashCount err=$($j.error)")

    # --- Still-behind after mocked/no-op pull ---
    $r = New-TestRepos 'skip-pull'
    Set-Content -LiteralPath (Join-Path $r.Other 'ahead.txt') -Value 'ahead' -Encoding utf8
    git -C $r.Other add ahead.txt
    git -C $r.Other commit -m 'ahead' --quiet
    git -C $r.Other push origin main --quiet
    $out = Invoke-Flow $r.Clone (Join-Path $r.Root 'result.json') -SkipPull
    $j = $out.Json
    Write-Case 'behind-gate-catches-skip-pull' (
        $out.ExitCode -ne 0 -and $j.ok -eq $false -and $j.failedStep -eq 'Pull' -and $j.behindCount -ge 1
    ) ("exit=$($out.ExitCode) ok=$($j.ok) step=$($j.failedStep) behind=$($j.behindCount) err=$($j.error)")

    # --- Stash pop conflict ---
    $r = New-TestRepos 'pop-conflict'
    Set-Content -LiteralPath (Join-Path $r.Clone 'README.md') -Value 'stashed-version' -Encoding utf8
    Set-Content -LiteralPath (Join-Path $r.Other 'README.md') -Value 'incoming-version' -Encoding utf8
    git -C $r.Other add README.md
    git -C $r.Other commit -m 'incoming' --quiet
    git -C $r.Other push origin main --quiet
    $out = Invoke-Flow $r.Clone (Join-Path $r.Root 'result.json')
    $j = $out.Json
    Write-Case 'stash-pop-conflict-fails' (
        $out.ExitCode -ne 0 -and $j.ok -eq $false -and $j.failedStep -eq 'StashPop'
    ) ("exit=$($out.ExitCode) ok=$($j.ok) step=$($j.failedStep) err=$($j.error)")

    # --- No upstream ---
    $r = New-TestRepos 'no-upstream'
    git -C $r.Clone checkout -b local-only --quiet
    $out = Invoke-Flow $r.Clone (Join-Path $r.Root 'result.json')
    $j = $out.Json
    Write-Case 'no-upstream-fails' (
        $out.ExitCode -ne 0 -and $j.ok -eq $false -and $j.failedStep -eq 'Upstream'
    ) ("exit=$($out.ExitCode) ok=$($j.ok) step=$($j.failedStep) err=$($j.error)")

    # --- Detached HEAD ---
    $r = New-TestRepos 'detached'
    git -C $r.Clone checkout --detach HEAD --quiet
    $out = Invoke-Flow $r.Clone (Join-Path $r.Root 'result.json')
    $j = $out.Json
    Write-Case 'detached-head-fails' (
        $out.ExitCode -ne 0 -and $j.ok -eq $false -and $j.failedStep -eq 'Detached'
    ) ("exit=$($out.ExitCode) ok=$($j.ok) step=$($j.failedStep) err=$($j.error)")

    # --- Merge in progress ---
    $r = New-TestRepos 'merge'
    git -C $r.Clone checkout -b other --quiet
    Set-Content -LiteralPath (Join-Path $r.Clone 'README.md') -Value 'other-side' -Encoding utf8
    git -C $r.Clone add README.md
    git -C $r.Clone commit -m 'other' --quiet
    git -C $r.Clone checkout main --quiet
    Set-Content -LiteralPath (Join-Path $r.Clone 'README.md') -Value 'main-side' -Encoding utf8
    git -C $r.Clone add README.md
    git -C $r.Clone commit -m 'main' --quiet
    git -C $r.Clone merge other --no-edit 2>$null
    $out = Invoke-Flow $r.Clone (Join-Path $r.Root 'result.json')
    $j = $out.Json
    Write-Case 'merge-in-progress-fails' (
        $out.ExitCode -ne 0 -and $j.ok -eq $false -and $j.failedStep -eq 'Merge'
    ) ("exit=$($out.ExitCode) ok=$($j.ok) step=$($j.failedStep) err=$($j.error)")
}
catch {
    $failed++
    Write-Host "FAIL  harness  $($_.Exception.Message)"
}
finally {
    if (Test-Path -LiteralPath $testRoot) {
        try { Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue } catch {}
    }
}

Write-Host ""
Write-Host "Passed: $passed  Failed: $failed"
if ($failed -gt 0) { exit 1 }
exit 0
