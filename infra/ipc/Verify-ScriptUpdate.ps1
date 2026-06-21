#Requires -Version 5.1
<#
.SYNOPSIS
    Verification layer for Quick Update Scripts: confirms every target script
    exists, is non-empty, and is readable before the update process finalizes.
.EXAMPLE
    .\Verify-ScriptUpdate.ps1 -ScriptsDir "C:\Users\fie7ca\Documents\scripts" -PathsFile "C:\Users\fie7ca\AppData\Local\Temp\quick-update-paths.txt"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ScriptsDir,

    [Parameter(Mandatory)]
    [string]$PathsFile,

    [Parameter()]
    [string]$ReportFile = ""
)

$ErrorActionPreference = 'Stop'
$script:Failed = @()
if ($ReportFile) {
    $script:ReportFile = $ReportFile
} else {
    $script:ReportFile = Join-Path $env:TEMP "quick-update-verify-report_{0}.txt" -f [Guid]::NewGuid().ToString("N").Substring(0, 8)
}

function Write-VerifyReport {
    if ($script:Failed.Count -eq 0) { return }
    $script:Failed | Set-Content -Path $script:ReportFile -Encoding UTF8
}

try {
    if (-not (Test-Path -LiteralPath $PathsFile -PathType Leaf)) {
        Write-Error "Paths file not found: $PathsFile"
        exit 2
    }

    $paths = Get-Content -LiteralPath $PathsFile -Encoding UTF8 | Where-Object { $_.Trim() -ne '' }
    if (-not $paths) {
        Write-Error "No script paths in file: $PathsFile"
        exit 2
    }

    foreach ($path in $paths) {
        $path = $path.Trim()
        $name = Split-Path -Leaf $path

        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            $script:Failed += "$name (not found)"
            continue
        }

        $item = Get-Item -LiteralPath $path -Force
        if ($null -eq $item -or $item.Length -le 0) {
            $script:Failed += "$name (empty file)"
            continue
        }

        try {
            $null = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
        } catch {
            $script:Failed += "$name (unreadable: $($_.Exception.Message))"
            continue
        }
    }

    Write-VerifyReport

    if ($script:Failed.Count -gt 0) {
        exit 1
    }
    exit 0
} catch {
    $script:Failed += "Verification error: $($_.Exception.Message)"
    Write-VerifyReport
    exit 2
}
