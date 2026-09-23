#Requires -Version 5.1
<#
.SYNOPSIS
  Elevated helper for Anti-Cat: disable notebook keyboard + trackpad only.

  Whitelist: Lenovo PS/2 KB (ACPI\IDEA0103) + Synaptics I2C trackpad (SYNA2BA6).
  Never disables any Bluetooth keyboard/mouse/HID (BTH*, BTHLE*, HOGP, class Bluetooth).
  IPC via -StateDir files: status, cmd, heartbeat, disabled_ids, pid.
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$StateDir
)

$ErrorActionPreference = "Stop"

function Test-IsAdmin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $p = New-Object Security.Principal.WindowsPrincipal($id)
    return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-IsAdmin)) {
    if (-not (Test-Path -LiteralPath $StateDir)) {
        New-Item -ItemType Directory -Path $StateDir -Force | Out-Null
    }
    Set-Content -LiteralPath (Join-Path $StateDir "status") -Value "awaiting-uac" -Encoding utf8
    $argList = @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", "`"$PSCommandPath`"",
        "-StateDir", "`"$StateDir`""
    ) -join " "
    try {
        Start-Process -FilePath "powershell.exe" -Verb RunAs -ArgumentList $argList | Out-Null
    } catch {
        Set-Content -LiteralPath (Join-Path $StateDir "status") -Value "error:UAC cancelled or elevation failed" -Encoding utf8
    }
    exit 0
}

if (-not (Test-Path -LiteralPath $StateDir)) {
    New-Item -ItemType Directory -Path $StateDir -Force | Out-Null
}

$statusPath = Join-Path $StateDir "status"
$cmdPath = Join-Path $StateDir "cmd"
$heartbeatPath = Join-Path $StateDir "heartbeat"
$idsPath = Join-Path $StateDir "disabled_ids"
$pidPath = Join-Path $StateDir "pid"

Set-Content -LiteralPath $pidPath -Value $PID -Encoding utf8
Set-Content -LiteralPath $statusPath -Value "starting" -Encoding utf8
if (Test-Path -LiteralPath $cmdPath) {
    Remove-Item -LiteralPath $cmdPath -Force -ErrorAction SilentlyContinue
}

function Test-IsBluetoothRelatedDevice($dev) {
    if (-not $dev) { return $true }
    $id = [string]$dev.InstanceId
    $name = [string]$dev.FriendlyName
    $cls = [string]$dev.Class
    if ($cls -eq 'Bluetooth') { return $true }
    # Bus / stack / HOGP (HID over GATT) — any Bluetooth keyboard or pointing device
    if ($id -match '(?i)^(BTH|BTHLE|BTHLEDEVICE)\\') { return $true }
    if ($id -match '(?i)\\BTH(LE)?(DEVICE)?\\') { return $true }
    if ($id -match '(?i)\{00001812-0000-1000-8000-00805F9B34FB\}') { return $true } # HOGP
    if ($id -match '(?i)BTHLEDEVICE') { return $true }
    if ($name -match '(?i)Bluetooth|Eyelash') { return $true }
    return $false
}

# Notebook built-in whitelist only (Lenovo PS/2 KB + Synaptics I2C trackpad).
# Hard reject: every Bluetooth keyboard and any other Bluetooth HID.
function Get-AntiCatTargets {
    Get-PnpDevice | Where-Object {
        $id = $_.InstanceId
        if (-not $id) { return $false }
        if (Test-IsBluetoothRelatedDevice($_)) { return $false }
        return (
            ($id -like 'ACPI\IDEA0103\*') -or
            ($id -like 'ACPI\SYNA2BA6*') -or
            ($id -like 'HID\SYNA2BA6*')
        )
    }
}

function Write-Status([string]$text) {
    Set-Content -LiteralPath $statusPath -Value $text -Encoding utf8
}

$script:DisabledIds = @()

function Disable-Targets {
    $script:DisabledIds = @()
    # Children before parents (HID\SYNA* before ACPI\SYNA*, then PS/2 KB).
    $targets = @(Get-AntiCatTargets | Where-Object { $_.Status -eq 'OK' } | Sort-Object {
            if ($_.InstanceId -like 'HID\SYNA2BA6*') { 0 }
            elseif ($_.InstanceId -like 'ACPI\SYNA2BA6*') { 1 }
            else { 2 }
        }, InstanceId)
    foreach ($dev in $targets) {
        if (Test-IsBluetoothRelatedDevice($dev)) { continue }
        try {
            Disable-PnpDevice -InstanceId $dev.InstanceId -Confirm:$false -ErrorAction Stop
            $script:DisabledIds += $dev.InstanceId
        } catch {
            # Continue; partial lock is still useful
        }
    }
    Set-Content -LiteralPath $idsPath -Value ($script:DisabledIds -join "`n") -Encoding utf8
    if ($script:DisabledIds.Count -eq 0) {
        Write-Status "error:no notebook keyboard/trackpad devices disabled (already off or access denied)"
        return $false
    }
    Write-Status "on"
    return $true
}

function Enable-DisabledTargets {
    $ids = @()
    if ($script:DisabledIds.Count -gt 0) {
        $ids = $script:DisabledIds
    } elseif (Test-Path -LiteralPath $idsPath) {
        $ids = @(Get-Content -LiteralPath $idsPath -Encoding utf8 | Where-Object { $_ -and $_.Trim() -ne "" })
    }
    # Parents before children when re-enabling
    $ids = @($ids | Sort-Object {
            if ($_ -like 'ACPI\*') { 0 }
            else { 1 }
        }, { $_ })
    foreach ($id in $ids) {
        try {
            Enable-PnpDevice -InstanceId $id.Trim() -Confirm:$false -ErrorAction SilentlyContinue
        } catch {
        }
    }
    # Also re-enable any matching targets still Error (best-effort restore)
    Get-AntiCatTargets | Where-Object { $_.Status -ne 'OK' } | ForEach-Object {
        try {
            Enable-PnpDevice -InstanceId $_.InstanceId -Confirm:$false -ErrorAction SilentlyContinue
        } catch {
        }
    }
    $script:DisabledIds = @()
    if (Test-Path -LiteralPath $idsPath) {
        Remove-Item -LiteralPath $idsPath -Force -ErrorAction SilentlyContinue
    }
    Write-Status "off"
}

$locked = $false
try {
    $locked = Disable-Targets
    if (-not $locked) {
        exit 1
    }

    $heartbeatTimeoutMs = 20000
    while ($true) {
        Start-Sleep -Milliseconds 250

        if (Test-Path -LiteralPath $cmdPath) {
            $cmd = (Get-Content -LiteralPath $cmdPath -Raw -Encoding utf8).Trim().ToLowerInvariant()
            Remove-Item -LiteralPath $cmdPath -Force -ErrorAction SilentlyContinue
            if ($cmd -eq "unlock") {
                break
            }
        }

        if (Test-Path -LiteralPath $heartbeatPath) {
            try {
                $age = ([DateTime]::UtcNow - (Get-Item -LiteralPath $heartbeatPath).LastWriteTimeUtc).TotalMilliseconds
                if ($age -gt $heartbeatTimeoutMs) {
                    break
                }
            } catch {
            }
        } else {
            # No heartbeat yet after lock — allow a grace window from status write
            try {
                $age = ([DateTime]::UtcNow - (Get-Item -LiteralPath $statusPath).LastWriteTimeUtc).TotalMilliseconds
                if ($age -gt $heartbeatTimeoutMs) {
                    break
                }
            } catch {
                break
            }
        }
    }
} finally {
    if ($locked -or (Test-Path -LiteralPath $idsPath)) {
        Enable-DisabledTargets
    } else {
        Write-Status "off"
    }
    if (Test-Path -LiteralPath $pidPath) {
        Remove-Item -LiteralPath $pidPath -Force -ErrorAction SilentlyContinue
    }
}
