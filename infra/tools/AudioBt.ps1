param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("list", "default", "enable", "disable", "connect", "disconnect", "isolate")]
    [string]$Action,

    [string]$Id = "",

    [string]$OutFile = ""
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$csPath = Join-Path $PSScriptRoot "AudioBtHelper.cs"
$dllPath = Join-Path $PSScriptRoot "AudioBtHelper.dll"

function Write-AudioBtResult([string]$text, [int]$code) {
    $out = $text.TrimEnd()
    if ($OutFile -ne "") {
        $utf8 = New-Object System.Text.UTF8Encoding $false
        [System.IO.File]::WriteAllText($OutFile, $out + [Environment]::NewLine, $utf8)
    } else {
        Write-Output $out
    }
    exit $code
}

function Import-AudioBtHelper {
    if (-not (Test-Path -LiteralPath $csPath)) {
        throw "Missing AudioBtHelper.cs"
    }
    $csTime = (Get-Item -LiteralPath $csPath).LastWriteTimeUtc
    $loaded = $false
    if (Test-Path -LiteralPath $dllPath) {
        $dllTime = (Get-Item -LiteralPath $dllPath).LastWriteTimeUtc
        if ($dllTime -ge $csTime) {
            try {
                Add-Type -Path $dllPath
                $loaded = $true
            } catch {
                $loaded = $false
            }
        }
    }
    if ($loaded) {
        return
    }
    $code = Get-Content -LiteralPath $csPath -Raw -Encoding UTF8
    try {
        Add-Type -TypeDefinition $code -OutputAssembly $dllPath -OutputType Library -Language CSharp
        Add-Type -Path $dllPath
    } catch {
        Add-Type -TypeDefinition $code -Language CSharp
    }
}

function Wait-AudioBtWinRT($async, [int]$timeoutMs = 15000) {
    if ($null -eq $async) {
        return $null
    }
    $sw = [Diagnostics.Stopwatch]::StartNew()
    while ($true) {
        $st = [string]$async.Status
        if ($st -eq "Completed") {
            try {
                return $async.GetResults()
            } catch {
                return $null
            }
        }
        if ($st -eq "Error" -or $st -eq "Canceled") {
            $code = ""
            try { $code = [string]$async.ErrorCode } catch { }
            throw "WinRT $st $code"
        }
        if ($sw.ElapsedMilliseconds -ge $timeoutMs) {
            throw "WinRT async timeout"
        }
        Start-Sleep -Milliseconds 40
    }
}

function Test-AudioBtMacMatch([string]$blob, [string]$hex) {
    if ([string]::IsNullOrWhiteSpace($blob) -or [string]::IsNullOrWhiteSpace($hex)) {
        return $false
    }
    $compact = $blob.ToUpperInvariant() -replace "[:\-_]", ""
    return $compact.Contains($hex.ToUpperInvariant())
}

function Invoke-AudioBtWinRTConnect([string]$Id) {
    $hex = [AudioBt.Helper]::BluetoothAddressHex($Id)
    if ([string]::IsNullOrWhiteSpace($hex)) {
        return @{ ok = $false; err = "No Bluetooth address for WinRT connect" }
    }
    $hex = $hex.ToUpperInvariant() -replace "[:\-_]", ""
    try {
        Add-Type -AssemblyName System.Runtime.WindowsRuntime -ErrorAction SilentlyContinue
        $null = [Windows.Media.Audio.AudioPlaybackConnection, Windows.Media.Audio, ContentType = WindowsRuntime]
        $null = [Windows.Devices.Enumeration.DeviceInformation, Windows.Devices.Enumeration, ContentType = WindowsRuntime]
    } catch {
        return @{ ok = $false; err = "AudioPlaybackConnection unavailable: $($_.Exception.Message)" }
    }

    try {
        $selector = [Windows.Media.Audio.AudioPlaybackConnection]::GetDeviceSelector()
        $devices = Wait-AudioBtWinRT ([Windows.Devices.Enumeration.DeviceInformation]::FindAllAsync($selector))
        $match = $null
        foreach ($d in $devices) {
            $blob = ""
            try { $blob = [string]$d.Id + " " + [string]$d.Name } catch { $blob = [string]$d.Id }
            if (Test-AudioBtMacMatch $blob $hex) {
                $match = $d
                break
            }
        }
        if ($null -eq $match) {
            return @{ ok = $false; err = "AudioPlaybackConnection: no device matching $hex" }
        }
        $conn = [Windows.Media.Audio.AudioPlaybackConnection]::TryCreateFromId($match.Id)
        if ($null -eq $conn) {
            return @{ ok = $false; err = "AudioPlaybackConnection.TryCreateFromId failed" }
        }
        Wait-AudioBtWinRT ($conn.StartAsync()) | Out-Null
        $open = Wait-AudioBtWinRT ($conn.OpenAsync())
        $status = ""
        if ($null -ne $open) {
            try { $status = [string]$open.Status } catch { $status = "" }
        }
        if ($status -ne "" -and $status -ne "0" -and $status -ne "Success") {
            return @{ ok = $false; err = "AudioPlaybackConnection.OpenAsync $status" }
        }
        return @{ ok = $true; err = "" }
    } catch {
        return @{ ok = $false; err = "AudioPlaybackConnection: $($_.Exception.Message)" }
    }
}

function Merge-AudioBtErr([string]$primary, [string]$winrtErr) {
    if ([string]::IsNullOrWhiteSpace($primary)) {
        $primary = "ERR`tBluetooth connect failed"
    }
    if ([string]::IsNullOrWhiteSpace($winrtErr)) {
        return $primary
    }
    if ($primary.StartsWith("ERR`t")) {
        return ($primary.TrimEnd() + "; WinRT: " + $winrtErr)
    }
    return "ERR`t$primary; WinRT: $winrtErr"
}

function Invoke-AudioBtConnectOrWinRT([string]$Id) {
    $r = [AudioBt.Helper]::Connect($Id)
    if ($r -and -not $r.StartsWith("ERR`t")) {
        return $r
    }
    $w = Invoke-AudioBtWinRTConnect $Id
    if ($w.ok) {
        $c = [AudioBt.Helper]::ConfirmConnected($Id, "AudioPlaybackConnection")
        if ($c) {
            return $c
        }
    }
    return (Merge-AudioBtErr $r $w.err)
}

function Invoke-AudioBtIsolateOrWinRT([string]$Id) {
    $r = [AudioBt.Helper]::Isolate($Id)
    if ($r -and -not $r.StartsWith("ERR`t")) {
        return $r
    }
    $w = Invoke-AudioBtWinRTConnect $Id
    if ($w.ok) {
        $c = [AudioBt.Helper]::ConfirmConnected($Id, "AudioPlaybackConnection")
        if ($c -and -not $c.StartsWith("ERR`t")) {
            $r2 = [AudioBt.Helper]::Isolate($Id)
            if ($r2) {
                return $r2
            }
        }
        if ($c) {
            return $c
        }
    }
    return (Merge-AudioBtErr $r $w.err)
}

try {
    Import-AudioBtHelper
} catch {
    Write-AudioBtResult ("ERR`tFailed to load AudioBt helper: " + $_.Exception.Message) 1
}

if ($null -eq $Id) { $Id = "" }
$Id = $Id.Trim().Trim("'").Trim('"')

if ($Action -ne "list" -and [string]::IsNullOrWhiteSpace($Id)) {
    Write-AudioBtResult "ERR`tMissing -Id" 1
}

$result = switch ($Action) {
    "list" { [AudioBt.Helper]::ListTsv() }
    "default" { [AudioBt.Helper]::SetDefault($Id) }
    "enable" { [AudioBt.Helper]::SetVisible($Id, $true) }
    "disable" { [AudioBt.Helper]::SetVisible($Id, $false) }
    "connect" { Invoke-AudioBtConnectOrWinRT $Id }
    "disconnect" { [AudioBt.Helper]::Disconnect($Id) }
    "isolate" { Invoke-AudioBtIsolateOrWinRT $Id }
}

if ($null -eq $result) {
    Write-AudioBtResult "ERR`tEmpty helper result" 1
}

$code = 0
if ($result.StartsWith("ERR`t")) {
    $code = 1
}
Write-AudioBtResult $result $code
