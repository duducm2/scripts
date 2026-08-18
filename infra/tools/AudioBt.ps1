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

function Get-AudioBtWinRTStatus($async) {
    try {
        return [string]$async.Status
    } catch {
        return "?"
    }
}

function Get-AudioBtTaskBlob($task, $async) {
    $ts = "?"
    $done = "?"
    $fault = "?"
    try { $ts = [string]$task.Status } catch { }
    try { $done = [string]$task.IsCompleted } catch { }
    try { $fault = [string]$task.IsFaulted } catch { }
    return "taskStatus=$ts completed=$done faulted=$fault winrtStatus=$(Get-AudioBtWinRTStatus $async)"
}

function Wait-AudioBtWinRT($async, [int]$timeoutMs = 4000, [type]$resultType = $null) {
    if ($null -eq $async) {
        return $null
    }
    Add-Type -AssemblyName System.Runtime.WindowsRuntime -ErrorAction Stop
    $task = $null
    foreach ($m in [System.WindowsRuntimeSystemExtensions].GetMethods()) {
        if ($m.Name -ne "AsTask") {
            continue
        }
        $ps = $m.GetParameters()
        if ($ps.Count -ne 1) {
            continue
        }
        $pn = $ps[0].ParameterType.Name
        if ($null -ne $resultType) {
            if (-not $m.IsGenericMethodDefinition) {
                continue
            }
            if ($m.GetGenericArguments().Count -ne 1) {
                continue
            }
            if ($pn -ne 'IAsyncOperation`1') {
                continue
            }
            $task = $m.MakeGenericMethod($resultType).Invoke($null, @($async))
            break
        }
        if ($m.IsGenericMethodDefinition) {
            continue
        }
        if ($pn -ne "IAsyncAction") {
            continue
        }
        $task = $m.Invoke($null, @($async))
        break
    }
    if ($null -eq $task) {
        throw "WinRT AsTask unavailable"
    }
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $finished = $false
    try {
        $finished = [bool]$task.Wait($timeoutMs)
    } catch {
        $inner = $_.Exception
        while ($null -ne $inner.InnerException) {
            $inner = $inner.InnerException
        }
        throw $inner
    }
    if (-not $finished) {
        try { $async.Cancel() } catch { }
        throw ("WinRT async timeout after " + $sw.ElapsedMilliseconds + "ms " + (Get-AudioBtTaskBlob $task $async))
    }
    if ($task.IsFaulted) {
        $ex = $task.Exception
        if ($null -ne $ex.InnerException) {
            $ex = $ex.InnerException
        }
        throw $ex
    }
    if ($null -eq $resultType) {
        return $null
    }
    return $task.Result
}

function Test-AudioBtNeedsConnectFallback([string]$result) {
    if ([string]::IsNullOrWhiteSpace($result)) {
        return $true
    }
    if ($result -match "SetDefault failed") {
        return $false
    }
    return $true
}

function Get-AudioBtConnectionStatus($dev) {
    try {
        return [string]$dev.ConnectionStatus
    } catch {
        return "?"
    }
}

function Invoke-AudioBtRfcommUncached($dev) {
    $resultType = [Windows.Devices.Bluetooth.Rfcomm.RfcommDeviceServicesResult]
    $uncached = [Windows.Devices.Bluetooth.BluetoothCacheMode]::Uncached
    try {
        return Wait-AudioBtWinRT ($dev.GetRfcommServicesAsync($uncached)) 8000 $resultType
    } catch {
        try {
            return Wait-AudioBtWinRT ($dev.GetRfcommServicesAsync()) 8000 $resultType
        } catch {
            return $null
        }
    }
}

function Invoke-AudioBtHfpStream($dev) {
    try {
        $hfp = [Windows.Devices.Bluetooth.Rfcomm.RfcommServiceId]::FromUuid([guid]"0000111e-0000-1000-8000-00805f9b34fb")
        $resultType = [Windows.Devices.Bluetooth.Rfcomm.RfcommDeviceServicesResult]
        $uncached = [Windows.Devices.Bluetooth.BluetoothCacheMode]::Uncached
        $result = $null
        try {
            $result = Wait-AudioBtWinRT ($dev.GetRfcommServicesForIdAsync($hfp, $uncached)) 8000 $resultType
        } catch {
            $result = Wait-AudioBtWinRT ($dev.GetRfcommServicesForIdAsync($hfp)) 8000 $resultType
        }
        $svc = $null
        foreach ($s in $result.Services) {
            $svc = $s
            break
        }
        if ($null -eq $svc) {
            return
        }
        $streamType = [Windows.Storage.Streams.IRandomAccessStream]
        $stream = $null
        try {
            $stream = Wait-AudioBtWinRT ($svc.GetStreamAsync()) 8000 $streamType
        } catch {
            try {
                $stream = Wait-AudioBtWinRT ($svc.GetStreamAsync()) 8000
            } catch {
                return
            }
        }
        if ($null -ne $stream) {
            try { $stream.Dispose() } catch { }
        }
    } catch {
    }
}

function Invoke-AudioBtAclConnect([string]$hex) {
    $addr = [uint64]::Parse($hex, [Globalization.NumberStyles]::HexNumber)
    $btType = [Windows.Devices.Bluetooth.BluetoothDevice]
    $dev = Wait-AudioBtWinRT ([Windows.Devices.Bluetooth.BluetoothDevice]::FromBluetoothAddressAsync($addr)) 4000 $btType
    if ($null -eq $dev) {
        return @{ ok = $false; err = "BluetoothDevice.FromBluetoothAddressAsync returned null"; status = "" }
    }
    $null = Invoke-AudioBtRfcommUncached $dev
    $afterRf = Get-AudioBtConnectionStatus $dev
    if ($afterRf -ne "Connected") {
        Invoke-AudioBtHfpStream $dev
        $afterRf = Get-AudioBtConnectionStatus $dev
    }
    $ok = ($afterRf -eq "Connected")
    return @{ ok = $ok; err = $(if ($ok) { "" } else { "BluetoothDevice still " + $afterRf }); status = $afterRf }
}

function Invoke-AudioBtWinRTConnect([string]$Id) {
    $hex = [AudioBt.Helper]::BluetoothAddressHex($Id)
    if ([string]::IsNullOrWhiteSpace($hex)) {
        return @{ ok = $false; err = "No Bluetooth address for WinRT connect" }
    }
    $hex = $hex.ToUpperInvariant() -replace "[:\-_]", ""
    try {
        Add-Type -AssemblyName System.Runtime.WindowsRuntime -ErrorAction SilentlyContinue
        $null = [Windows.Devices.Bluetooth.BluetoothDevice, Windows.Devices.Bluetooth, ContentType = WindowsRuntime]
        $null = [Windows.Devices.Bluetooth.BluetoothCacheMode, Windows.Devices.Bluetooth, ContentType = WindowsRuntime]
        $null = [Windows.Devices.Bluetooth.Rfcomm.RfcommDeviceService, Windows.Devices.Bluetooth.Rfcomm, ContentType = WindowsRuntime]
        $null = [Windows.Devices.Bluetooth.Rfcomm.RfcommServiceId, Windows.Devices.Bluetooth.Rfcomm, ContentType = WindowsRuntime]
        $null = [Windows.Devices.Bluetooth.Rfcomm.RfcommDeviceServicesResult, Windows.Devices.Bluetooth.Rfcomm, ContentType = WindowsRuntime]
        $null = [Windows.Storage.Streams.IRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
    } catch {
        return @{ ok = $false; err = "BluetoothDevice unavailable: $($_.Exception.Message)" }
    }

    try {
        $acl = Invoke-AudioBtAclConnect $hex
        if (-not $acl.ok) {
            return @{ ok = $false; err = $acl.err }
        }
        $null = [AudioBt.Helper]::EnableBtAudioServices($Id)
        return @{ ok = $true; err = "" }
    } catch {
        return @{ ok = $false; err = "BluetoothDevice ACL: $($_.Exception.Message)" }
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
        $c = [AudioBt.Helper]::ConfirmConnected($Id, "BluetoothDevice")
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
    if (-not (Test-AudioBtNeedsConnectFallback $r)) {
        return $r
    }
    $w = Invoke-AudioBtWinRTConnect $Id
    if ($w.ok) {
        $c = [AudioBt.Helper]::ConfirmConnected($Id, "BluetoothDevice")
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
