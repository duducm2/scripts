param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("list", "default", "enable", "disable", "connect", "disconnect", "isolate")]
    [string]$Action,

    [string]$Id = "",

    [string]$OutFile = "",

    [string]$LogDir = "",

    [string]$Note = ""
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$csPath = Join-Path $PSScriptRoot "AudioBtHelper.cs"
$dllPath = Join-Path $PSScriptRoot "AudioBtHelper.dll"

function Write-AudioBtPsLog([string]$msg) {
    try {
        [AudioBt.Helper]::LogLine($msg)
    } catch {
    }
}

function New-AudioBtSessionLog([string]$dir, [string]$action, [string]$id, [string]$note) {
    if ([string]::IsNullOrWhiteSpace($dir)) {
        return $null
    }
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $session = Join-Path $dir ("audio-bt-connect-" + $stamp + ".log")
    $index = Join-Path $dir "audio-bt-connect.log"
    $latest = Join-Path $dir "audio-bt-connect-latest.log"
    [AudioBt.Helper]::SetLogPath($session)
    Write-AudioBtPsLog "======== AudioBt session $stamp ========"
    Write-AudioBtPsLog ("computer=" + $env:COMPUTERNAME + " user=" + $env:USERNAME)
    Write-AudioBtPsLog ("os=" + [Environment]::OSVersion.VersionString)
    Write-AudioBtPsLog ("ps=" + $PSVersionTable.PSVersion.ToString() + " sta=" + [threading.thread]::CurrentThread.GetApartmentState())
    Write-AudioBtPsLog ("action=" + $action + " id=" + $id + " note=" + $note)
    [AudioBt.Helper]::LogSnapshot("before $action")
    return @{ session = $session; index = $index; latest = $latest; stamp = $stamp }
}

function Complete-AudioBtSessionLog($ctx, [string]$result, [int]$code) {
    if ($null -eq $ctx) {
        return
    }
    try { [AudioBt.Helper]::LogSnapshot("after") } catch { }
    Write-AudioBtPsLog ("result code=" + $code + " " + $result)
    Write-AudioBtPsLog ("======== end " + $ctx.stamp + " ========")
    $leaf = Split-Path -Leaf $ctx.session
    $line = (Get-Date -Format "yyyy-MM-dd HH:mm:ss") + "  " + $ctx.stamp + "  code=" + $code + "  " + $leaf + "  " + $result
    $utf8 = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::AppendAllText($ctx.index, $line + [Environment]::NewLine, $utf8)
    Copy-Item -LiteralPath $ctx.session -Destination $ctx.latest -Force
}

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

function Wait-AudioBtWinRT($async, [int]$timeoutMs = 12000, [type]$resultType = $null) {
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
        $sigs = @()
        foreach ($m in [System.WindowsRuntimeSystemExtensions].GetMethods()) {
            if ($m.Name -eq "AsTask") {
                $sigs += $m.ToString()
            }
        }
        Write-AudioBtPsLog ("WinRT AsTask unavailable methods=" + ($sigs -join "; "))
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
        Write-AudioBtPsLog ("WinRT AsTask fault elapsedMs=" + $sw.ElapsedMilliseconds + " status=" + (Get-AudioBtWinRTStatus $async) + " " + $inner.Message)
        throw $inner
    }
    $st = Get-AudioBtWinRTStatus $async
    if (-not $finished) {
        Write-AudioBtPsLog ("WinRT AsTask timeout elapsedMs=" + $sw.ElapsedMilliseconds + " status=" + $st)
        throw ("WinRT async timeout after " + $sw.ElapsedMilliseconds + "ms status=" + $st)
    }
    if ($task.IsFaulted) {
        $ex = $task.Exception
        if ($null -ne $ex.InnerException) {
            $ex = $ex.InnerException
        }
        Write-AudioBtPsLog ("WinRT AsTask fault elapsedMs=" + $sw.ElapsedMilliseconds + " status=" + $st + " " + $ex.Message)
        throw $ex
    }
    Write-AudioBtPsLog ("WinRT AsTask ok elapsedMs=" + $sw.ElapsedMilliseconds + " status=" + $st)
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

function Test-AudioBtMacMatch([string]$blob, [string]$hex) {
    if ([string]::IsNullOrWhiteSpace($blob) -or [string]::IsNullOrWhiteSpace($hex)) {
        return $false
    }
    $compact = $blob.ToUpperInvariant() -replace "[:\-_]", ""
    return $compact.Contains($hex.ToUpperInvariant())
}

function Get-AudioBtWinRTDeviceBlob($d) {
    $parts = New-Object System.Collections.Generic.List[string]
    try { $parts.Add([string]$d.Id) } catch { }
    try { $parts.Add([string]$d.Name) } catch { }
    try {
        $keys = @(
            "System.Devices.Aep.DeviceAddress",
            "System.DeviceInterface.Bluetooth.DeviceAddress",
            "System.Devices.Aep.IsConnected"
        )
        foreach ($k in $keys) {
            $v = $null
            try { $v = $d.Properties[$k] } catch { $v = $null }
            if ($null -ne $v) {
                $parts.Add($k + "=" + [string]$v)
            }
        }
    } catch { }
    return ($parts -join " ")
}

function Find-AudioBtWinRTDevices([string]$label, [string]$aqs, [bool]$aep) {
    Write-AudioBtPsLog ("WinRT FindAllAsync label=" + $label + " aqs=" + $aqs)
    $collType = [Windows.Devices.Enumeration.DeviceInformationCollection]
    try {
        $op = $null
        if ($aep) {
            $props = [string[]]@(
                "System.Devices.Aep.DeviceAddress",
                "System.Devices.Aep.IsConnected",
                "System.Devices.Aep.ProtocolId"
            )
            $kind = [Windows.Devices.Enumeration.DeviceInformationKind]::AssociationEndpoint
            try {
                $op = [Windows.Devices.Enumeration.DeviceInformation]::FindAllAsync($aqs, $props, $kind)
            } catch {
                Write-AudioBtPsLog ("WinRT FindAllAsync AEP kind overload fail, retry aqs-only: " + $_.Exception.Message)
                $op = [Windows.Devices.Enumeration.DeviceInformation]::FindAllAsync($aqs)
            }
        } else {
            $op = [Windows.Devices.Enumeration.DeviceInformation]::FindAllAsync($aqs)
        }
        $devices = Wait-AudioBtWinRT $op 12000 $collType
        $count = 0
        try { $count = @($devices).Count } catch { $count = -1 }
        Write-AudioBtPsLog ("WinRT FindAllAsync label=" + $label + " count=" + $count)
        return $devices
    } catch {
        Write-AudioBtPsLog ("WinRT FindAllAsync label=" + $label + " fail: " + $_.Exception.Message)
        return $null
    }
}

function Invoke-AudioBtWinRTOpen([string]$deviceId) {
    Write-AudioBtPsLog ("WinRT TryCreateFromId " + $deviceId)
    $conn = [Windows.Media.Audio.AudioPlaybackConnection]::TryCreateFromId($deviceId)
    if ($null -eq $conn) {
        Write-AudioBtPsLog "WinRT TryCreateFromId returned null"
        return @{ ok = $false; err = "AudioPlaybackConnection.TryCreateFromId failed" }
    }
    Wait-AudioBtWinRT ($conn.StartAsync()) 12000 | Out-Null
    Write-AudioBtPsLog "WinRT StartAsync completed"
    $openType = [Windows.Media.Audio.AudioPlaybackConnectionOpenResult]
    $open = Wait-AudioBtWinRT ($conn.OpenAsync()) 12000 $openType
    $status = ""
    if ($null -ne $open) {
        try { $status = [string]$open.Status } catch { $status = "" }
    }
    Write-AudioBtPsLog ("WinRT OpenAsync status=" + $status)
    if ($status -ne "" -and $status -ne "0" -and $status -ne "Success") {
        return @{ ok = $false; err = "AudioPlaybackConnection.OpenAsync $status" }
    }
    Write-AudioBtPsLog "WinRT ok"
    return @{ ok = $true; err = "" }
}

function Invoke-AudioBtWinRTConnect([string]$Id) {
    $hex = [AudioBt.Helper]::BluetoothAddressHex($Id)
    Write-AudioBtPsLog ("WinRT begin id=" + $Id + " hex=" + $hex)
    if ([string]::IsNullOrWhiteSpace($hex)) {
        Write-AudioBtPsLog "WinRT no bluetooth address"
        return @{ ok = $false; err = "No Bluetooth address for WinRT connect" }
    }
    $hex = $hex.ToUpperInvariant() -replace "[:\-_]", ""
    try {
        Add-Type -AssemblyName System.Runtime.WindowsRuntime -ErrorAction SilentlyContinue
        $null = [Windows.Media.Audio.AudioPlaybackConnection, Windows.Media.Audio, ContentType = WindowsRuntime]
        $null = [Windows.Devices.Enumeration.DeviceInformation, Windows.Devices.Enumeration, ContentType = WindowsRuntime]
        $null = [Windows.Devices.Enumeration.DeviceInformationKind, Windows.Devices.Enumeration, ContentType = WindowsRuntime]
        $null = [Windows.Media.Audio.AudioPlaybackConnectionOpenResult, Windows.Media.Audio, ContentType = WindowsRuntime]
    } catch {
        Write-AudioBtPsLog ("WinRT types unavailable: " + $_.Exception.Message)
        return @{ ok = $false; err = "AudioPlaybackConnection unavailable: $($_.Exception.Message)" }
    }

    try {
        $selector = [Windows.Media.Audio.AudioPlaybackConnection]::GetDeviceSelector()
        Write-AudioBtPsLog ("WinRT selector=" + $selector)
        $unfiltered = ($selector -replace "(?i)\s*AND\s+System\.Devices\.InterfaceEnabled:=System\.StructuredQueryType\.Boolean#True", "").Trim()
        $queries = New-Object System.Collections.Generic.List[object]
        $queries.Add(@{ label = "AudioPlaybackConnection"; aqs = $selector; aep = $false })
        if ($unfiltered -ne $selector -and -not [string]::IsNullOrWhiteSpace($unfiltered)) {
            $queries.Add(@{ label = "AudioPlaybackUnfiltered"; aqs = $unfiltered; aep = $false })
        }
        $queries.Add(@{ label = "AepBluetooth"; aqs = 'System.Devices.Aep.ProtocolId:="{e0cbf06c-cd8b-4647-bb8a-263b43f0f974}"'; aep = $true })
        $lastErr = "AudioPlaybackConnection: no device matching $hex"
        foreach ($q in $queries) {
            if ([string]::IsNullOrWhiteSpace($q.aqs)) {
                continue
            }
            $devices = Find-AudioBtWinRTDevices $q.label $q.aqs $q.aep
            if ($null -eq $devices) {
                $lastErr = "AudioPlaybackConnection: FindAllAsync $($q.label) failed"
                continue
            }
            $match = $null
            foreach ($d in $devices) {
                $blob = Get-AudioBtWinRTDeviceBlob $d
                Write-AudioBtPsLog ("WinRT candidate " + $q.label + " " + $blob)
                if (Test-AudioBtMacMatch $blob $hex) {
                    $match = $d
                    break
                }
            }
            if ($null -eq $match) {
                Write-AudioBtPsLog ("WinRT no match for " + $hex + " via " + $q.label)
                continue
            }
            Write-AudioBtPsLog ("WinRT match label=" + $q.label + " id=" + $match.Id)
            $opened = Invoke-AudioBtWinRTOpen $match.Id
            if ($opened.ok) {
                return $opened
            }
            $lastErr = $opened.err
        }
        return @{ ok = $false; err = $lastErr }
    } catch {
        Write-AudioBtPsLog ("WinRT exception: " + $_.Exception.Message)
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
    Write-AudioBtPsLog ("helper Connect => " + $r)
    if ($r -and -not $r.StartsWith("ERR`t")) {
        return $r
    }
    Write-AudioBtPsLog "trying WinRT connect fallback"
    $w = Invoke-AudioBtWinRTConnect $Id
    Write-AudioBtPsLog ("WinRT connect ok=" + $w.ok + " err=" + $w.err)
    if ($w.ok) {
        $c = [AudioBt.Helper]::ConfirmConnected($Id, "AudioPlaybackConnection")
        Write-AudioBtPsLog ("ConfirmConnected => " + $c)
        if ($c) {
            return $c
        }
    }
    return (Merge-AudioBtErr $r $w.err)
}

function Invoke-AudioBtIsolateOrWinRT([string]$Id) {
    $r = [AudioBt.Helper]::Isolate($Id)
    Write-AudioBtPsLog ("helper Isolate => " + $r)
    if ($r -and -not $r.StartsWith("ERR`t")) {
        return $r
    }
    if (-not (Test-AudioBtNeedsConnectFallback $r)) {
        Write-AudioBtPsLog "skip WinRT fallback (SetDefault failure)"
        return $r
    }
    Write-AudioBtPsLog "trying WinRT isolate fallback"
    $w = Invoke-AudioBtWinRTConnect $Id
    Write-AudioBtPsLog ("WinRT connect ok=" + $w.ok + " err=" + $w.err)
    if ($w.ok) {
        $c = [AudioBt.Helper]::ConfirmConnected($Id, "AudioPlaybackConnection")
        Write-AudioBtPsLog ("ConfirmConnected => " + $c)
        if ($c -and -not $c.StartsWith("ERR`t")) {
            $r2 = [AudioBt.Helper]::Isolate($Id)
            Write-AudioBtPsLog ("Isolate after WinRT => " + $r2)
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

$logCtx = $null
if ($Action -in @("connect", "isolate", "disconnect") -and -not [string]::IsNullOrWhiteSpace($LogDir)) {
    $logCtx = New-AudioBtSessionLog $LogDir $Action $Id $Note
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
    Complete-AudioBtSessionLog $logCtx "ERR`tEmpty helper result" 1
    Write-AudioBtResult "ERR`tEmpty helper result" 1
}

$code = 0
if ($result.StartsWith("ERR`t")) {
    $code = 1
}
Complete-AudioBtSessionLog $logCtx $result $code
Write-AudioBtResult $result $code
