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
        Write-AudioBtPsLog ("WinRT AsTask fault elapsedMs=" + $sw.ElapsedMilliseconds + " " + (Get-AudioBtTaskBlob $task $async) + " " + $inner.Message)
        throw $inner
    }
    if (-not $finished) {
        Write-AudioBtPsLog ("WinRT AsTask timeout elapsedMs=" + $sw.ElapsedMilliseconds + " " + (Get-AudioBtTaskBlob $task $async))
        try { $async.Cancel() } catch { }
        throw ("WinRT async timeout after " + $sw.ElapsedMilliseconds + "ms " + (Get-AudioBtTaskBlob $task $async))
    }
    if ($task.IsFaulted) {
        $ex = $task.Exception
        if ($null -ne $ex.InnerException) {
            $ex = $ex.InnerException
        }
        Write-AudioBtPsLog ("WinRT AsTask fault elapsedMs=" + $sw.ElapsedMilliseconds + " " + (Get-AudioBtTaskBlob $task $async) + " " + $ex.Message)
        throw $ex
    }
    Write-AudioBtPsLog ("WinRT AsTask ok elapsedMs=" + $sw.ElapsedMilliseconds + " " + (Get-AudioBtTaskBlob $task $async))
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

function Write-AudioBtAclStatus([string]$when, $dev) {
    $st = Get-AudioBtConnectionStatus $dev
    $did = ""
    try { $did = [string]$dev.DeviceId } catch { }
    Write-AudioBtPsLog ("ACL status " + $when + "=" + $st + " id=" + $did)
    return $st
}

function Get-AudioBtRfcommCount($result) {
    try {
        return @($result.Services).Count
    } catch {
        return -1
    }
}

function Invoke-AudioBtRfcommUncached($dev) {
    $resultType = [Windows.Devices.Bluetooth.Rfcomm.RfcommDeviceServicesResult]
    $uncached = [Windows.Devices.Bluetooth.BluetoothCacheMode]::Uncached
    try {
        return Wait-AudioBtWinRT ($dev.GetRfcommServicesAsync($uncached)) 8000 $resultType
    } catch {
        Write-AudioBtPsLog ("ACL GetRfcommServicesAsync(Uncached) fail: " + $_.Exception.Message)
        try {
            return Wait-AudioBtWinRT ($dev.GetRfcommServicesAsync()) 8000 $resultType
        } catch {
            Write-AudioBtPsLog ("ACL GetRfcommServicesAsync fail: " + $_.Exception.Message)
            return $null
        }
    }
}

function Invoke-AudioBtHfpStream($dev) {
    Write-AudioBtPsLog "ACL trying HFP 111E GetStreamAsync"
    try {
        $hfp = [Windows.Devices.Bluetooth.Rfcomm.RfcommServiceId]::FromUuid([guid]"0000111e-0000-1000-8000-00805f9b34fb")
        $resultType = [Windows.Devices.Bluetooth.Rfcomm.RfcommDeviceServicesResult]
        $uncached = [Windows.Devices.Bluetooth.BluetoothCacheMode]::Uncached
        $result = $null
        try {
            $result = Wait-AudioBtWinRT ($dev.GetRfcommServicesForIdAsync($hfp, $uncached)) 8000 $resultType
        } catch {
            Write-AudioBtPsLog ("ACL GetRfcommServicesForIdAsync Uncached fail: " + $_.Exception.Message)
            $result = Wait-AudioBtWinRT ($dev.GetRfcommServicesForIdAsync($hfp)) 8000 $resultType
        }
        $err = ""
        try { $err = [string]$result.Error } catch { }
        $n = Get-AudioBtRfcommCount $result
        Write-AudioBtPsLog ("ACL HFP services error=" + $err + " count=" + $n)
        $svc = $null
        foreach ($s in $result.Services) {
            $svc = $s
            break
        }
        if ($null -eq $svc) {
            Write-AudioBtPsLog "ACL HFP no RfcommDeviceService"
            return
        }
        $streamType = [Windows.Storage.Streams.IRandomAccessStream]
        $stream = $null
        try {
            $stream = Wait-AudioBtWinRT ($svc.GetStreamAsync()) 8000 $streamType
        } catch {
            Write-AudioBtPsLog ("ACL GetStreamAsync typed fail, retry untyped: " + $_.Exception.Message)
            try {
                $stream = Wait-AudioBtWinRT ($svc.GetStreamAsync()) 8000
            } catch {
                Write-AudioBtPsLog ("ACL GetStreamAsync fail: " + $_.Exception.Message)
                return
            }
        }
        Write-AudioBtPsLog ("ACL GetStreamAsync got=" + ($null -ne $stream))
        if ($null -ne $stream) {
            try { $stream.Dispose() } catch { }
        }
    } catch {
        Write-AudioBtPsLog ("ACL HFP stream fail: " + $_.Exception.Message)
    }
}

function Invoke-AudioBtAclConnect([string]$hex) {
    Write-AudioBtPsLog ("ACL FromBluetoothAddressAsync hex=" + $hex)
    $addr = [uint64]::Parse($hex, [Globalization.NumberStyles]::HexNumber)
    $btType = [Windows.Devices.Bluetooth.BluetoothDevice]
    $dev = Wait-AudioBtWinRT ([Windows.Devices.Bluetooth.BluetoothDevice]::FromBluetoothAddressAsync($addr)) 4000 $btType
    if ($null -eq $dev) {
        Write-AudioBtPsLog "ACL FromBluetoothAddressAsync returned null"
        return @{ ok = $false; err = "BluetoothDevice.FromBluetoothAddressAsync returned null"; status = "" }
    }
    $nm = ""
    try { $nm = [string]$dev.Name } catch { }
    Write-AudioBtPsLog ("ACL BluetoothDevice name=" + $nm)
    $before = Write-AudioBtAclStatus "before" $dev
    $result = Invoke-AudioBtRfcommUncached $dev
    if ($null -ne $result) {
        $rfErr = ""
        try { $rfErr = [string]$result.Error } catch { }
        Write-AudioBtPsLog ("ACL GetRfcommServicesAsync error=" + $rfErr + " count=" + (Get-AudioBtRfcommCount $result))
    }
    $afterRf = Write-AudioBtAclStatus "after-rfcomm" $dev
    if ($afterRf -ne "Connected") {
        Invoke-AudioBtHfpStream $dev
        $afterRf = Write-AudioBtAclStatus "after-hfp-stream" $dev
    }
    $ok = ($afterRf -eq "Connected")
    Write-AudioBtPsLog ("ACL done ok=" + $ok + " before=" + $before + " after=" + $afterRf)
    return @{ ok = $ok; err = $(if ($ok) { "" } else { "BluetoothDevice still " + $afterRf }); status = $afterRf }
}

function Invoke-AudioBtWinRTConnect([string]$Id) {
    $hex = [AudioBt.Helper]::BluetoothAddressHex($Id)
    Write-AudioBtPsLog ("WinRT begin id=" + $Id + " hex=" + $hex)
    try { [AudioBt.Helper]::DumpConnectDiagnostics($Id) } catch { Write-AudioBtPsLog ("DumpConnectDiagnostics: " + $_.Exception.Message) }
    if ([string]::IsNullOrWhiteSpace($hex)) {
        Write-AudioBtPsLog "WinRT no bluetooth address"
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
        Write-AudioBtPsLog ("WinRT BluetoothDevice types unavailable: " + $_.Exception.Message)
        return @{ ok = $false; err = "BluetoothDevice unavailable: $($_.Exception.Message)" }
    }

    try {
        Write-AudioBtPsLog "WinRT skip TryCreateFromId on A2DP sink/AEP ids (wrong API)"
        Write-AudioBtPsLog "WinRT skip AepBluetooth FindAllAsync (hangs)"
        $acl = Invoke-AudioBtAclConnect $hex
        Write-AudioBtPsLog ("ACL connect ok=" + $acl.ok + " err=" + $acl.err + " status=" + $acl.status)
        if (-not $acl.ok) {
            return @{ ok = $false; err = $acl.err }
        }
        $en = [AudioBt.Helper]::EnableBtAudioServices($Id)
        Write-AudioBtPsLog ("EnableBtAudioServices => " + $en)
        return @{ ok = $true; err = "" }
    } catch {
        Write-AudioBtPsLog ("WinRT exception: " + $_.Exception.Message)
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
    Write-AudioBtPsLog ("helper Connect => " + $r)
    if ($r -and -not $r.StartsWith("ERR`t")) {
        return $r
    }
    Write-AudioBtPsLog "trying WinRT connect fallback"
    $w = Invoke-AudioBtWinRTConnect $Id
    Write-AudioBtPsLog ("WinRT connect ok=" + $w.ok + " err=" + $w.err)
    if ($w.ok) {
        $c = [AudioBt.Helper]::ConfirmConnected($Id, "BluetoothDevice")
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
        $c = [AudioBt.Helper]::ConfirmConnected($Id, "BluetoothDevice")
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
