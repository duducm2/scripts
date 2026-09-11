# Build ClipAngelDb.exe (x86) against vendored System.Data.SQLite.dll
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$src = Join-Path $PSScriptRoot "ClipAngelDb.cs"
$bin = Join-Path $root "bin"
$dll = Join-Path $bin "System.Data.SQLite.dll"
$out = Join-Path $bin "ClipAngelDb.exe"
$csc = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe"

if (!(Test-Path $csc)) { throw "csc.exe not found: $csc" }
if (!(Test-Path $src)) { throw "source not found: $src" }
if (!(Test-Path $dll)) { throw "SQLite DLL not found: $dll (copy from ClipAngel install)" }

New-Item -ItemType Directory -Force -Path $bin | Out-Null

& $csc /nologo /optimize+ /platform:x86 /target:exe /out:"$out" /r:"$dll" "$src"
if ($LASTEXITCODE -ne 0) { throw "csc failed with exit $LASTEXITCODE" }

Write-Host "Built: $out"
Get-Item $out | Select-Object FullName, Length, LastWriteTime
