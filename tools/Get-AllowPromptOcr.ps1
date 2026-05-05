param(
    [Parameter(Mandatory = $true)][int]$X,
    [Parameter(Mandatory = $true)][int]$Y,
    [Parameter(Mandatory = $true)][int]$Width,
    [Parameter(Mandatory = $true)][int]$Height,
    [Parameter(Mandatory = $true)][string]$OutFile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Write-Result {
    param([string]$Line)
    Set-Content -Path $OutFile -Value $Line -Encoding UTF8
}

function Await-WinRT {
    param([object]$AsyncOp)
    return [System.WindowsRuntimeSystemExtensions]::AsTask($AsyncOp).GetAwaiter().GetResult()
}

if ($Width -lt 1 -or $Height -lt 1) {
    Write-Result "NONE|0|0|invalid-crop"
    exit 1
}

$bitmap = $null
$graphics = $null
$tmpPng = Join-Path $env:TEMP ("aib_allow_ocr_" + [Guid]::NewGuid().ToString("N") + ".png")

try {
    Add-Type -AssemblyName System.Drawing

    $bitmap = New-Object System.Drawing.Bitmap $Width, $Height
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.CopyFromScreen($X, $Y, 0, 0, $bitmap.Size)
    $bitmap.Save($tmpPng, [System.Drawing.Imaging.ImageFormat]::Png)

    Add-Type -AssemblyName System.Runtime.WindowsRuntime
    $null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
    $null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Foundation, ContentType = WindowsRuntime]
    $null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
    $null = [Windows.Storage.FileAccessMode, Windows.Storage, ContentType = WindowsRuntime]

    $file = Await-WinRT ([Windows.Storage.StorageFile]::GetFileFromPathAsync($tmpPng))
    $stream = Await-WinRT ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read))
    $decoder = Await-WinRT ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream))
    $softwareBitmap = Await-WinRT ($decoder.GetSoftwareBitmapAsync())

    $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
    if (-not $engine) {
        Write-Result "NONE|0|0|ocr-engine-unavailable"
        exit 1
    }

    $ocrResult = Await-WinRT ($engine.RecognizeAsync($softwareBitmap))
    $allText = (($ocrResult.Text -replace "`r", " " -replace "`n", " ").Trim())
    $normalized = $allText.ToLowerInvariant()

    if ($normalized -notmatch '(^|\W)allow(\W|$)') {
        Write-Result ("NONE|0|0|" + ($normalized -replace "\|", " "))
        exit 0
    }

    $hitX = 0
    $hitY = 0
    $foundWord = $false

    foreach ($line in $ocrResult.Lines) {
        foreach ($word in $line.Words) {
            $w = $word.Text.ToLowerInvariant()
            if ($w -match '(^|\W)allow(\W|$)') {
                $rect = $word.BoundingRect
                $hitX = [int]([Math]::Round($X + $rect.X + ($rect.Width / 2.0)))
                $hitY = [int]([Math]::Round($Y + $rect.Y + ($rect.Height / 2.0)))
                $foundWord = $true
                break
            }
        }
        if ($foundWord) { break }
    }

    if (-not $foundWord) {
        $hitX = [int]([Math]::Round($X + ($Width / 2.0)))
        $hitY = [int]([Math]::Round($Y + ($Height / 2.0)))
    }

    $safeText = ($normalized -replace "\|", " ")
    Write-Result ("FOUND|$hitX|$hitY|$safeText")
    exit 0
}
catch {
    $msg = $_.Exception.Message
    if (-not $msg) { $msg = "ocr-error" }
    Write-Result ("NONE|0|0|" + ($msg -replace "\|", " "))
    exit 1
}
finally {
    if ($graphics) { $graphics.Dispose() }
    if ($bitmap) { $bitmap.Dispose() }
    if (Test-Path -LiteralPath $tmpPng) {
        Remove-Item -LiteralPath $tmpPng -Force -ErrorAction SilentlyContinue
    }
}
