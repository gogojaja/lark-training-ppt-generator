# PowerShell Windows OCR helper (offline, no install).
# Uses Windows 10/11 built-in WinRT OCR. Pure ASCII to avoid PS 5.1 encoding issues.
param(
  [Parameter(Mandatory=$true)][string]$ImagePath,
  [Parameter(Mandatory=$true)][string]$OutJson,
  [string]$Lang = "zh-Hans-CN"
)

$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Runtime.WindowsRuntime

$asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
})[0]

Function Await($WinRtTask, $ResultType) {
    $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
    $netTask = $asTask.Invoke($null, @($WinRtTask))
    $netTask.Wait() | Out-Null
    $netTask.Result
}

[Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime] > $null
[Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime] > $null
[Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics, ContentType = WindowsRuntime] > $null
[Windows.Globalization.Language, Windows.Globalization, ContentType = WindowsRuntime] > $null

$file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($ImagePath)) ([Windows.Storage.StorageFile])
$stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
$decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
$bitmap = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])

$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage([Windows.Globalization.Language]::new($Lang))
if ($null -eq $engine) {
    $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
}
if ($null -eq $engine) {
    Write-Error "no OCR engine available for: $Lang"
    exit 1
}

$result = Await ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])

$words = New-Object System.Collections.ArrayList
foreach ($line in $result.Lines) {
    foreach ($w in $line.Words) {
        $r = $w.BoundingRect
        [void]$words.Add([PSCustomObject]@{
            text   = $w.Text
            x      = [math]::Round($r.X, 1)
            y      = [math]::Round($r.Y, 1)
            w      = [math]::Round($r.Width, 1)
            h      = [math]::Round($r.Height, 1)
        })
    }
}

$resultObj = @{
    imageWidth  = $bitmap.PixelWidth
    imageHeight = $bitmap.PixelHeight
    words       = $words
    engineLang  = $engine.RecognizerLanguage.LanguageTag
}
$resultObj | ConvertTo-Json -Depth 5 | Out-File -FilePath $OutJson -Encoding utf8
Write-Host "OCR done: $($words.Count) words, lang=$($engine.RecognizerLanguage.LanguageTag)"
