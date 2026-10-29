$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Drawing

$source = 'D:\Huawei Handover\PAT DOC_LP89M_2026 Capacity Expansion NW Modernization_EXP_BKK_hidden_images_removed.xlsx'
$backup = $source -replace '\.xlsx$', '_before_image_compression.xlsx'
$temp = $source -replace '\.xlsx$', '_image_compression_tmp.xlsx'

if (-not (Test-Path -LiteralPath $source)) { throw "Target workbook not found: $source" }
if (Test-Path -LiteralPath $temp) { Remove-Item -LiteralPath $temp -Force }
if (-not (Test-Path -LiteralPath $backup)) { Copy-Item -LiteralPath $source -Destination $backup }

function Get-Bytes($entry) {
    $stream = $entry.Open()
    try { $memory = New-Object IO.MemoryStream; $stream.CopyTo($memory); return $memory.ToArray() } finally { $stream.Dispose() }
}

function Compress-ImageBytes([byte[]]$bytes, [string]$extension) {
    $input = New-Object IO.MemoryStream(,$bytes)
    $image = $null
    try {
        $image = [Drawing.Image]::FromStream($input, $true, $true)
        $output = New-Object IO.MemoryStream
        try {
            if ($extension -match '\.jpe?g$') {
                $codec = [Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object MimeType -eq 'image/jpeg' | Select-Object -First 1
                $parameters = New-Object Drawing.Imaging.EncoderParameters(1)
                $parameters.Param[0] = New-Object Drawing.Imaging.EncoderParameter([Drawing.Imaging.Encoder]::Quality, [long]90)
                $image.Save($output, $codec, $parameters)
                $parameters.Dispose()
            } elseif ($extension -match '\.png$') {
                $image.Save($output, [Drawing.Imaging.ImageFormat]::Png)
            } else { return $bytes }
            $candidate = $output.ToArray()
            if ($candidate.Length -lt $bytes.Length) { return $candidate }
            return $bytes
        } finally { $output.Dispose() }
    } catch { return $bytes }
    finally { if ($null -ne $image) { $image.Dispose() }; $input.Dispose() }
}

$sourceZip = [IO.Compression.ZipFile]::OpenRead($source)
$outputStream = [IO.File]::Open($temp, [IO.FileMode]::CreateNew)
$outputZip = New-Object IO.Compression.ZipArchive($outputStream, [IO.Compression.ZipArchiveMode]::Create, $false)
$mediaBefore = 0
$mediaAfterBytes = 0
$mediaBeforeBytes = 0
$changed = 0
try {
    foreach ($entry in $sourceZip.Entries) {
        $newEntry = $outputZip.CreateEntry($entry.FullName, [IO.Compression.CompressionLevel]::Optimal)
        try { $newEntry.LastWriteTime = $entry.LastWriteTime } catch {}
        $targetStream = $newEntry.Open()
        try {
            $bytes = Get-Bytes $entry
            if ($entry.FullName -match '^xl/media/[^/]+\.(jpe?g|png)$') {
                $mediaBefore++
                $mediaBeforeBytes += $bytes.Length
                $compressed = Compress-ImageBytes $bytes ([IO.Path]::GetExtension($entry.FullName))
                if ($compressed.Length -lt $bytes.Length) { $changed++; $bytes = $compressed }
                $mediaAfterBytes += $bytes.Length
            }
            $targetStream.Write($bytes, 0, $bytes.Length)
        } finally { $targetStream.Dispose() }
    }
} finally {
    $outputZip.Dispose(); $outputStream.Dispose(); $sourceZip.Dispose()
}

$beforeLength = (Get-Item -LiteralPath $source).Length
$afterLength = (Get-Item -LiteralPath $temp).Length
if ($afterLength -ge $beforeLength) {
    Remove-Item -LiteralPath $temp -Force
    throw "Compression did not reduce the workbook size. Original was preserved."
}

# Validate the temporary package before replacing the prepared template.
$check = [IO.Compression.ZipFile]::OpenRead($temp)
try {
    $names = @($check.Entries | Select-Object -ExpandProperty FullName)
    if ($names -notcontains 'xl/workbook.xml' -or $names -notcontains '[Content_Types].xml') { throw 'Temporary workbook package validation failed.' }
    $mediaAfterCount = @($check.Entries | Where-Object FullName -match '^xl/media/').Count
} finally { $check.Dispose() }

Move-Item -LiteralPath $temp -Destination $source -Force

[pscustomobject]@{
    Workbook = $source
    Backup = $backup
    MediaFiles = $mediaBefore
    ImagesSmaller = $changed
    MediaBytesBefore = $mediaBeforeBytes
    MediaBytesAfter = $mediaAfterBytes
    WorkbookBytesBefore = $beforeLength
    WorkbookBytesAfter = $afterLength
    WorkbookReductionBytes = ($beforeLength - $afterLength)
    WorkbookReductionPercent = [math]::Round((($beforeLength - $afterLength) / $beforeLength) * 100, 2)
    MediaEntriesAfter = $mediaAfterCount
} | ConvertTo-Json
