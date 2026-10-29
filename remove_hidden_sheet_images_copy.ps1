$ErrorActionPreference = 'Stop'

$source = 'D:\Huawei Handover\PAT DOC_LP89M_2026 Capacity Expansion NW Modernization_EXP_BKK.xlsx'
$baseOutput = 'D:\Huawei Handover\PAT DOC_LP89M_2026 Capacity Expansion NW Modernization_EXP_BKK_hidden_images_removed.xlsx'
$output = $baseOutput
$suffix = 2
while (Test-Path -LiteralPath $output) {
    $output = $baseOutput -replace '\.xlsx$', ("_v{0}.xlsx" -f $suffix)
    $suffix++
}
$resultPath = 'D:\Huawei Handover\reports\hidden_image_removal_result.md'

$sourceHashBefore = (Get-FileHash -Algorithm SHA256 -LiteralPath $source).Hash
$sourceSize = (Get-Item -LiteralPath $source).Length
Copy-Item -LiteralPath $source -Destination $output

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $null
$deleted = 0
$hiddenPictureBefore = 0
$visiblePictureBefore = 0
$hiddenSheetCount = 0
$visibleSheetCount = 0
$hiddenSheetsWithPictures = 0

try {
    # UpdateLinks=0 and ReadOnly=false for the copy only.
    $wb = $excel.Workbooks.Open($output, 0, $false)
    foreach ($ws in $wb.Worksheets) {
        $isVisible = ([int]$ws.Visible -eq -1)
        if ($isVisible) { $visibleSheetCount++ } else { $hiddenSheetCount++ }
        $countBefore = 0
        for ($i = [int]$ws.Shapes.Count; $i -ge 1; $i--) {
            $shape = $ws.Shapes.Item($i)
            $type = [int]$shape.Type
            if ($type -eq 13 -or $type -eq 11) {
                $countBefore++
                if ($isVisible) { $visiblePictureBefore++ } else { $hiddenPictureBefore++ }
                if (-not $isVisible) {
                    $shape.Delete()
                    $deleted++
                }
            }
        }
        if (-not $isVisible -and $countBefore -gt 0) { $hiddenSheetsWithPictures++ }
    }
    $wb.Save()
    $wb.Close($true)
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
    $wb = $null
}
finally {
    if ($null -ne $wb) { $wb.Close($false); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) }
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

# Reopen the saved copy read-only for validation.
$excel2 = New-Object -ComObject Excel.Application
$excel2.Visible = $false
$excel2.DisplayAlerts = $false
$checkWb = $null
$hiddenPictureAfter = 0
$visiblePictureAfter = 0
$visibleFormulaErrorCount = 0
$copyExternalLinks = 0
try {
    $checkWb = $excel2.Workbooks.Open($output, 0, $true)
    foreach ($ws in $checkWb.Worksheets) {
        $isVisible = ([int]$ws.Visible -eq -1)
        for ($i = 1; $i -le [int]$ws.Shapes.Count; $i++) {
            $type = [int]$ws.Shapes.Item($i).Type
            if ($type -eq 13 -or $type -eq 11) {
                if ($isVisible) { $visiblePictureAfter++ } else { $hiddenPictureAfter++ }
            }
        }
        if ($isVisible) {
            foreach ($term in @('#REF!','#DIV/0!','#VALUE!','#NAME?','#N/A','#NUM!','#NULL!','#SPILL!','#CALC!')) {
                try {
                    $found = $ws.UsedRange.Find($term, $null, -4163, 1, 1, 1, $false, $false, $false)
                    if ($null -ne $found) { $visibleFormulaErrorCount++ }
                } catch {}
            }
        }
    }
    $links = $checkWb.LinkSources(1)
    if ($null -ne $links) { $copyExternalLinks = @($links).Count }
    $checkWb.Close($false)
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($checkWb)
    $checkWb = $null
}
finally {
    if ($null -ne $checkWb) { $checkWb.Close($false); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($checkWb) }
    $excel2.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel2)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

$sourceHashAfter = (Get-FileHash -Algorithm SHA256 -LiteralPath $source).Hash
$outputSize = (Get-Item -LiteralPath $output).Length
$reductionBytes = $sourceSize - $outputSize
$reductionPercent = if ($sourceSize -gt 0) { [math]::Round(($reductionBytes / $sourceSize) * 100, 2) } else { 0 }

$reportLines = @(
    '# Hidden Sheet Image Removal Result',
    '',
    "Source file: $source",
    "Output copy: $output",
    'Original workbook was not saved or modified.',
    '',
    "Visible Worksheets: $visibleSheetCount",
    "Hidden/Very Hidden Worksheets: $hiddenSheetCount",
    "Hidden sheets with pictures before: $hiddenSheetsWithPictures",
    "Hidden picture objects before: $hiddenPictureBefore",
    "Hidden picture objects deleted: $deleted",
    "Hidden picture objects after: $hiddenPictureAfter",
    "Visible picture objects before: $visiblePictureBefore",
    "Visible picture objects after: $visiblePictureAfter",
    "Visible formula-error matches after save: $visibleFormulaErrorCount",
    "External Links reported in output copy: $copyExternalLinks",
    "Original size (bytes): $sourceSize",
    "Output size (bytes): $outputSize",
    "Size reduction (bytes): $reductionBytes",
    "Size reduction (%): $reductionPercent",
    "Original SHA-256 before: $sourceHashBefore",
    "Original SHA-256 after: $sourceHashAfter"
)
Set-Content -LiteralPath $resultPath -Value ($reportLines -join [Environment]::NewLine) -Encoding UTF8

[pscustomobject]@{
    Output = $output
    DeletedHiddenPictures = $deleted
    HiddenPicturesAfter = $hiddenPictureAfter
    VisiblePicturesBefore = $visiblePictureBefore
    VisiblePicturesAfter = $visiblePictureAfter
    VisibleFormulaErrorMatches = $visibleFormulaErrorCount
    ExternalLinksAfter = $copyExternalLinks
    SourceSize = $sourceSize
    OutputSize = $outputSize
    ReductionBytes = $reductionBytes
    ReductionPercent = $reductionPercent
    OriginalHashUnchanged = ($sourceHashBefore -eq $sourceHashAfter)
    ResultReport = $resultPath
} | ConvertTo-Json
