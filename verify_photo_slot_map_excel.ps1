$ErrorActionPreference = 'Stop'

$workbookPath = 'D:\Huawei Handover\PAT DOC_LP89M_2026 Capacity Expansion NW Modernization_EXP_BKK.xlsx'
$reportDir = 'D:\Huawei Handover\reports'
$mapPath = Join-Path $reportDir 'photo_slot_map.csv'
$summaryPath = Join-Path $reportDir 'visible_worksheet_summary.csv'
$verifiedPath = Join-Path $reportDir 'photo_slot_map_verified.csv'
$reviewPath = Join-Path $reportDir 'photo_slot_map_review.csv'
$reportPath = Join-Path $reportDir 'photo_slot_map_verified_summary.md'

function Get-Category([string]$name) {
    $n = $name.Trim().ToLowerInvariant()
    if ($n -match '^sow') { return 'SOW' }
    if ($n -match 'dismantle') { return 'Dismantle' }
    if ($n -match '\bsn\b|ntc|sticker') { return 'SN' }
    if ($n -match 'card') { return 'Card' }
    if ($n -match 'dcdu') { return 'DCDU' }
    if ($n -match 'breaker') { return 'Breaker' }
    if ($n -match 'antenna|\bant\b') { return 'Antenna' }
    if ($n -match 'rru|aau') { return 'RRU/AAU' }
    if ($n -match 'bayface') { return 'BAYFACE' }
    if ($n -match 'content') { return 'Content' }
    if ($n -match 'checklist|boq') { return 'Checklist/BOQ' }
    if ($n -match 'alarm') { return 'Alarm' }
    if ($n -match 'tower|map|j-hanger|ladder') { return 'Site/Structure' }
    if ($n -match 'assessment|assestment|ehs|deficiency|clean|poe') { return 'Audit/Deficiency' }
    return 'Other'
}

function Get-CellRow([string]$address) {
    $m = [regex]::Match($address, '\d+')
    if ($m.Success) { return [int]$m.Value }
    return 0
}

function Get-CellColumn([string]$address) {
    $m = [regex]::Match($address, '^[A-Z]+', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $m.Success) { return 0 }
    $value = 0
    foreach ($ch in $m.Value.ToUpperInvariant().ToCharArray()) { $value = ($value * 26) + ([int][char]$ch - 64) }
    return $value
}

function Get-ValueAt($values, [int]$row, [int]$col, [int]$rowCount, [int]$colCount) {
    if ($rowCount -eq 1 -and $colCount -eq 1) { return $values }
    return $values[$row, $col]
}

function Add-Text($cache, [int]$row, [int]$col, [object]$value) {
    if ($null -eq $value) { return }
    $text = ([string]$value).Trim()
    if (-not $text) { return }
    $key = "$row`:$col"
    if (-not $cache.ContainsKey($key)) { $cache[$key] = $text }
}

function Get-TextFromWindow($cache, [int]$topRow, [int]$bottomRow, [int]$topCol, [int]$bottomCol) {
    $parts = New-Object System.Collections.Generic.List[string]
    $startRow = [Math]::Max(1, $topRow - 2)
    $endRow = $bottomRow + 3
    $startCol = [Math]::Max(1, $topCol - 2)
    $endCol = $bottomCol + 2
    foreach ($key in $cache.Keys) {
        $pieces = $key.Split(':')
        $r = [int]$pieces[0]
        $c = [int]$pieces[1]
        if ($r -ge $startRow -and $r -le $endRow -and $c -ge $startCol -and $c -le $endCol) {
            $text = [string]$cache[$key]
            if (-not $parts.Contains($text)) { [void]$parts.Add($text) }
        }
    }
    return ($parts -join ' | ')
}

function Get-Evidence([string]$text) {
    $terms = @('Before','After','Sector','SN','Activity','Remark','Installation','Article','Tilt','Dismantle','Label','View','Power')
    $parts = @($text -split '\s*\|\s*' | Where-Object { $_ -and (($terms | Where-Object { $_ -and $_.Length -gt 0 -and $_ -and $_ -as [string] -and $_.ToString() -and $_.ToString().Length -gt 0 -and $_.ToString().ToLowerInvariant() -and ([string]$text).ToLowerInvariant().Contains($_.ToString().ToLowerInvariant()) }).Count -gt 0) })
    return (($parts | Select-Object -Unique) -join ' | ')
}

function Get-Role([string]$evidence) {
    $hasBefore = $evidence -match '(?i)\bBefore\b'
    $hasAfter = $evidence -match '(?i)\bAfter\b'
    if ($hasBefore -and $hasAfter) { return 'Before_and_After_context' }
    if ($hasBefore) { return 'Before' }
    if ($hasAfter) { return 'After' }
    if ($evidence -match '(?i)\bSN\b|sticker') { return 'SN' }
    if ($evidence -match '(?i)\bRemark\b') { return 'Remark_context' }
    return 'UNKNOWN'
}

function Get-SectorEvidence([string]$evidence) {
    $matches = [regex]::Matches($evidence, '(?i)\bSector\s+([0-9]+(?:\.[0-9]+)?)')
    $values = @($matches | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique)
    if ($values.Count -eq 1) { return $values[0] }
    if ($values.Count -gt 1) { return ('MULTIPLE_EXPLICIT_VALUES:' + ($values -join ',')) }
    return 'UNKNOWN'
}

$mapRows = @(Import-Csv -LiteralPath $mapPath)
$summaryRows = @(Import-Csv -LiteralPath $summaryPath)
$summaryByName = @{}
foreach ($row in $summaryRows) { $summaryByName[[string]$row.Worksheet] = $row }
$mapSheetNames = @($mapRows | Select-Object -ExpandProperty Worksheet -Unique)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $null
$visibleByName = @{}
$textCacheBySheet = @{}
$shapeBySheetAndIndex = @{}
$verified = @()
$review = @()
$outOfScope = @()

try {
    $wb = $excel.Workbooks.Open($workbookPath, $null, $true)
    $visibleSheets = @($wb.Worksheets | Where-Object { $_.Visible -eq -1 })

    foreach ($ws in $visibleSheets) {
        $visibleByName[[string]$ws.Name] = $ws
        $cache = @{}
        $pa = [string]$ws.PageSetup.PrintArea
        if (-not $pa) { $pa = [string]$ws.UsedRange.Address() }
        $range = $ws.Range($pa)
        $values = $range.Value2
        $rowCount = [int]$range.Rows.Count
        $colCount = [int]$range.Columns.Count
        for ($r = 1; $r -le $rowCount; $r++) {
            for ($c = 1; $c -le $colCount; $c++) {
                Add-Text $cache ([int]$range.Row + $r - 1) ([int]$range.Column + $c - 1) (Get-ValueAt $values $r $c $rowCount $colCount)
            }
        }
        $textCacheBySheet[[string]$ws.Name] = $cache
        for ($i = 1; $i -le [int]$ws.Shapes.Count; $i++) {
            $shape = $ws.Shapes.Item($i)
            $shapeBySheetAndIndex[([string]$ws.Name + '|' + $i)] = $shape
        }
    }

    foreach ($row in $mapRows) {
        $sheetName = [string]$row.Worksheet
        if (-not $visibleByName.ContainsKey($sheetName)) {
            $outOfScope += $row
            continue
        }
        $key = $sheetName + '|' + [string]$row.ShapeIndex
        $shape = $shapeBySheetAndIndex[$key]
        $tl = $shape.TopLeftCell
        $br = $shape.BottomRightCell
        $topRow = [int]$tl.Row
        $bottomRow = [int]$br.Row
        $topCol = [int]$tl.Column
        $bottomCol = [int]$br.Column
        $nearby = Get-TextFromWindow $textCacheBySheet[$sheetName] $topRow $bottomRow $topCol $bottomCol
        $evidence = Get-Evidence $nearby
        $role = Get-Role $evidence
        $sector = Get-SectorEvidence $evidence
        $isHeaderCandidate = ($topRow -le 4)
        $hasSemanticEvidence = ($evidence -match '(?i)\bBefore\b|\bAfter\b|\bSN\b|sticker|\bSector\b|Installation|Article|Tilt|Dismantle|Label|Power')
        $shapeType = [int]$shape.Type
        $isPicture = ($shapeType -eq 13 -or $shapeType -eq 11)
        $base = [ordered]@{
            WorksheetIndex = $row.WorksheetIndex
            Worksheet = $sheetName
            WorksheetCategory = Get-Category $sheetName
            ShapeIndex = $row.ShapeIndex
            ShapeName = $row.ShapeName
            ShapeType = $shapeType
            IsPicture = $isPicture
            TopLeftCell = [string]$tl.Address($false, $false)
            BottomRightCell = [string]$br.Address($false, $false)
            Width = [math]::Round([double]$shape.Width, 1)
            Height = [math]::Round([double]$shape.Height, 1)
            NearbyText = $nearby
            CaptionEvidence = $evidence
            Activity = 'UNKNOWN'
            Sector = $sector
            ImageRole = $role
            SN = 'UNKNOWN'
            VerificationBasis = ''
            Confidence = ''
            ReviewRequired = ''
        }
        if ($isPicture -and -not $isHeaderCandidate -and $hasSemanticEvidence) {
            $base.VerificationBasis = 'Visible worksheet; picture object; explicit nearby worksheet text supports a caption/role or sector context.'
            $base.Confidence = 'MEDIUM_STRUCTURAL_TEXT_EVIDENCE'
            $base.ReviewRequired = 'YES_FOR_FINAL_SEMANTIC_ASSIGNMENT'
            $verified += [pscustomobject]$base
            continue
        }
        if (-not $isPicture) { $reason = 'CSV entry is not confirmed as an Excel picture after re-checking the shape type.' }
        elseif ($isHeaderCandidate) { $reason = 'Picture is anchored in the top header area; likely logo/header/decorative, but no semantic slot is assigned.' }
        elseif (-not $hasSemanticEvidence) { $reason = 'Picture has no nearby explicit Before/After/Sector/SN/Activity/Remark-style evidence in the inspected worksheet text window.' }
        else { $reason = 'Structure exists but semantic assignment remains ambiguous.' }
        $base.VerificationBasis = $reason
        $base.Confidence = 'INSUFFICIENT_FOR_SEMANTIC_SLOT'
        $base.ReviewRequired = 'YES'
        $review += [pscustomobject]$base
    }
}
finally {
    if ($null -ne $wb) {
        $wb.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
    }
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$verified | Export-Csv -LiteralPath $verifiedPath -NoTypeInformation -Encoding UTF8
$review | Export-Csv -LiteralPath $reviewPath -NoTypeInformation -Encoding UTF8

$missingMapSheets = @($summaryRows | Where-Object { $mapSheetNames -notcontains $_.Worksheet })
$categorySummary = @($summaryRows | Group-Object { Get-Category $_.Worksheet } | Sort-Object Name | ForEach-Object {
    $pictures = ($_.Group | Measure-Object -Property PictureCount -Sum).Sum
    "| $($_.Name) | $($_.Count) | $pictures |"
})
$missingLines = if ($missingMapSheets.Count -gt 0) { ($missingMapSheets | ForEach-Object { "- ``$($_.Worksheet)`` (PictureCount=$($_.PictureCount), ShapeCount=$($_.ShapeCount))" }) -join "`n" } else { '- None' }
$outScopeLines = if ($outOfScope.Count -gt 0) { "`nOut-of-scope map rows: $($outOfScope.Count)" } else { '' }

$report = @"
# Verified Photo Slot Map Summary

## Scope and safety

- The Excel workbook was opened Read-only.
- Only Visible Worksheets were inspected for cell text, shape type, shape bounds, and nearby text.
- Hidden Worksheets were not inspected, changed, or used for classification.
- No pictures were moved, deleted, or replaced. The workbook was not saved.

## Counts

| Item | Count |
|---|---:|
| Visible Worksheets in summary CSV | $($summaryRows.Count) |
| Unique worksheet names in input photo_slot_map.csv | $($mapSheetNames.Count) |
| Visible worksheet picture rows in input map | $($mapRows.Count) |
| Verified structural/text-evidence rows | $($verified.Count) |
| Rows requiring review | $($review.Count) |
| Map rows outside the Visible Worksheet set | $($outOfScope.Count) |

The input reports list 91 Visible Worksheets but 90 worksheet names in `photo_slot_map.csv` because the missing sheet is:

$missingLines

This is consistent with the prior report: `Content` is visible but has no embedded pictures, so it contributes no rows to the picture map.

## Worksheet groups

| Group | Visible Worksheets | Pictures in summary |
|---|---:|---:|
$($categorySummary -join "`n")

## Verification rule

A row is included in `photo_slot_map_verified.csv` only when the re-checked Excel object is a picture, it is not anchored in the top four rows, and nearby worksheet text contains explicit evidence such as `Before`, `After`, `Sector`, `SN`, `Installation`, `Article`, `Tilt`, `Dismantle`, `Label`, or `Power`.

This confirms a structural/text relationship only. It does not prove that the image is evidence for a particular Activity, Equipment Model, or Action. `Activity` and `SN` remain `UNKNOWN` unless the inspected evidence is explicit enough to support a future manual interpretation.

## Review requirements

The review file includes header/decorative candidates, pictures without nearby semantic text, and any object whose type did not re-confirm as a picture. Final assignment still requires a human check of the image, merged-cell layout, caption, and Print Preview.

## Output files

- `photo_slot_map_verified.csv` - $($verified.Count) structurally supported rows.
- `photo_slot_map_review.csv` - $($review.Count) rows requiring human review.
- The original Excel workbook was not modified.
"@
Set-Content -LiteralPath $reportPath -Value $report -Encoding UTF8

[pscustomobject]@{
    VisibleWorksheets = $summaryRows.Count
    InputMapWorksheetNames = $mapSheetNames.Count
    InputMapRows = $mapRows.Count
    VerifiedRows = $verified.Count
    ReviewRows = $review.Count
    MissingMapSheets = (($missingMapSheets | Select-Object -ExpandProperty Worksheet) -join ', ')
    OutOfScopeRows = $outOfScope.Count
    VerifiedPath = $verifiedPath
    ReviewPath = $reviewPath
    SummaryPath = $reportPath
} | ConvertTo-Json
