$ErrorActionPreference = 'Stop'

$reportDir = 'D:\Huawei Handover\reports'
$mapPath = Join-Path $reportDir 'visible_worksheet_map.csv'
$summaryPath = Join-Path $reportDir 'visible_worksheet_summary.csv'
$readmePath = Join-Path $reportDir 'visible_worksheet_map_readme.md'

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

function Get-RowFromCell([string]$address) {
    $m = [regex]::Match($address, '\d+')
    if ($m.Success) { return [int]$m.Value }
    return $null
}

function Get-Signature($row) {
    return ('PrintArea={0};Pics={1};FitW={2};FitH={3};Orientation={4};HBreaks={5};VBreaks={6}' -f $row.PrintArea, $row.PictureCount, $row.FitToPagesWide, $row.FitToPagesTall, $row.Orientation, $row.HorizontalPageBreaks, $row.VerticalPageBreaks)
}

$map = @(Import-Csv -LiteralPath $mapPath)
$sheetSummary = @(Import-Csv -LiteralPath $summaryPath)
$readme = Get-Content -Raw -LiteralPath $readmePath

$summaryBySheet = @{}
foreach ($row in $sheetSummary) { $summaryBySheet[[string]$row.Worksheet] = $row }

$slotRows = @()
$reviewRows = @()
foreach ($row in $map) {
    $sheet = [string]$row.Worksheet
    $sheetRow = $summaryBySheet[$sheet]
    $category = Get-Category $sheet
    $topRow = Get-RowFromCell ([string]$row.TopLeftCell)
    $slotKind = 'physical_image_object'
    $candidateType = 'photo_area_candidate'
    if ($null -ne $topRow -and $topRow -le 4) {
        $candidateType = 'header_or_decoration_candidate'
    }
    $basis = 'IsPicture=True and cell anchors copied from visible_worksheet_map.csv'
    $slotRows += [pscustomobject]@{
        WorksheetIndex = $row.WorksheetIndex
        Worksheet = $sheet
        WorksheetCategory = $category
        ShapeIndex = $row.ShapeIndex
        ShapeName = $row.ShapeName
        TopLeftCell = $row.TopLeftCell
        BottomRightCell = $row.BottomRightCell
        Width = $row.Width
        Height = $row.Height
        CandidateType = $candidateType
        Caption = 'UNKNOWN'
        Activity = 'UNKNOWN'
        EquipmentType = 'UNKNOWN'
        Sector = 'UNKNOWN'
        ImageRole = 'UNKNOWN'
        EvidenceType = 'UNKNOWN'
        SlotStatus = 'PHYSICAL_IMAGE_OBJECT_CONFIRMED_SEMANTIC_ROLE_UNCONFIRMED'
        Confidence = 'HIGH_PHYSICAL_NONE_SEMANTIC'
        SourceBasis = $basis
        ReviewRequired = 'YES'
    }
    $reviewRows += [pscustomobject]@{
        WorksheetIndex = $row.WorksheetIndex
        Worksheet = $sheet
        WorksheetCategory = $category
        ShapeIndex = $row.ShapeIndex
        ShapeName = $row.ShapeName
        TopLeftCell = $row.TopLeftCell
        BottomRightCell = $row.BottomRightCell
        CandidateType = $candidateType
        CaptionStatus = 'MISSING_IN_CSV'
        ReviewReason = 'CSV confirms a picture object and coordinates, but provides no caption, activity, equipment, sector, SN, or Before/After role for this object.'
        RequiredExcelCheck = 'Open the visible worksheet in Excel and inspect the caption/merged cells, image content, and print preview before assigning semantic meaning.'
        SemanticFieldsUnknown = 'Activity;EquipmentType;Sector;ImageRole;EvidenceType;SN'
        Confidence = 'NOT_ENOUGH_DATA'
        RepeatedStructureSignature = Get-Signature $sheetRow
    }
}

$slotRows | Export-Csv -LiteralPath (Join-Path $reportDir 'photo_slot_map.csv') -NoTypeInformation -Encoding UTF8
$reviewRows | Export-Csv -LiteralPath (Join-Path $reportDir 'photo_slot_map_review.csv') -NoTypeInformation -Encoding UTF8

$photoCandidates = @($slotRows | Where-Object CandidateType -eq 'photo_area_candidate')
$headerCandidates = @($slotRows | Where-Object CandidateType -eq 'header_or_decoration_candidate')
$categories = @($sheetSummary | ForEach-Object { Get-Category $_.Worksheet } | Sort-Object -Unique)
$categoryLines = foreach ($category in $categories) {
    $rows = @($sheetSummary | Where-Object { (Get-Category $_.Worksheet) -eq $category })
    $pics = ($rows | Measure-Object -Property PictureCount -Sum).Sum
    $sheets = $rows.Count
    "| $category | $sheets | $pics |"
}
$repeated = @($sheetSummary | Group-Object { Get-Signature $_ } | Where-Object Count -gt 1 | Sort-Object Count -Descending)
$repeatedLines = foreach ($group in $repeated | Select-Object -First 15) {
    $names = (($group.Group | Select-Object -ExpandProperty Worksheet) -join ', ')
    "| $($group.Count) | $names | $($group.Name) |"
}
if ($repeatedLines.Count -eq 0) { $repeatedLines = '| 0 | No exact repeated signature | - |' }

$md = @"
# Photo Slot Map Summary - Visible Worksheets Only

This analysis used only the three report files specified for step 2. The Excel workbook was not opened or read, and Hidden Worksheets were not used.

## Results

| Item | Count |
|---|---:|
| Visible Worksheets from summary CSV | $($sheetSummary.Count) |
| Picture objects from map CSV | $($map.Count) |
| Physical image objects with confirmed CSV coordinates | $($slotRows.Count) |
| Photo-area candidates (TopLeft row > 4) | $($photoCandidates.Count) |
| Header/decoration candidates (TopLeft row <= 4) | $($headerCandidates.Count) |
| Semantic photo slots confirmed from CSV | 0 |
| Items requiring additional review | $($reviewRows.Count) |

Confirmed means only that the CSV records `IsPicture=True` and provides cell anchors and image dimensions. It does not confirm an Activity, Equipment, Sector, SN, or Before/After role.

## Worksheet groups found

| Group | Worksheets | Pictures |
|---|---:|---:|
$($categoryLines -join "`n")

## Exact structural signatures repeated in the CSV

| Worksheets | Example/member sheets | Signature |
|---:|---|---|
$($repeatedLines -join "`n")

The signature uses PrintArea, PictureCount, FitToPagesWide/High, Orientation, and page-break counts. It indicates structural similarity only; it does not confirm that slots have the same meaning.

## What can and cannot be determined

- Determined: worksheet-name groups, picture counts, TopLeft/BottomRight anchors, image dimensions, and a conservative split between top-row header/decorative candidates and lower-row photo-area candidates.
- Not determined from CSV: captions, merged-cell context, image content, Activity, Equipment Model, Sector, tape color, SN, and Before/After role.
- No coordinates were invented, no captions were filled in, and no picture was assigned to an activity.

## Output files

- `photo_slot_map.csv` - $($slotRows.Count) physical image objects with semantic fields set to `UNKNOWN` and separate physical/semantic confidence.
- `photo_slot_map_review.csv` - $($reviewRows.Count) items requiring additional Excel review.

## Required follow-up checks

Open only Visible Worksheets in Excel to read captions/merged cells and verify Print Preview before creating the next semantic slot map, especially repeated image layouts in Antenna, RRU/AAU, SN, Card, DCDU, and Dismantle groups.

The original Excel workbook was not modified in step 2.
"@
Set-Content -LiteralPath (Join-Path $reportDir 'photo_slot_map_summary.md') -Value $md -Encoding UTF8

[pscustomobject]@{
    InputMapRows = $map.Count
    InputSummaryRows = $sheetSummary.Count
    PhysicalImageObjects = $slotRows.Count
    PhotoAreaCandidates = $photoCandidates.Count
    HeaderDecorationCandidates = $headerCandidates.Count
    SemanticSlotsConfirmed = 0
    ReviewRows = $reviewRows.Count
    OutputSlotMap = (Join-Path $reportDir 'photo_slot_map.csv')
    OutputReview = (Join-Path $reportDir 'photo_slot_map_review.csv')
    OutputSummary = (Join-Path $reportDir 'photo_slot_map_summary.md')
} | ConvertTo-Json
