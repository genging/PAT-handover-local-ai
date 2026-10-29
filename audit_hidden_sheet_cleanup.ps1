$ErrorActionPreference = 'Stop'

$diagnosticPath = 'D:\Huawei Handover\reports\hidden_audit_diagnostic.log'
try { Remove-Item -LiteralPath $diagnosticPath -Force -ErrorAction SilentlyContinue } catch {}
trap {
    $_ | Out-File -LiteralPath $diagnosticPath -Append -Encoding UTF8
    throw
}

$workbookPath = 'D:\Huawei Handover\PAT DOC_LP89M_2026 Capacity Expansion NW Modernization_EXP_BKK.xlsx'
$reportDir = 'D:\Huawei Handover\reports'
New-Item -ItemType Directory -Path $reportDir -Force | Out-Null

$auditPath = Join-Path $reportDir 'hidden_sheet_cleanup_audit.csv'
$depsPath = Join-Path $reportDir 'hidden_sheet_dependencies.csv'
$namesPath = Join-Path $reportDir 'defined_names_audit.csv'
$linksPath = Join-Path $reportDir 'external_links_audit.csv'
$summaryPath = Join-Path $reportDir 'hidden_sheet_cleanup_summary.md'

function Get-VisibilityLabel([int]$value) {
    if ($value -eq -1) { return 'VISIBLE' }
    if ($value -eq 2) { return 'VERY_HIDDEN' }
    return 'HIDDEN'
}

function Get-SheetRefPattern($sheetNames) {
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($name in ($sheetNames | Sort-Object Length -Descending)) {
        $escaped = [regex]::Escape([string]$name)
        [void]$parts.Add("(?:'$escaped'|$escaped)!")
    }
    if ($parts.Count -eq 0) { return '(?!)' }
    return '(?i)(?:' + ($parts -join '|') + ')'
}

function Get-ReferencedSheetNames([string]$formula, [hashtable]$knownByLower, [regex]$refRegex) {
    $found = New-Object System.Collections.Generic.List[string]
    foreach ($match in $refRegex.Matches($formula)) {
        $raw = $match.Value.TrimEnd('!')
        if ($raw.StartsWith("'")) { $raw = $raw.Substring(1, $raw.Length - 2).Replace("''", "'") }
        $key = $raw.ToLowerInvariant()
        if ($knownByLower.ContainsKey($key) -and -not $found.Contains($knownByLower[$key])) { [void]$found.Add($knownByLower[$key]) }
    }
    return @($found)
}

function Get-FormulaRecords($ws, [hashtable]$knownByLower, [regex]$refRegex) {
    $records = @()
    $formulaRange = $null
    try { $formulaRange = $ws.UsedRange.SpecialCells(-4123) } catch { return @() }
    foreach ($area in $formulaRange.Areas) {
        $rowCount = [int]$area.Rows.Count
        $colCount = [int]$area.Columns.Count
        $values = $area.Formula
        for ($r = 1; $r -le $rowCount; $r++) {
            for ($c = 1; $c -le $colCount; $c++) {
                $formula = $null
                if ($rowCount -eq 1 -and $colCount -eq 1) { $formula = $values } else { $formula = $values[$r, $c] }
                if ($null -eq $formula) { continue }
                $formulaText = [string]$formula
                if (-not $formulaText.StartsWith('=')) { continue }
                $cell = $area.Cells.Item($r, $c)
                $refs = Get-ReferencedSheetNames $formulaText $knownByLower $refRegex
                $records += [pscustomobject]@{
                    Worksheet = [string]$ws.Name
                    Cell = [string]$cell.Address($false, $false)
                    Formula = $formulaText
                    ReferencedWorksheets = ($refs -join '; ')
                    HasExternalReference = ($formulaText -match '\[')
                    HasIndirectReference = ($formulaText -match '(?i)INDIRECT|ADDRESS\s*\(')
                }
            }
        }
    }
    return $records
}

function Get-CountSafe($scriptBlock) {
    try { return [int](& $scriptBlock) } catch { return 0 }
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $null
$allSheets = @()
$sheetByLower = @{}
$formulaRecords = @()
$auditRows = @()
$dependencyRows = @()
$nameRows = @()
$linkRows = @()
$scanLimitations = New-Object System.Collections.Generic.List[string]

try {
    # UpdateLinks=0 and ReadOnly=$true. No Save operation is performed.
    $wb = $excel.Workbooks.Open($workbookPath, 0, $true)
    foreach ($ws in $wb.Worksheets) {
        $allSheets += $ws
        $sheetByLower[[string]$ws.Name.ToLowerInvariant()] = [string]$ws.Name
    }

    $refRegex = [regex](Get-SheetRefPattern $sheetByLower.Values)
    foreach ($ws in $allSheets) {
        $formulaRecords += Get-FormulaRecords $ws $sheetByLower $refRegex
    }

    foreach ($name in $wb.Names) {
        $nameText = [string]$name.Name
        $refersTo = ''
        try { $refersTo = [string]$name.RefersTo } catch {}
        $refs = Get-ReferencedSheetNames $refersTo $sheetByLower $refRegex
        $nameRows += [pscustomobject]@{
            Name = $nameText
            Visible = [bool]$name.Visible
            RefersTo = $refersTo
            ReferencedWorksheets = ($refs -join '; ')
            HasExternalReference = ($refersTo -match '\[')
            HasIndirectReference = ($refersTo -match '(?i)INDIRECT|ADDRESS\s*\(')
        }
    }

    $links = $wb.LinkSources(1)
    if ($null -ne $links) {
        $linkIndex = 0
        foreach ($link in @($links)) {
            $linkIndex++
            $linkRows += [pscustomobject]@{ LinkIndex = $linkIndex; LinkSource = [string]$link; LinkType = 'ExcelWorkbookLink'; Status = 'NOT_CHANGED_READ_ONLY_AUDIT' }
        }
    }

    foreach ($ws in $allSheets) {
        $visible = ([int]$ws.Visible -eq -1)
        $visibilityLabel = Get-VisibilityLabel ([int]$ws.Visible)
        $ur = $ws.UsedRange
        $formulaForSheet = @($formulaRecords | Where-Object { $_.Worksheet -eq $ws.Name })
        $incomingFormula = @($formulaRecords | Where-Object { $_.ReferencedWorksheets -split '; ' -contains [string]$ws.Name })
        $incomingNames = @($nameRows | Where-Object { $_.ReferencedWorksheets -split '; ' -contains [string]$ws.Name })
        $externalFormulaCount = @($formulaForSheet | Where-Object HasExternalReference).Count
        $indirectFormulaCount = @($formulaForSheet | Where-Object HasIndirectReference).Count
        $hyperlinkCount = Get-CountSafe { $ws.Hyperlinks.Count }
        $shapeCount = Get-CountSafe { $ws.Shapes.Count }
        $pictureCount = 0
        for ($i = 1; $i -le $shapeCount; $i++) {
            try { $shapeType = [int]$ws.Shapes.Item($i).Type; if ($shapeType -eq 13 -or $shapeType -eq 11) { $pictureCount++ } } catch { [void]$scanLimitations.Add("$($ws.Name): shape type scan incomplete") }
        }
        $chartCount = Get-CountSafe { $ws.ChartObjects().Count }
        $pivotCount = Get-CountSafe { $ws.PivotTables().Count }
        $queryTableCount = Get-CountSafe { $ws.QueryTables.Count }
        $listObjectCount = Get-CountSafe { $ws.ListObjects.Count }
        $formatConditionCount = Get-CountSafe { $ur.FormatConditions.Count }
        $validationType = 'NONE_OR_UNKNOWN'
        try { $validationType = [string]$ur.Validation.Type; if (-not $validationType) { $validationType = 'NONE_OR_UNKNOWN' } } catch {}
        $hyperlinkHasExternal = $false
        try { foreach ($h in $ws.Hyperlinks) { if ([string]$h.Address -match '^(https?|file|mailto):|\[') { $hyperlinkHasExternal = $true } } } catch {}
        $hasContent = ([int64]$ur.CountLarge -gt 1 -or [int]$ur.Row -gt 1 -or [int]$ur.Column -gt 1)
        $hasObjects = ($shapeCount -gt 0 -or $chartCount -gt 0 -or $pivotCount -gt 0 -or $queryTableCount -gt 0 -or $listObjectCount -gt 0 -or $hyperlinkCount -gt 0 -or $formatConditionCount -gt 0 -or $validationType -ne 'NONE_OR_UNKNOWN')
        $hasDependency = ($incomingFormula.Count -gt 0 -or $incomingNames.Count -gt 0 -or $formulaForSheet.Count -gt 0 -or $externalFormulaCount -gt 0 -or $indirectFormulaCount -gt 0)
        $riskFlags = New-Object System.Collections.Generic.List[string]
        if ($incomingFormula.Count -gt 0) { [void]$riskFlags.Add('INCOMING_FORMULA_REFERENCES') }
        if ($incomingNames.Count -gt 0) { [void]$riskFlags.Add('DEFINED_NAMES_REFERENCE_SHEET') }
        if ($formulaForSheet.Count -gt 0) { [void]$riskFlags.Add('SHEET_HAS_FORMULAS') }
        if ($externalFormulaCount -gt 0) { [void]$riskFlags.Add('EXTERNAL_FORMULA_REFERENCES') }
        if ($indirectFormulaCount -gt 0) { [void]$riskFlags.Add('INDIRECT_OR_ADDRESS_REFERENCE_REQUIRES_REVIEW') }
        if ($shapeCount -gt 0) { [void]$riskFlags.Add('SHAPES_OR_IMAGES') }
        if ($chartCount -gt 0) { [void]$riskFlags.Add('CHARTS') }
        if ($pivotCount -gt 0) { [void]$riskFlags.Add('PIVOTS') }
        if ($queryTableCount -gt 0) { [void]$riskFlags.Add('QUERY_TABLES') }
        if ($listObjectCount -gt 0) { [void]$riskFlags.Add('TABLES') }
        if ($hyperlinkCount -gt 0) { [void]$riskFlags.Add('HYPERLINKS') }
        if ($formatConditionCount -gt 0) { [void]$riskFlags.Add('CONDITIONAL_FORMATTING') }
        if ($validationType -ne 'NONE_OR_UNKNOWN') { [void]$riskFlags.Add('DATA_VALIDATION_PRESENT_OR_UNKNOWN') }
        if ($hyperlinkHasExternal) { [void]$riskFlags.Add('EXTERNAL_HYPERLINK') }

        $status = 'REVIEW_REQUIRED'
        $reason = 'Conservative default: hidden-sheet deletion is not approved without complete dependency proof.'
        $confidence = 'LOW'
        if ($visibilityLabel -eq 'VISIBLE') {
            $status = 'KEEP_REQUIRED'
            $reason = 'Visible Worksheet is out of scope for hidden-sheet cleanup.'
            $confidence = 'HIGH'
        } elseif (-not $hasContent -and -not $hasObjects -and -not $hasDependency) {
            $status = 'DELETE_SAFE'
            $reason = 'Hidden worksheet is empty by UsedRange and has no detected formulas, incoming references, names, objects, links, validation, or conditional formatting.'
            $confidence = 'MEDIUM'
        } elseif ($hasDependency) {
            $status = 'REPAIR_THEN_DELETE'
            $reason = 'Detected formula/name dependency or formula content. References must be repaired or proven unnecessary before deletion.'
            $confidence = 'MEDIUM'
        } elseif ($hasObjects -or $hasContent) {
            $status = 'REVIEW_REQUIRED'
            $reason = 'Hidden worksheet contains data or Excel objects but no complete proof was found that deletion is safe.'
            $confidence = 'MEDIUM'
        }

        $auditRows += [pscustomobject]@{
            WorksheetIndex = [int]$ws.Index
            Worksheet = [string]$ws.Name
            Visibility = $visibilityLabel
            UsedRange = [string]$ur.Address($false, $false)
            UsedRows = [int]$ur.Rows.Count
            UsedColumns = [int]$ur.Columns.Count
            UsedCellCount = [int64]$ur.CountLarge
            FormulaCount = $formulaForSheet.Count
            IncomingFormulaReferenceCount = $incomingFormula.Count
            DefinedNameReferenceCount = $incomingNames.Count
            ExternalFormulaCount = $externalFormulaCount
            IndirectFormulaCount = $indirectFormulaCount
            ShapeCount = $shapeCount
            PictureCount = $pictureCount
            ChartCount = $chartCount
            PivotTableCount = $pivotCount
            QueryTableCount = $queryTableCount
            ListObjectCount = $listObjectCount
            HyperlinkCount = $hyperlinkCount
            ExternalHyperlinkDetected = $hyperlinkHasExternal
            ConditionalFormattingCount = $formatConditionCount
            DataValidationStatus = $validationType
            IncomingFormulaSources = (($incomingFormula | Select-Object -ExpandProperty Worksheet -Unique) -join '; ')
            IncomingDefinedNames = (($incomingNames | Select-Object -ExpandProperty Name) -join '; ')
            RiskFlags = ($riskFlags -join '; ')
            ExpectedImpact = $reason
            RequiredRepair = if ($status -eq 'REPAIR_THEN_DELETE') { 'Trace every incoming formula/name and replace or retain the source before deletion.' } else { 'None identified at this stage; human review may still be required.' }
            ManagementStatus = $status
            Confidence = $confidence
            HumanReviewRequired = if ($status -eq 'DELETE_SAFE') { 'NO_FOR_CURRENT_SCAN; APPROVAL_STILL_REQUIRED' } else { 'YES' }
        }

        foreach ($ref in $incomingFormula) {
            $dependencyRows += [pscustomobject]@{ DependencyType = 'FORMULA_INCOMING'; TargetWorksheet = [string]$ws.Name; SourceWorksheet = $ref.Worksheet; SourceCell = $ref.Cell; ReferenceText = $ref.Formula; ReferenceKind = if ($ref.HasIndirectReference) { 'INDIRECT_OR_ADDRESS' } elseif ($ref.HasExternalReference) { 'EXTERNAL_FORMULA' } else { 'DIRECT_OR_PARSED_SHEET_REFERENCE' } }
        }
        foreach ($refName in $incomingNames) {
            $dependencyRows += [pscustomobject]@{ DependencyType = 'DEFINED_NAME_INCOMING'; TargetWorksheet = [string]$ws.Name; SourceWorksheet = '[Workbook Name]'; SourceCell = $refName.Name; ReferenceText = $refName.RefersTo; ReferenceKind = if ($refName.HasIndirectReference) { 'INDIRECT_OR_ADDRESS' } elseif ($refName.HasExternalReference) { 'EXTERNAL_REFERENCE' } else { 'DEFINED_NAME' } }
        }
    }
}
finally {
    if ($null -ne $wb) { $wb.Close($false); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) }
    $excel.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

$auditRows | Export-Csv -LiteralPath $auditPath -NoTypeInformation -Encoding UTF8
$dependencyRows | Export-Csv -LiteralPath $depsPath -NoTypeInformation -Encoding UTF8
$nameRows | Export-Csv -LiteralPath $namesPath -NoTypeInformation -Encoding UTF8
$linkRows | Export-Csv -LiteralPath $linksPath -NoTypeInformation -Encoding UTF8

$hidden = @($auditRows | Where-Object Visibility -ne 'VISIBLE')
$statusLines = foreach ($s in @('DELETE_SAFE','REPAIR_THEN_DELETE','KEEP_REQUIRED','REVIEW_REQUIRED')) { "| $s | $(@($hidden | Where-Object ManagementStatus -eq $s).Count) |" }
$hiddenList = ($hidden | Sort-Object WorksheetIndex | ForEach-Object { "| $($_.Worksheet) | $($_.Visibility) | $($_.ManagementStatus) | $($_.RiskFlags) |" }) -join "`n"
$summary = @"
# Hidden Sheet Cleanup Audit

## Scope and safety

- Source workbook: `$workbookPath`
- Opened with Excel Read-only and external-link updates disabled.
- All 189 Worksheets were checked for visibility; dependency inspection covered the 98 hidden/very-hidden sheets and their references.
- No worksheet was deleted, moved, saved, or modified.
- Hidden Worksheets were inspected for cleanup planning only and were not used as photo sources.

## Management classification

| Status | Hidden sheets |
|---|---:|
$($statusLines -join "`n")

No deletion has been performed. Even `DELETE_SAFE` rows require user approval before any copy is created or any sheet is removed.

## Dependency checks performed

- Formulas in all sheets, including incoming references to hidden sheets
- Workbook Defined Names and their worksheet references
- External workbook links reported by Excel
- Hyperlinks
- Shape/image, chart, PivotTable, QueryTable, and ListObject counts
- Conditional Formatting count and Data Validation presence/status
- Formula flags for external references and `INDIRECT`/`ADDRESS`
- VBA was not present in the `.xlsx` file format; no VBA project was changed

## Hidden worksheet classification

| Worksheet | Visibility | Status | Risk flags |
|---|---|---|---|
$hiddenList

## Important limitations

- Excel COM does not provide a complete, lossless dependency graph for every feature through this audit alone, especially PivotCaches, chart series, query/connection internals, indirect references, and workbook-level external-link semantics.
- Any sheet with such unresolved risk remains `REVIEW_REQUIRED` or `REPAIR_THEN_DELETE`; it is not treated as safe merely because it is hidden or visually empty.
- No formulas were rewritten and no links were removed.

## Output files

- `hidden_sheet_cleanup_audit.csv` — one row per worksheet, including all 98 hidden/very-hidden sheets
- `hidden_sheet_dependencies.csv` — incoming formula and Defined Name dependencies
- `defined_names_audit.csv` — workbook Defined Names and references
- `external_links_audit.csv` — external links reported by Excel

The original workbook remains unchanged. The process stops here pending review and approval.
"@
Set-Content -LiteralPath $summaryPath -Value $summary -Encoding UTF8

[pscustomobject]@{
    TotalWorksheets = $allSheets.Count
    HiddenWorksheets = $hidden.Count
    VisibleWorksheets = @($auditRows | Where-Object Visibility -eq 'VISIBLE').Count
    DeleteSafe = @($hidden | Where-Object ManagementStatus -eq 'DELETE_SAFE').Count
    RepairThenDelete = @($hidden | Where-Object ManagementStatus -eq 'REPAIR_THEN_DELETE').Count
    KeepRequired = @($hidden | Where-Object ManagementStatus -eq 'KEEP_REQUIRED').Count
    ReviewRequired = @($hidden | Where-Object ManagementStatus -eq 'REVIEW_REQUIRED').Count
    FormulaDependencyRows = $dependencyRows.Count
    DefinedNames = $nameRows.Count
    ExternalLinks = $linkRows.Count
    AuditPath = $auditPath
    DependenciesPath = $depsPath
    NamesPath = $namesPath
    LinksPath = $linksPath
    SummaryPath = $summaryPath
} | ConvertTo-Json
