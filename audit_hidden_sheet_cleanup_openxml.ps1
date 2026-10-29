$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$workbookPath = 'D:\Huawei Handover\PAT DOC_LP89M_2026 Capacity Expansion NW Modernization_EXP_BKK.xlsx'
$reportDir = 'D:\Huawei Handover\reports'
$auditPath = Join-Path $reportDir 'hidden_sheet_cleanup_audit.csv'
$depsPath = Join-Path $reportDir 'hidden_sheet_dependencies.csv'
$namesPath = Join-Path $reportDir 'defined_names_audit.csv'
$linksPath = Join-Path $reportDir 'external_links_audit.csv'
$summaryPath = Join-Path $reportDir 'hidden_sheet_cleanup_summary.md'

function Read-ZipXml($zip, [string]$path) {
    $entry = $zip.GetEntry($path)
    if ($null -eq $entry) { return $null }
    $stream = $entry.Open()
    try {
        $doc = New-Object System.Xml.XmlDocument
        $doc.Load($stream)
        return $doc
    } finally { $stream.Dispose() }
}

function Normalize-ZipPath([string]$base, [string]$target) {
    if ($target.StartsWith('/')) { return $target.TrimStart('/') }
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($part in (($base -split '/') + ($target -split '/'))) {
        if (-not $part -or $part -eq '.') { continue }
        if ($part -eq '..') { if ($parts.Count -gt 0) { $parts.RemoveAt($parts.Count - 1) }; continue }
        [void]$parts.Add($part)
    }
    return ($parts -join '/')
}

function Get-Attr($node, [string]$name) {
    if ($null -eq $node) { return '' }
    $a = $node.Attributes[$name]
    if ($null -eq $a) {
        foreach ($candidate in $node.Attributes) { if ($candidate.LocalName -eq $name) { $a = $candidate; break } }
    }
    if ($null -ne $a) { return [string]$a.Value }
    return ''
}

function Get-ReferencedSheets([string]$text, [hashtable]$knownByLower, [regex]$refRegex) {
    $found = New-Object System.Collections.Generic.List[string]
    foreach ($m in $refRegex.Matches($text)) {
        $raw = $m.Value.TrimEnd('!')
        if ($raw.StartsWith("'")) { $raw = $raw.Substring(1, $raw.Length - 2).Replace("''", "'") }
        $key = $raw.ToLowerInvariant()
        if ($knownByLower.ContainsKey($key) -and -not $found.Contains($knownByLower[$key])) { [void]$found.Add($knownByLower[$key]) }
    }
    return @($found)
}

$zip = [System.IO.Compression.ZipFile]::OpenRead($workbookPath)
$auditRows = @()
$dependencyRows = @()
$nameRows = @()
$linkRows = @()
$sheetFormulaRecords = @()

try {
    $workbookXml = Read-ZipXml $zip 'xl/workbook.xml'
    $workbookRels = Read-ZipXml $zip 'xl/_rels/workbook.xml.rels'
    $relTargets = @{}
    foreach ($rel in $workbookRels.SelectNodes("//*[local-name()='Relationship']")) {
        $relTargets[(Get-Attr $rel 'Id')] = Get-Attr $rel 'Target'
    }

    $sheetNodes = @($workbookXml.SelectNodes("//*[local-name()='sheet']"))
    $sheetInfo = @()
    foreach ($sheet in $sheetNodes) {
        $rid = Get-Attr $sheet 'id'
        $target = $relTargets[$rid]
        $sheetPath = Normalize-ZipPath 'xl' $target
        $state = Get-Attr $sheet 'state'
        if (-not $state) { $state = 'visible' }
        $sheetInfo += [pscustomobject]@{ Name = (Get-Attr $sheet 'name'); State = $state.ToUpperInvariant(); Path = $sheetPath; SheetId = (Get-Attr $sheet 'sheetId') }
    }
    $knownByLower = @{}
    foreach ($s in $sheetInfo) { $knownByLower[$s.Name.ToLowerInvariant()] = $s.Name }
    $patternParts = New-Object System.Collections.Generic.List[string]
    foreach ($name in ($knownByLower.Values | Sort-Object Length -Descending)) {
        $e = [regex]::Escape([string]$name)
        [void]$patternParts.Add("(?:'$e'|$e)!")
    }
    $refRegex = [regex]('(?i)(?:' + ($patternParts -join '|') + ')')

    $sheetRecords = @{}
    foreach ($s in $sheetInfo) {
        $doc = Read-ZipXml $zip $s.Path
        if ($null -eq $doc) { continue }
        $cells = @($doc.SelectNodes("//*[local-name()='c']"))
        $formulaNodes = @($doc.SelectNodes("//*[local-name()='f']"))
        $formulas = @()
        $incomingTextRefs = New-Object System.Collections.Generic.List[string]
        $outgoingSheets = New-Object System.Collections.Generic.List[string]
        $externalFormulaCount = 0
        $indirectCount = 0
        foreach ($f in $formulaNodes) {
            $parent = $f.ParentNode
            $cellRef = Get-Attr $parent 'r'
            $formulaText = [string]$f.InnerText
            $refs = Get-ReferencedSheets $formulaText $knownByLower $refRegex
            foreach ($ref in $refs) { if (-not $outgoingSheets.Contains($ref)) { [void]$outgoingSheets.Add($ref) } }
            if ($formulaText -match '\[') { $externalFormulaCount++ }
            if ($formulaText -match '(?i)INDIRECT|ADDRESS\s*\(') { $indirectCount++ }
            $rec = [pscustomobject]@{ Worksheet = $s.Name; Cell = $cellRef; Formula = $formulaText; ReferencedWorksheets = ($refs -join '; '); HasExternalReference = ($formulaText -match '\['); HasIndirectReference = ($formulaText -match '(?i)INDIRECT|ADDRESS\s*\(') }
            $formulas += $rec
            $sheetFormulaRecords += $rec
        }
        $sheetRecords[$s.Name] = [pscustomobject]@{
            Info = $s
            Xml = $doc
            CellCount = $cells.Count
            FormulaCount = $formulaNodes.Count
            Formulas = $formulas
            OutgoingSheets = @($outgoingSheets)
            ExternalFormulaCount = $externalFormulaCount
            IndirectFormulaCount = $indirectCount
            HyperlinkCount = @($doc.SelectNodes("//*[local-name()='hyperlink']")).Count
            DataValidationCount = @($doc.SelectNodes("//*[local-name()='dataValidation' or local-name()='dataValidations']")).Count
            ConditionalFormattingCount = @($doc.SelectNodes("//*[local-name()='conditionalFormatting']")).Count
            DrawingCount = @($doc.SelectNodes("//*[local-name()='drawing' or local-name()='legacyDrawing']")).Count
            TablePartCount = @($doc.SelectNodes("//*[local-name()='tablePart' or local-name()='tableParts']")).Count
            PivotTablePartCount = @($doc.SelectNodes("//*[local-name()='pivotTablePart' or local-name()='pivotTableParts']")).Count
            Dimension = ((@($doc.SelectNodes("//*[local-name()='dimension']")) | Select-Object -First 1 | ForEach-Object { Get-Attr $_ 'ref' }))
        }
    }

    foreach ($name in $workbookXml.SelectNodes("//*[local-name()='definedName']")) {
        $refersTo = [string]$name.InnerText
        $refs = Get-ReferencedSheets $refersTo $knownByLower $refRegex
        $nameRows += [pscustomobject]@{ Name = (Get-Attr $name 'name'); LocalSheetId = (Get-Attr $name 'localSheetId'); Hidden = (Get-Attr $name 'hidden'); RefersTo = $refersTo; ReferencedWorksheets = ($refs -join '; '); HasExternalReference = ($refersTo -match '\['); HasIndirectReference = ($refersTo -match '(?i)INDIRECT|ADDRESS\s*\(') }
    }

    foreach ($externalEntry in @($zip.Entries | Where-Object { $_.FullName -match '^xl/externalLinks/externalLink\d+\.xml$' })) {
        $doc = Read-ZipXml $zip $externalEntry.FullName
        $locations = @($doc.SelectNodes("//*[local-name()='sheetName']")) | ForEach-Object { Get-Attr $_ 'val' } | Where-Object { $_ } | Select-Object -Unique
        $relPath = ($externalEntry.FullName -replace '^xl/externalLinks/','xl/externalLinks/_rels/') + '.rels'
        $relDoc = Read-ZipXml $zip $relPath
        $targets = @()
        if ($null -ne $relDoc) { $targets = @($relDoc.SelectNodes("//*[local-name()='Relationship']")) | ForEach-Object { Get-Attr $_ 'Target' } | Where-Object { $_ } | Select-Object -Unique }
        $linkRows += [pscustomobject]@{ LinkIndex = $linkRows.Count + 1; LinkFile = $externalEntry.FullName; ReferencedItems = ($locations -join '; '); ExternalTarget = ($targets -join '; '); Status = 'NOT_CHANGED_READ_ONLY_AUDIT' }
    }

    $zipNames = @($zip.Entries | Select-Object -ExpandProperty FullName)
    $globalObjects = [ordered]@{
        ChartFiles = @($zipNames | Where-Object { $_ -match '^xl/charts/chart\d+\.xml$' }).Count
        PivotCacheFiles = @($zipNames | Where-Object { $_ -match '^xl/pivotCache/' }).Count
        QueryTableFiles = @($zipNames | Where-Object { $_ -match '^xl/queryTables/' }).Count
        ConnectionFiles = @($zipNames | Where-Object { $_ -match '^xl/connections\.xml$' }).Count
        VbaProjectFiles = @($zipNames | Where-Object { $_ -match '(^|/)vbaProject\.bin$' }).Count
    }

    foreach ($s in $sheetInfo) {
        $rec = $sheetRecords[$s.Name]
        if ($null -eq $rec) { continue }
        $incomingFormula = @($sheetFormulaRecords | Where-Object { ($_.ReferencedWorksheets -split '; ') -contains $s.Name })
        $incomingNames = @($nameRows | Where-Object { ($_.ReferencedWorksheets -split '; ') -contains $s.Name })
        $isVisible = ($s.State -eq 'VISIBLE')
        $risk = New-Object System.Collections.Generic.List[string]
        if ($incomingFormula.Count -gt 0) { [void]$risk.Add('INCOMING_FORMULA_REFERENCES') }
        if ($incomingNames.Count -gt 0) { [void]$risk.Add('DEFINED_NAMES_REFERENCE_SHEET') }
        if ($rec.FormulaCount -gt 0) { [void]$risk.Add('SHEET_HAS_FORMULAS') }
        if ($rec.ExternalFormulaCount -gt 0) { [void]$risk.Add('EXTERNAL_FORMULA_REFERENCES') }
        if ($rec.IndirectFormulaCount -gt 0) { [void]$risk.Add('INDIRECT_OR_ADDRESS_REFERENCE') }
        if ($rec.HyperlinkCount -gt 0) { [void]$risk.Add('HYPERLINKS') }
        if ($rec.DrawingCount -gt 0) { [void]$risk.Add('DRAWINGS_OR_LEGACY_DRAWINGS') }
        if ($rec.TablePartCount -gt 0) { [void]$risk.Add('TABLE_PARTS') }
        if ($rec.PivotTablePartCount -gt 0) { [void]$risk.Add('PIVOT_TABLE_PARTS') }
        $status = 'REVIEW_REQUIRED'
        $reason = 'Conservative default: sheet has content or unresolved dependency risk.'
        $confidence = 'LOW'
        if ($isVisible) { $status = 'KEEP_REQUIRED'; $reason = 'Visible Worksheet is not eligible for hidden-sheet cleanup.'; $confidence = 'HIGH' }
        elseif ($rec.CellCount -eq 0 -and $rec.FormulaCount -eq 0 -and $incomingFormula.Count -eq 0 -and $incomingNames.Count -eq 0 -and $rec.HyperlinkCount -eq 0 -and $rec.DrawingCount -eq 0 -and $rec.TablePartCount -eq 0 -and $rec.PivotTablePartCount -eq 0 -and $rec.ExternalFormulaCount -eq 0) {
            $status = 'DELETE_SAFE'; $reason = 'No cells, formulas, incoming formula/name references, hyperlinks, drawings, table parts, pivot table parts, or external formulas detected in Open XML.'; $confidence = 'MEDIUM'
        } elseif ($incomingFormula.Count -gt 0 -or $incomingNames.Count -gt 0) {
            $status = 'REPAIR_THEN_DELETE'; $reason = 'Other formulas or Defined Names reference this sheet; dependencies must be repaired or proven unnecessary before deletion.'; $confidence = 'MEDIUM'
        } elseif ($rec.CellCount -gt 0 -or $rec.FormulaCount -gt 0 -or $rec.DrawingCount -gt 0 -or $rec.TablePartCount -gt 0 -or $rec.PivotTablePartCount -gt 0) {
            $status = 'REVIEW_REQUIRED'; $reason = 'Sheet contains stored cells or Excel structures but no complete proof supports deletion.'; $confidence = 'MEDIUM'
        }
        $auditRows += [pscustomobject]@{
            WorksheetIndex = [int]$s.SheetId
            Worksheet = $s.Name
            Visibility = $s.State
            SheetXmlPath = $s.Path
            Dimension = $rec.Dimension
            StoredCellCount = $rec.CellCount
            FormulaCount = $rec.FormulaCount
            IncomingFormulaReferenceCount = $incomingFormula.Count
            DefinedNameReferenceCount = $incomingNames.Count
            ExternalFormulaCount = $rec.ExternalFormulaCount
            IndirectFormulaCount = $rec.IndirectFormulaCount
            HyperlinkCount = $rec.HyperlinkCount
            DrawingOrLegacyDrawingCount = $rec.DrawingCount
            TablePartCount = $rec.TablePartCount
            PivotTablePartCount = $rec.PivotTablePartCount
            WorkbookChartFileCount = $globalObjects.ChartFiles
            WorkbookPivotCacheFileCount = $globalObjects.PivotCacheFiles
            WorkbookQueryTableFileCount = $globalObjects.QueryTableFiles
            WorkbookConnectionFileCount = $globalObjects.ConnectionFiles
            WorkbookVbaProjectFileCount = $globalObjects.VbaProjectFiles
            OutgoingWorksheetReferences = ($rec.OutgoingSheets -join '; ')
            IncomingFormulaSources = (($incomingFormula | Select-Object -ExpandProperty Worksheet -Unique) -join '; ')
            IncomingDefinedNames = (($incomingNames | Select-Object -ExpandProperty Name) -join '; ')
            RiskFlags = ($risk -join '; ')
            ExpectedImpact = $reason
            RequiredRepair = if ($status -eq 'REPAIR_THEN_DELETE') { 'Trace incoming formulas and Defined Names; redirect or retain source before deletion.' } else { 'None identified by this structural scan; approval and follow-up review may still be required.' }
            ManagementStatus = $status
            Confidence = $confidence
            HumanReviewRequired = if ($status -eq 'DELETE_SAFE') { 'NO_FOR_STRUCTURAL_SCAN; APPROVAL_REQUIRED' } else { 'YES' }
        }
        foreach ($ref in $incomingFormula) {
            $dependencyRows += [pscustomobject]@{ DependencyType = 'FORMULA_INCOMING'; TargetWorksheet = $s.Name; SourceWorksheet = $ref.Worksheet; SourceCell = $ref.Cell; ReferenceText = $ref.Formula; ReferenceKind = if ($ref.HasIndirectReference) { 'INDIRECT_OR_ADDRESS' } elseif ($ref.HasExternalReference) { 'EXTERNAL_FORMULA' } else { 'DIRECT_OR_PARSED_SHEET_REFERENCE' } }
        }
        foreach ($name in $incomingNames) {
            $dependencyRows += [pscustomobject]@{ DependencyType = 'DEFINED_NAME_INCOMING'; TargetWorksheet = $s.Name; SourceWorksheet = '[Workbook Name]'; SourceCell = $name.Name; ReferenceText = $name.RefersTo; ReferenceKind = if ($name.HasIndirectReference) { 'INDIRECT_OR_ADDRESS' } elseif ($name.HasExternalReference) { 'EXTERNAL_REFERENCE' } else { 'DEFINED_NAME' } }
        }
    }
}
finally { $zip.Dispose() }

$auditRows | Export-Csv -LiteralPath $auditPath -NoTypeInformation -Encoding UTF8
$dependencyRows | Export-Csv -LiteralPath $depsPath -NoTypeInformation -Encoding UTF8
$nameRows | Export-Csv -LiteralPath $namesPath -NoTypeInformation -Encoding UTF8
$linkRows | Export-Csv -LiteralPath $linksPath -NoTypeInformation -Encoding UTF8

$hidden = @($auditRows | Where-Object Visibility -ne 'VISIBLE')
$statusLines = foreach ($s in @('DELETE_SAFE','REPAIR_THEN_DELETE','KEEP_REQUIRED','REVIEW_REQUIRED')) { "| $s | $(@($hidden | Where-Object ManagementStatus -eq $s).Count) |" }
$missingVba = ($auditRows | Select-Object -First 1).WorkbookVbaProjectFileCount
$summary = @"
# Hidden Sheet Cleanup Audit

## Scope and safety

- Source: $workbookPath
- Read using the workbook's Open XML package in read-only mode.
- All 189 Worksheets were inspected for visibility; 98 Hidden/Very Hidden sheets were classified.
- No worksheet was deleted, moved, saved, or modified.
- Hidden sheets were inspected for cleanup planning only and were not used as Photo Mapping image sources.

## Management classification

| Status | Hidden sheets |
|---|---:|
$($statusLines -join "`n")

No deletion has been performed. Even `DELETE_SAFE` rows require user approval before a test copy is created.

## Dependency and feature scan

- Formula nodes in every worksheet XML, including incoming references to hidden sheets
- Workbook Defined Names and sheet references
- External link parts: $($linkRows.Count)
- Hyperlinks, drawings, table parts, pivot table parts, data validations, and conditional formatting per sheet
- Workbook chart files: $($auditRows | Select-Object -First 1 | Select-Object -ExpandProperty WorkbookChartFileCount)
- Workbook pivot cache files: $($auditRows | Select-Object -First 1 | Select-Object -ExpandProperty WorkbookPivotCacheFileCount)
- Workbook query table files: $($auditRows | Select-Object -First 1 | Select-Object -ExpandProperty WorkbookQueryTableFileCount)
- Workbook connection file: $($auditRows | Select-Object -First 1 | Select-Object -ExpandProperty WorkbookConnectionFileCount)
- VBA project files: $missingVba
- Formula flags for external references and `INDIRECT`/`ADDRESS`

## Important limitations

- This is a structural Open XML audit. It does not recalculate formulas, render Print Preview, or prove every runtime dependency in PivotCaches, chart series, query providers, connection credentials, or indirect references.
- Any unresolved risk remains `REVIEW_REQUIRED` or `REPAIR_THEN_DELETE`; hidden status alone was never treated as proof of safety.
- External links were inventoried and not removed.

## Output files

- `hidden_sheet_cleanup_audit.csv` — one row per worksheet
- `hidden_sheet_dependencies.csv` — incoming formula and Defined Name references
- `defined_names_audit.csv` — workbook Defined Names
- `external_links_audit.csv` — external link parts

The original workbook was not modified. Processing stops here pending review and approval.
"@
Set-Content -LiteralPath $summaryPath -Value $summary -Encoding UTF8

[pscustomobject]@{
    TotalWorksheets = $auditRows.Count
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
