$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$diagnosticPath = 'D:\Huawei Handover\reports\sheet_consolidation_diagnostic.log'
try { Remove-Item -LiteralPath $diagnosticPath -Force -ErrorAction SilentlyContinue } catch {}
trap { $_ | Out-File -LiteralPath $diagnosticPath -Append -Encoding UTF8; throw }

$workbookPath = 'D:\Huawei Handover\PAT DOC_LP89M_2026 Capacity Expansion NW Modernization_EXP_BKK_hidden_images_removed.xlsx'
$reportDir = 'D:\Huawei Handover\reports'
$planPath = Join-Path $reportDir 'sheet_consolidation_plan.csv'
$summaryPath = Join-Path $reportDir 'sheet_consolidation_summary.md'

function Read-ZipXml($zip, [string]$path) {
    $entry = $zip.GetEntry($path)
    if ($null -eq $entry) { return $null }
    $stream = $entry.Open()
    try { $doc = New-Object System.Xml.XmlDocument; $doc.Load($stream); return $doc } finally { $stream.Dispose() }
}

function Get-Attr($node, [string]$name) {
    if ($null -eq $node) { return '' }
    $a = $node.Attributes[$name]
    if ($null -eq $a) { foreach ($candidate in $node.Attributes) { if ($candidate.LocalName -eq $name) { $a = $candidate; break } } }
    if ($null -ne $a) { return [string]$a.Value }
    return ''
}

function Normalize-ZipPath([string]$path) {
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($part in ($path -split '/')) {
        if (-not $part -or $part -eq '.') { continue }
        if ($part -eq '..') { if ($parts.Count -gt 0) { $parts.RemoveAt($parts.Count - 1) }; continue }
        [void]$parts.Add($part)
    }
    return ($parts -join '/')
}

function Get-Hash([string]$text) {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($text))).Replace('-', '').ToLowerInvariant()) } finally { $sha.Dispose() }
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

function Get-CellValueText($cell) {
    $inline = $cell.SelectSingleNode("*[local-name()='is']//*[local-name()='t']")
    if ($null -ne $inline) { return [string]$inline.InnerText }
    $v = $cell.SelectSingleNode("*[local-name()='v']")
    if ($null -ne $v) { return [string]$v.InnerText }
    return ''
}

$zip = [IO.Compression.ZipFile]::OpenRead($workbookPath)
$planRows = @()
try {
    $wb = Read-ZipXml $zip 'xl/workbook.xml'
    $rels = Read-ZipXml $zip 'xl/_rels/workbook.xml.rels'
    $relTargets = @{}
    foreach ($rel in $rels.SelectNodes("//*[local-name()='Relationship']")) { $relTargets[(Get-Attr $rel 'Id')] = Get-Attr $rel 'Target' }

    $sheetInfo = @()
    foreach ($sheet in $wb.SelectNodes("//*[local-name()='sheet']")) {
        $state = Get-Attr $sheet 'state'; if (-not $state) { $state = 'visible' }
        $target = Normalize-ZipPath ('xl/' + $relTargets[(Get-Attr $sheet 'id')])
        $sheetInfo += [pscustomobject]@{ Index = $sheetInfo.Count; Name = Get-Attr $sheet 'name'; State = $state.ToUpperInvariant(); Path = $target }
    }
    $knownByLower = @{}
    foreach ($s in $sheetInfo) { $knownByLower[$s.Name.ToLowerInvariant()] = $s.Name }
    $patternParts = New-Object Collections.Generic.List[string]
    foreach ($name in ($knownByLower.Values | Sort-Object Length -Descending)) {
        $e = [regex]::Escape([string]$name)
        [void]$patternParts.Add("(?:'$e'|$e)!")
    }
    $refRegex = [regex]('(?i)(?:' + ($patternParts -join '|') + ')')

    $definedNames = @()
    foreach ($dn in $wb.SelectNodes("//*[local-name()='definedName']")) {
        $refs = Get-ReferencedSheets ([string]$dn.InnerText) $knownByLower $refRegex
        $definedNames += [pscustomobject]@{ Name = Get-Attr $dn 'name'; LocalSheetId = Get-Attr $dn 'localSheetId'; RefersTo = [string]$dn.InnerText; ReferencedWorksheets = ($refs -join '; '); IsPrintArea = ((Get-Attr $dn 'name') -match '(?i)Print_Area'); IsPrintTitles = ((Get-Attr $dn 'name') -match '(?i)Print_Titles'); HasExternalReference = ([string]$dn.InnerText -match '\['); HasIndirectReference = ([string]$dn.InnerText -match '(?i)INDIRECT|ADDRESS\s*\(') }
    }

    $records = @()
    $formulaRecords = @()
    foreach ($s in $sheetInfo) {
        $doc = Read-ZipXml $zip $s.Path
        if ($null -eq $doc) { continue }
        $cells = @($doc.SelectNodes("//*[local-name()='c']"))
        $formulaNodes = @($doc.SelectNodes("//*[local-name()='f']"))
        $formulaPattern = New-Object Collections.Generic.List[string]
        $valuePattern = New-Object Collections.Generic.List[string]
        $stylePattern = New-Object Collections.Generic.List[string]
        $outgoing = New-Object Collections.Generic.List[string]
        $externalFormulaCount = 0
        $indirectCount = 0
        foreach ($cell in $cells) {
            $addr = Get-Attr $cell 'r'
            $style = Get-Attr $cell 's'
            if ($style) { [void]$stylePattern.Add($style) }
            $f = $cell.SelectSingleNode("*[local-name()='f']")
            if ($null -ne $f) {
                $formula = [string]$f.InnerText
                [void]$formulaPattern.Add("$addr=$formula")
                $refs = Get-ReferencedSheets $formula $knownByLower $refRegex
                foreach ($ref in $refs) { if (-not $outgoing.Contains($ref)) { [void]$outgoing.Add($ref) } }
                if ($formula -match '\[') { $externalFormulaCount++ }
                if ($formula -match '(?i)INDIRECT|ADDRESS\s*\(') { $indirectCount++ }
                $formulaRecords += [pscustomobject]@{ Worksheet = $s.Name; Cell = $addr; Formula = $formula; ReferencedWorksheets = ($refs -join '; '); HasExternalReference = ($formula -match '\['); HasIndirectReference = ($formula -match '(?i)INDIRECT|ADDRESS\s*\(') }
            } else {
                $v = Get-CellValueText $cell
                [void]$valuePattern.Add("$addr=$v")
            }
        }
        $localNames = @($definedNames | Where-Object { $_.LocalSheetId -ne '' -and [int]$_.LocalSheetId -eq $s.Index })
        $sheetNames = @($definedNames | Where-Object { $_.ReferencedWorksheets -split '; ' -contains $s.Name })
        $records += [pscustomobject]@{
            Index = $s.Index
            Worksheet = $s.Name
            Visibility = $s.State
            SheetXmlPath = $s.Path
            Dimension = ((@($doc.SelectNodes("//*[local-name()='dimension']")) | Select-Object -First 1 | ForEach-Object { Get-Attr $_ 'ref' }))
            CellCount = $cells.Count
            FormulaCount = $formulaNodes.Count
            FormulaPatternHash = Get-Hash (($formulaPattern | Sort-Object) -join '|')
            ValuePatternHash = Get-Hash (($valuePattern | Sort-Object) -join '|')
            StylePatternHash = Get-Hash ((@($stylePattern | Group-Object | Sort-Object Name | ForEach-Object { "$($_.Name):$($_.Count)" })) -join '|')
            HyperlinkCount = @($doc.SelectNodes("//*[local-name()='hyperlink']")).Count
            DataValidationCount = @($doc.SelectNodes("//*[local-name()='dataValidation']")).Count
            ConditionalFormattingCount = @($doc.SelectNodes("//*[local-name()='conditionalFormatting']")).Count
            DrawingCount = @($doc.SelectNodes("//*[local-name()='drawing' or local-name()='legacyDrawing']")).Count
            TablePartCount = @($doc.SelectNodes("//*[local-name()='tablePart']")).Count
            PivotTablePartCount = @($doc.SelectNodes("//*[local-name()='pivotTablePart']")).Count
            LocalNameCount = $localNames.Count
            ReferencedNameCount = $sheetNames.Count
            PrintAreaNames = (($sheetNames | Where-Object IsPrintArea | Select-Object -ExpandProperty Name) -join '; ')
            PrintTitleNames = (($sheetNames | Where-Object IsPrintTitles | Select-Object -ExpandProperty Name) -join '; ')
            FormulaExternalCount = $externalFormulaCount
            FormulaIndirectCount = $indirectCount
            OutgoingWorksheets = ($outgoing -join '; ')
        }
    }

    foreach ($record in $records) {
        $visible = @($records | Where-Object Visibility -eq 'VISIBLE')
        $hidden = ($record.Visibility -ne 'VISIBLE')
        $formulaSources = @($formulaRecords | Where-Object { ($_.ReferencedWorksheets -split '; ') -contains $record.Worksheet })
        $nameSources = @($definedNames | Where-Object { ($_.ReferencedWorksheets -split '; ') -contains $record.Worksheet })
        $exactCandidates = @($visible | Where-Object { $_.CellCount -eq $record.CellCount -and $_.Dimension -eq $record.Dimension -and $_.FormulaPatternHash -eq $record.FormulaPatternHash -and $_.ValuePatternHash -eq $record.ValuePatternHash })
        $formulaCandidates = @($visible | Where-Object { $_.CellCount -eq $record.CellCount -and $_.Dimension -eq $record.Dimension -and $_.FormulaPatternHash -eq $record.FormulaPatternHash })
        $candidate = $null
        $matchType = 'NONE'
        if ($exactCandidates.Count -gt 0) { $candidate = $exactCandidates[0]; $matchType = 'EXACT_CONTENT_AND_FORMULA_PATTERN' }
        elseif ($formulaCandidates.Count -gt 0) { $candidate = $formulaCandidates[0]; $matchType = 'FORMULA_PATTERN_AND_STRUCTURE_ONLY' }
        $differences = New-Object Collections.Generic.List[string]
        if ($candidate) {
            if ($candidate.ValuePatternHash -ne $record.ValuePatternHash) { [void]$differences.Add('CELL_VALUES_DIFFER') }
            if ($candidate.StylePatternHash -ne $record.StylePatternHash) { [void]$differences.Add('STYLE_DISTRIBUTION_DIFFERS') }
            if ($candidate.PrintAreaNames -ne $record.PrintAreaNames) { [void]$differences.Add('PRINT_AREA_NAMES_DIFFER') }
            if ($candidate.PrintTitleNames -ne $record.PrintTitleNames) { [void]$differences.Add('PRINT_TITLE_NAMES_DIFFER') }
            if ($candidate.HyperlinkCount -ne $record.HyperlinkCount) { [void]$differences.Add('HYPERLINK_COUNT_DIFFERS') }
            if ($candidate.DataValidationCount -ne $record.DataValidationCount) { [void]$differences.Add('DATA_VALIDATION_COUNT_DIFFERS') }
            if ($candidate.ConditionalFormattingCount -ne $record.ConditionalFormattingCount) { [void]$differences.Add('CONDITIONAL_FORMATTING_COUNT_DIFFERS') }
            if ($candidate.DrawingCount -ne $record.DrawingCount) { [void]$differences.Add('DRAWING_COUNT_DIFFERS') }
        }
        $status = 'KEEP'
        $risk = New-Object Collections.Generic.List[string]
        if (-not $hidden) { $status = 'KEEP'; $confidence = 'HIGH'; $reason = 'Visible sheet remains the user-facing candidate and is not changed in this read-only plan.' }
        else {
            if ($formulaSources.Count -gt 0) { [void]$risk.Add('INCOMING_FORMULA_REFERENCES') }
            if ($nameSources.Count -gt 0) { [void]$risk.Add('DEFINED_NAME_REFERENCES') }
            if ($record.FormulaExternalCount -gt 0) { [void]$risk.Add('EXTERNAL_FORMULAS') }
            if ($record.FormulaIndirectCount -gt 0) { [void]$risk.Add('INDIRECT_OR_ADDRESS') }
            if ($record.DataValidationCount -gt 0) { [void]$risk.Add('DATA_VALIDATION') }
            if ($record.ConditionalFormattingCount -gt 0) { [void]$risk.Add('CONDITIONAL_FORMATTING') }
            if ($record.DrawingCount -gt 0) { [void]$risk.Add('DRAWINGS') }
            if ($record.HyperlinkCount -gt 0) { [void]$risk.Add('HYPERLINKS') }
            if ($matchType -eq 'EXACT_CONTENT_AND_FORMULA_PATTERN' -and $differences.Count -eq 0) {
                $status = 'MERGE_AND_RELINK'; $confidence = 'MEDIUM'; $reason = 'Exact content/formula/structure match to a Visible Sheet; relink incoming formulas and Defined Names before retiring the hidden copy.'
            } elseif ($matchType -eq 'FORMULA_PATTERN_AND_STRUCTURE_ONLY') {
                $status = 'REVIEW_REQUIRED'; $confidence = 'LOW'; $reason = 'Formula pattern and dimensions match a Visible Sheet, but cell values or other structures differ.'
                [void]$risk.Add('CONTENT_DIFFERENCE_REQUIRES_HUMAN_REVIEW')
            } elseif ($formulaSources.Count -gt 0 -or $nameSources.Count -gt 0) {
                $status = 'DELETE_AFTER_RELINK'; $confidence = 'LOW'; $reason = 'No verified Visible duplicate was found, but incoming references exist; a valid replacement source must be selected before deletion.'
                [void]$risk.Add('NO_VERIFIED_VISIBLE_REPLACEMENT')
            } else {
                $status = 'REVIEW_REQUIRED'; $confidence = 'LOW'; $reason = 'Hidden sheet contains content or structures without a verified Visible replacement.'
            }
        }
        $recordsByName = $records | Where-Object Worksheet -eq $record.Worksheet
        $planRows += [pscustomobject]@{
            SourceWorksheet = $record.Worksheet
            SourceVisibility = $record.Visibility
            ProposedVisibleTarget = if ($candidate) { $candidate.Worksheet } else { '' }
            MatchType = $matchType
            SourceDimension = $record.Dimension
            TargetDimension = if ($candidate) { $candidate.Dimension } else { '' }
            SourceCellCount = $record.CellCount
            TargetCellCount = if ($candidate) { $candidate.CellCount } else { '' }
            SourceFormulaCount = $record.FormulaCount
            TargetFormulaCount = if ($candidate) { $candidate.FormulaCount } else { '' }
            DataAndFormulaRange = $record.Dimension
            FormulaReferencesToRewrite = (($formulaSources | ForEach-Object { "$($_.Worksheet)!$($_.Cell)" }) -join '; ')
            DefinedNamesToReview = (($nameSources | Select-Object -ExpandProperty Name) -join '; ')
            PrintAreasAffected = $record.PrintAreaNames
            PrintTitlesAffected = $record.PrintTitleNames
            HyperlinkCount = $record.HyperlinkCount
            DataValidationCount = $record.DataValidationCount
            ConditionalFormattingCount = $record.ConditionalFormattingCount
            DrawingCountAfterImageCleanup = $record.DrawingCount
            PivotTablePartCount = $record.PivotTablePartCount
            OutgoingWorksheetReferences = $record.OutgoingWorksheets
            Differences = (($differences | Select-Object -Unique) -join '; ')
            RiskFlags = (($risk | Select-Object -Unique) -join '; ')
            Recommendation = $status
            Confidence = $confidence
            Reason = $reason
            HumanReviewRequired = if ($hidden) { 'YES' } else { 'NO_FOR_THIS_PLAN' }
        }
    }
}
finally { $zip.Dispose() }

$planRows | Export-Csv -LiteralPath $planPath -NoTypeInformation -Encoding UTF8
$hiddenPlan = @($planRows | Where-Object SourceVisibility -ne 'VISIBLE')
$exact = @($hiddenPlan | Where-Object MatchType -eq 'EXACT_CONTENT_AND_FORMULA_PATTERN')
$lowRisk = @($exact | Where-Object { $_.RiskFlags -notmatch 'INCOMING_FORMULA_REFERENCES|EXTERNAL|INDIRECT|DATA_VALIDATION|CONDITIONAL|HYPERLINK' })
$statusLines = foreach ($s in @('KEEP','MERGE_AND_RELINK','DELETE_AFTER_RELINK','REVIEW_REQUIRED')) { "| $s | $(@($planRows | Where-Object Recommendation -eq $s).Count) |" }
$candidateLines = if ($lowRisk.Count -gt 0) { ($lowRisk | Select-Object -First 20 | ForEach-Object { "- ``$($_.SourceWorksheet)`` -> ``$($_.ProposedVisibleTarget)`` ($($_.MatchType))" }) -join "`n" } else { '- None identified by the read-only structural scan.' }
$summary = @(
    '# Sheet Consolidation Plan',
    '',
    "Source workbook: $workbookPath",
    'Read-only Open XML analysis only. No formulas, links, values, formatting, worksheets, or external links were changed.',
    '',
    '## Recommendation counts',
    '',
    '| Recommendation | Sheets |',
    '|---|---:|',
    ($statusLines -join "`n"),
    '',
    '## Lowest-risk trial candidates',
    '',
    $candidateLines,
    '',
    'A candidate is low-risk only when the scan found an exact visible content/formula/structure match and no direct incoming formula, external, indirect, validation, conditional-formatting, or hyperlink risk. These are candidates for a separate test-copy experiment, not approval for deletion.',
    '',
    '## Constraints and limitations',
    '',
    '- Formula results were not recalculated.',
    '- External link targets were not opened or validated; they remain in the workbook.',
    '- Pivot caches, chart series, query providers, connection credentials, and VBA runtime behavior require Excel-side validation before any merge or delete.',
    '- A matching formula pattern does not prove matching business meaning; differences are retained in the CSV.',
    '- No source/target cell range was moved and no reference was rewritten.',
    '',
    'The process stops here pending approval of a specific trial group.'
)
Set-Content -LiteralPath $summaryPath -Value ($summary -join [Environment]::NewLine) -Encoding UTF8

[pscustomobject]@{
    TotalRows = $planRows.Count
    HiddenRows = $hiddenPlan.Count
    VisibleRows = @($planRows | Where-Object SourceVisibility -eq 'VISIBLE').Count
    ExactMatches = $exact.Count
    LowRiskTrialCandidates = $lowRisk.Count
    PlanPath = $planPath
    SummaryPath = $summaryPath
} | ConvertTo-Json
