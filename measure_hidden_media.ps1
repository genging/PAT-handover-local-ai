Add-Type -AssemblyName System.IO.Compression.FileSystem
$path = 'D:\Huawei Handover\PAT DOC_LP89M_2026 Capacity Expansion NW Modernization_EXP_BKK.xlsx'
$zip = [IO.Compression.ZipFile]::OpenRead($path)
function ReadXml($entry) {
    $s = $entry.Open()
    try { $x = New-Object Xml.XmlDocument; $x.Load($s); return $x } finally { $s.Dispose() }
}
function Attr($node, [string]$name) {
    $a = $node.Attributes[$name]
    if ($null -eq $a) { foreach ($q in $node.Attributes) { if ($q.LocalName -eq $name) { $a = $q; break } } }
    if ($null -ne $a) { return [string]$a.Value }
    return ''
}
function Normalize([string]$path) {
    $parts = New-Object Collections.Generic.List[string]
    foreach ($p in ($path -split '/')) {
        if (-not $p -or $p -eq '.') { continue }
        if ($p -eq '..') { if ($parts.Count -gt 0) { $parts.RemoveAt($parts.Count - 1) }; continue }
        [void]$parts.Add($p)
    }
    return ($parts -join '/')
}
try {
    $wb = ReadXml($zip.GetEntry('xl/workbook.xml'))
    $rels = ReadXml($zip.GetEntry('xl/_rels/workbook.xml.rels'))
    $relTargets = @{}
    foreach ($r in $rels.SelectNodes("//*[local-name()='Relationship']")) { $relTargets[(Attr $r 'Id')] = Attr $r 'Target' }
    $sheetMedia = @{}
    $sheetState = @{}
    foreach ($sheet in $wb.SelectNodes("//*[local-name()='sheet']")) {
        $name = Attr $sheet 'name'
        $state = Attr $sheet 'state'
        if (-not $state) { $state = 'visible' }
        $sheetState[$name] = $state
        $sheetPath = Normalize ('xl/' + $relTargets[(Attr $sheet 'id')])
        $sheetRels = $zip.GetEntry('xl/worksheets/_rels/' + [IO.Path]::GetFileName($sheetPath) + '.rels')
        $media = New-Object Collections.Generic.List[string]
        if ($null -ne $sheetRels) {
            $sr = ReadXml($sheetRels)
            foreach ($rel in $sr.SelectNodes("//*[local-name()='Relationship']")) {
                if ((Attr $rel 'Type') -notmatch '/drawing$') { continue }
                $drawingPath = Normalize ('xl/worksheets/' + (Attr $rel 'Target'))
                $drawingRels = $zip.GetEntry('xl/drawings/_rels/' + [IO.Path]::GetFileName($drawingPath) + '.rels')
                if ($null -eq $drawingRels) { continue }
                $dr = ReadXml($drawingRels)
                foreach ($mediaRel in $dr.SelectNodes("//*[local-name()='Relationship']")) {
                    if ((Attr $mediaRel 'Target') -match 'media/') { [void]$media.Add((Normalize ('xl/drawings/' + (Attr $mediaRel 'Target')))) }
                }
            }
        }
        $sheetMedia[$name] = @($media | Select-Object -Unique)
    }
    $hiddenMedia = New-Object Collections.Generic.HashSet[string]
    $visibleMedia = New-Object Collections.Generic.HashSet[string]
    foreach ($name in $sheetMedia.Keys) {
        foreach ($media in $sheetMedia[$name]) {
            if ($sheetState[$name] -eq 'visible') { [void]$visibleMedia.Add($media) } else { [void]$hiddenMedia.Add($media) }
        }
    }
    $onlyHidden = @($hiddenMedia | Where-Object { -not $visibleMedia.Contains($_) })
    $shared = @($hiddenMedia | Where-Object { $visibleMedia.Contains($_) })
    $bytes = 0
    foreach ($media in $onlyHidden) { $entry = $zip.GetEntry($media); if ($null -ne $entry) { $bytes += $entry.Length } }
    [pscustomobject]@{
        HiddenSheetsWithMedia = @($sheetMedia.Keys | Where-Object { $sheetState[$_] -ne 'visible' -and $sheetMedia[$_].Count -gt 0 }).Count
        HiddenUniqueMediaFiles = $hiddenMedia.Count
        VisibleUniqueMediaFiles = $visibleMedia.Count
        MediaSharedByHiddenAndVisible = $shared.Count
        MediaReferencedOnlyByHidden = $onlyHidden.Count
        BytesReferencedOnlyByHidden = $bytes
        MBReferencedOnlyByHidden = [math]::Round($bytes / 1MB, 2)
    } | Format-List
} finally { $zip.Dispose() }
