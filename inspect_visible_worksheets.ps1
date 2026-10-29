$ErrorActionPreference = 'Stop'

$workbookPath = 'D:\Huawei Handover\PAT DOC_LP89M_2026 Capacity Expansion NW Modernization_EXP_BKK.xlsx'
$reportDir = 'D:\Huawei Handover\reports'
New-Item -ItemType Directory -Path $reportDir -Force | Out-Null

function Get-CellText([object]$ws, [int]$row, [int]$col) {
    if ($row -lt 1 -or $col -lt 1) { return '' }
    try {
        $value = $ws.Cells.Item($row, $col).Text
        if ($null -eq $value) { return '' }
        return ([string]$value).Trim()
    } catch { return '' }
}

function Get-NearbyText([object]$ws, [int]$row, [int]$col) {
    $parts = New-Object System.Collections.Generic.List[string]
    $startRow = [Math]::Max(1, $row - 2)
    $startCol = [Math]::Max(1, $col - 2)
    $endRow = $row + 2
    $endCol = $col + 2
    try {
        $values = $ws.Range($ws.Cells.Item($startRow, $startCol), $ws.Cells.Item($endRow, $endCol)).Value2
        for ($r = 1; $r -le 5; $r++) {
            for ($c = 1; $c -le 5; $c++) {
                $value = $values[$r, $c]
                if ($null -ne $value) {
                    $text = ([string]$value).Trim()
                    if ($text -and -not $parts.Contains($text)) { [void]$parts.Add($text) }
                }
            }
        }
    } catch {
        $text = Get-CellText $ws $row $col
        if ($text) { [void]$parts.Add($text) }
    }
    return ($parts -join ' | ')
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $null
$summary = @()
$map = @()

try {
    $wb = $excel.Workbooks.Open($workbookPath, $null, $true)
    $visibleSheets = @($wb.Worksheets | Where-Object { $_.Visible -eq -1 })
    $hiddenSheets = @($wb.Worksheets | Where-Object { $_.Visible -ne -1 })
    $visiblePictureCount = 0
    $visibleShapeCount = 0

    foreach ($ws in $visibleSheets) {
        $ur = $ws.UsedRange
        $shapeCount = [int]$ws.Shapes.Count
        $pictureCount = 0
        for ($i = 1; $i -le $shapeCount; $i++) {
            $shape = $ws.Shapes.Item($i)
            if ($shape.Type -eq 13 -or $shape.Type -eq 11) { $pictureCount++ }
        }
        $visibleShapeCount += $shapeCount
        $visiblePictureCount += $pictureCount

        $orientation = 'Portrait'
        if ([int]$ws.PageSetup.Orientation -eq 2) { $orientation = 'Landscape' }
        $summary += [pscustomobject]@{
            WorksheetIndex = [int]$ws.Index
            Worksheet = [string]$ws.Name
            UsedRange = [string]$ur.Address($false, $false)
            UsedRows = [int]$ur.Rows.Count
            UsedColumns = [int]$ur.Columns.Count
            ShapeCount = $shapeCount
            PictureCount = $pictureCount
            PrintArea = [string]$ws.PageSetup.PrintArea
            Orientation = $orientation
            FitToPagesWide = [int]$ws.PageSetup.FitToPagesWide
            FitToPagesTall = [int]$ws.PageSetup.FitToPagesTall
            HorizontalPageBreaks = [int]$ws.HPageBreaks.Count
            VerticalPageBreaks = [int]$ws.VPageBreaks.Count
        }

        for ($i = 1; $i -le $shapeCount; $i++) {
            $shape = $ws.Shapes.Item($i)
            $topLeft = $shape.TopLeftCell
            $bottomRight = $shape.BottomRightCell
            $shapeType = [int]$shape.Type
            $isPicture = ($shapeType -eq 13 -or $shapeType -eq 11)
            if (-not $isPicture) { continue }
            $altText = ''
            try { $altText = [string]$shape.AlternativeText } catch {}
            $map += [pscustomobject]@{
                WorksheetIndex = [int]$ws.Index
                Worksheet = [string]$ws.Name
                ShapeIndex = $i
                ShapeName = [string]$shape.Name
                ShapeType = $shapeType
                IsPicture = $isPicture
                TopLeftCell = [string]$topLeft.Address($false, $false)
                BottomRightCell = [string]$bottomRight.Address($false, $false)
                Left = [math]::Round([double]$shape.Left, 1)
                Top = [math]::Round([double]$shape.Top, 1)
                Width = [math]::Round([double]$shape.Width, 1)
                Height = [math]::Round([double]$shape.Height, 1)
                AlternativeText = $altText
                NearbyCellText = Get-NearbyText $ws ([int]$topLeft.Row) ([int]$topLeft.Column)
            }
        }
    }

    $summary | Export-Csv -LiteralPath (Join-Path $reportDir 'visible_worksheet_summary.csv') -NoTypeInformation -Encoding UTF8
    $map | Export-Csv -LiteralPath (Join-Path $reportDir 'visible_worksheet_map.csv') -NoTypeInformation -Encoding UTF8

    $links = $wb.LinkSources(1)
    $linkCount = 0
    if ($null -ne $links) { $linkCount = @($links).Count }
    [pscustomobject]@{
        Workbook = $workbookPath
        TotalWorksheets = [int]$wb.Worksheets.Count
        VisibleWorksheets = $visibleSheets.Count
        HiddenWorksheets = $hiddenSheets.Count
        VisibleShapes = $visibleShapeCount
        VisiblePictures = $visiblePictureCount
        ExternalLinks = $linkCount
        SummaryCsv = (Join-Path $reportDir 'visible_worksheet_summary.csv')
        MapCsv = (Join-Path $reportDir 'visible_worksheet_map.csv')
    } | ConvertTo-Json
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

