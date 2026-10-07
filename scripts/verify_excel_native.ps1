param([string]$ProjectRoot = (Split-Path -Parent $PSScriptRoot))
$ErrorActionPreference = 'Stop'
$reportPath = Join-Path $ProjectRoot 'outputs\example_retail\findings_report.xlsx'
$previewFolder = Join-Path $ProjectRoot 'tmp\qa'
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($reportPath, 0, $true)
    $excel.CalculateFullRebuild()
    $stats = $book.Worksheets.Item('Data_Summary')
    $clean = $book.Worksheets.Item('Cleaned_Data')
    $baseline = [double]$stats.Range('B3').Value2
    $qty = [double]$clean.Range('F2').Value2
    $price = [double]$clean.Range('G2').Value2
    $clean.Range('F2').Value2 = $qty + 1
    $excel.CalculateFullRebuild()
    $changed = [double]$stats.Range('B3').Value2
    if ([Math]::Abs($changed - $baseline - $price) -gt 0.000001) { throw 'Native Excel input-change test failed.' }
    $clean.Range('F2').Value2 = $qty
    $excel.CalculateFullRebuild()
    foreach ($sheetName in @('Charts','Cleaned_Data')) {
        $sheet = $book.Worksheets.Item($sheetName)
        $sheet.PageSetup.PrintArea = if ($sheetName -eq 'Charts') { '$A$1:$B$126' } else { '$A$1:$L$10' }
        $sheet.PageSetup.Zoom = $false
        $sheet.PageSetup.FitToPagesWide = 1
        $sheet.PageSetup.FitToPagesTall = if ($sheetName -eq 'Charts') { 5 } else { 1 }
        $sheet.ExportAsFixedFormat(0, (Join-Path $previewFolder ('native_' + $sheetName + '.pdf')))
    }
    Write-Output ('Native Excel recalculation passed: total=' + $baseline + ', changed=' + $changed + ', expected delta=' + $price)
    Write-Output ('Embedded chart images=' + $book.Worksheets.Item('Charts').Shapes.Count)
} finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $excel) { $excel.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
}
