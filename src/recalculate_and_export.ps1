$ErrorActionPreference = 'Stop'
$taskRoot = Split-Path -Parent $PSScriptRoot
$bookPath = Join-Path $taskRoot 'deliverables\LITE_Bull_Case_DCF.xlsx'
$deckPath = Join-Path $taskRoot 'deliverables\LITE_DCF_Slides.pptx'
$pdfPath = Join-Path $taskRoot 'deliverables\LITE_DCF_Slides.pdf'
$renderPath = Join-Path $taskRoot '.render\lite'
New-Item -ItemType Directory -Force -Path $renderPath | Out-Null
$excelApp = $null
$book = $null
try {
    $excelApp = New-Object -ComObject Excel.Application
    $excelApp.Visible = $false
    $excelApp.DisplayAlerts = $false
    $book = $excelApp.Workbooks.Open($bookPath, 0, $false)
    $excelApp.CalculateFullRebuild()
    $book.Save()
    Write-Output ('Excel recalculated. Bull Gordon: ' + $book.Worksheets.Item('Bull').Range('B66').Value2)
    $book.Close($false)
    $book = $null
} finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $excelApp) { $excelApp.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excelApp) }
}
$powerpointApp = $null
$deck = $null
try {
    $powerpointApp = New-Object -ComObject PowerPoint.Application
    $deck = $powerpointApp.Presentations.Open($deckPath, $true, $false, $false)
    $deck.SaveAs($pdfPath, 32)
    $deck.Export($renderPath, 'PNG', 1600, 900)
    Write-Output ('PowerPoint exported: ' + $deck.Slides.Count + ' slides')
    $deck.Close()
    $deck = $null
} finally {
    if ($null -ne $deck) { $deck.Close() }
    if ($null -ne $powerpointApp) { $powerpointApp.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($powerpointApp) }
}
