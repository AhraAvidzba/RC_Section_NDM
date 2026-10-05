# Снимает печатные представления реальной Excel-книги для визуальной приемки.
# Открывает только read-only копию; PrintArea/outline не сохраняются, исходные
# данные и ширины столбцов не меняются. PDF служат QA-артефактами, не отчетом.
param(
    [Parameter(Mandatory=$true)][string]$WorkbookPath,
    [Parameter(Mandatory=$true)][string]$ReportDirectory
)
$ErrorActionPreference = 'Stop'
$path = (Resolve-Path -LiteralPath $WorkbookPath).Path
$directory = [IO.Path]::GetFullPath($ReportDirectory)
New-Item -ItemType Directory -Path $directory -Force | Out-Null
$before = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
$excel = $null
$book = $null
$records = New-Object 'System.Collections.Generic.List[object]'

# Печатает указанную область с ее реальными merged cells, шрифтами и картинками.
# Одна ширина страницы не обрезает строки; высота остается многостраничной.
function Export-Range([object]$Sheet, [object]$Range, [string]$Name, [bool]$SinglePage) {
    $Sheet.PageSetup.PrintArea = $Range.Address()
    $Sheet.PageSetup.Orientation = 2
    $Sheet.PageSetup.PaperSize = 8
    $Sheet.PageSetup.Zoom = $false
    $Sheet.PageSetup.FitToPagesWide = 1
    $Sheet.PageSetup.FitToPagesTall = $false
    if ($SinglePage) { $Sheet.PageSetup.FitToPagesTall = 1 }
    $Sheet.PageSetup.LeftMargin = 15
    $Sheet.PageSetup.RightMargin = 15
    $Sheet.PageSetup.TopMargin = 15
    $Sheet.PageSetup.BottomMargin = 15
    $file = Join-Path $directory ($Name + '.pdf')
    $Sheet.ExportAsFixedFormat(0, $file, 0, $true, $false)
    $records.Add(@{name=$Name; range=$Range.Address(); file=$file})
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($path, 0, $true)
    $anchor = $book.Names.Item('rngCrackSummaryAnchor').RefersToRange
    $sheet = $anchor.Worksheet
    Export-Range $sheet $anchor.Offset(-4,0).Resize(10,29) 'Crack_Formation_Current' $true
    Export-Range $sheet $anchor.Offset(-4,30).Resize(10,14) 'Crack_SP35' $true
    Export-Range $sheet $anchor.Offset(-4,45).Resize(10,27) 'Crack_SP63_Longitudinal' $true
    $guide = $book.Worksheets.Item('Справка')
    $data = $guide.UsedRange.Value2
    $rows = @{}
    for ($row = 1; $row -le $data.GetLength(0); $row++) {
        $text = [string]$data[$row,1]
        foreach ($title in @('Расчет ширины раскрытия нормальных трещин', 'Подготовка геометрии в AutoCAD и импорт', 'Расчет устойчивости', 'Результаты расчета')) {
            if ($text -eq $title) { $rows[$title] = $row }
        }
        if ($text -match '^СП 35:') { $rows['SP35'] = $row }
    }
    $records.Add(@{name='HelpSections'; rows=$rows; shapes=$guide.Shapes.Count})
    $guide.Outline.ShowLevels(8)
    $crack = [int]$rows['Расчет ширины раскрытия нормальных трещин']
    $next = 0
    for ($row = $crack + 1; $row -le $data.GetLength(0); $row++) {
        if ([int]$guide.Rows.Item($row).OutlineLevel -eq 1 -and [string]$data[$row,1] -ne '') { $next = $row; break }
    }
    if ($crack -le 0 -or $next -le $crack) { throw 'Не найдена общая сворачиваемая группа справки трещин.' }
    Export-Range $guide $guide.Range("A${crack}:H$($next-1)") 'Help_Crack' $false
    $cad = [int]$rows['Подготовка геометрии в AutoCAD и импорт']
    $next = 0
    for ($row = $cad + 1; $row -le $data.GetLength(0); $row++) {
        if ([int]$guide.Rows.Item($row).OutlineLevel -eq 1 -and [string]$data[$row,1] -ne '') { $next = $row; break }
    }
    if ($cad -le 0 -or $next -le $cad) { throw 'Не найдена сворачиваемая группа подготовки AutoCAD.' }
    Export-Range $guide $guide.Range("A${cad}:H$($next-1)") 'Help_AutoCAD' $false
    $records | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $directory 'RenderManifest.json') -Encoding UTF8
}
finally {
    if ($null -ne $book) { $book.Close($false); [Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null }
    if ($null -ne $excel) { $excel.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
$after = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
if ($before -ne $after) { throw 'Read-only render changed the workbook.' }
Write-Output "READ_ONLY_RENDER_OK $directory"
