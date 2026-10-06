# Печатает только измененные блоки справки из сохраненной книги.
# Свойства печати меняются в отдельном read-only экземпляре и не сохраняются.
param(
    [string]$WorkbookPath = 'docs/regression/ContourExportHelp_2026-10-06/Publication/RC_Section_NDM.xlsm',
    [string]$SettingKey = 'SLS.Crack.SP35.NeighborRatioLimit',
    [string]$ReportDirectory = ''
)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path
$path = Join-Path $root $WorkbookPath
$directory = Join-Path $PSScriptRoot 'Visual'
if ($ReportDirectory) { $directory = Join-Path $root $ReportDirectory; New-Item -ItemType Directory -Path $directory -Force | Out-Null }
$hash = (Get-FileHash -LiteralPath $path).Hash
$excel = $null; $book = $null

# Сохраняет PDF с настоящим форматированием Excel, не изменяя ширину столбцов.
function Export-HelpRange([object]$Sheet, [int]$Start, [int]$End, [string]$Name) {
    $Sheet.PageSetup.PrintArea = "A${Start}:H${End}"
    $Sheet.PageSetup.Orientation = 2
    $Sheet.PageSetup.PaperSize = 8
    $Sheet.PageSetup.Zoom = $false
    $Sheet.PageSetup.FitToPagesWide = 1
    $Sheet.PageSetup.FitToPagesTall = $false
    $Sheet.ExportAsFixedFormat(0, (Join-Path $directory "$Name.pdf"), 0, $true, $false)
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false; $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($path, 0, $true)
    $guide = $book.Worksheets.Item('Справка')
    $guide.Outline.ShowLevels(8) | Out-Null
    $data = $guide.UsedRange.Value2
    $settings = $book.Names.Item('rngSystemSettings').RefersToRange
    $start = 0
    for ($row = 1; $row -le $settings.Rows.Count; $row++) {
        if ([string]$settings.Cells.Item($row,1).Value2 -ne $SettingKey) { continue }
        $link = [string]$settings.Cells.Item($row,5).Hyperlinks.Item(1).SubAddress
        if ($link -notmatch '!\$?A\$?(\d+)$') { throw "Unexpected Help link: $link" }
        $start = [int]$Matches[1]; break
    }
    if ($start -eq 0) { throw "$SettingKey Help was not found." }
    $end = $start
    while ([string]$data[($end + 1),1] -match '^\d+$') { $end++ }
    Export-HelpRange $guide $start $end ('Help_' + ($SettingKey -split '\.')[-1])
    $start = 0
    for ($row = 1; $row -le $data.GetLength(0); $row++) {
        if ([string]$data[$row,1] -eq 'Справка по результатам расчета листа Results') { $start = $row; break }
    }
    if ($start -eq 0) { throw 'Results Help was not found.' }
    $end = $data.GetLength(0)
    for ($row = $start + 1; $row -le $data.GetLength(0); $row++) {
        if ([int]$guide.Rows.Item($row).OutlineLevel -eq 1 -and [string]$data[$row,1] -ne '') { $end = $row - 1; break }
    }
    Export-HelpRange $guide $start $end 'Help_Results'
} finally {
    if ($null -ne $book) { $book.Close($false); [Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null }
    if ($null -ne $excel) { $excel.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
if ((Get-FileHash -LiteralPath $path).Hash -ne $hash) { throw 'Read-only render changed the workbook.' }
Write-Output 'CHANGED_HELP_READ_ONLY_RENDER_OK'
