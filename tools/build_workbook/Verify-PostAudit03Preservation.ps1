# Сверяет выпуск с заранее сохраненным пользовательским снимком.
# Читает книгу read-only: Config, верхние результаты и ширины не исправляются.
param([Parameter(Mandatory=$true)][string]$WorkbookPath,
      [Parameter(Mandatory=$true)][string]$ReportPath)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$path=(Resolve-Path -LiteralPath (Join-Path $root $WorkbookPath)).Path
$report=[IO.Path]::GetFullPath((Join-Path $root $ReportPath))
$allowed=[IO.Path]::GetFullPath((Join-Path $root 'docs/regression/PostAudit03'))+[IO.Path]::DirectorySeparatorChar
if(-not $path.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase) -or
   -not $report.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)) {throw 'Preservation evidence must remain inside PostAudit03.'}
$baseline=@(Import-Clixml -LiteralPath (Join-Path $root 'docs/regression/PostAudit03/Baseline/UserSheets.clixml'))
$hash=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
$excel=$null; $book=$null; $differences=@(); $widthDifferences=@(); $compared=0; $widths=0
# Две пустые ячейки считаются одинаковыми; числа сравниваются как точные Double.
function Test-ExactCell([object]$Expected,[object]$Actual) {
    if($null -eq $Expected) {$Expected=''}
    if($null -eq $Actual) {$Actual=''}
    if($Expected -is [ValueType] -and $Actual -is [ValueType]) {return [double]$Expected -eq [double]$Actual}
    return [string]$Expected -ceq [string]$Actual
}
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false; $excel.AutomationSecurity=3
    $book=$excel.Workbooks.Open($path,0,$true)
    if(-not $book.ReadOnly) {throw 'Preservation check must be read-only.'}
    $results=$book.Worksheets.Item('Results')
    $crack=$results.Range('rngCrackSummaryAnchor')
    $geometry=$results.Range('rngNDMSectionGeometry')
    foreach($saved in $baseline) {
        $sheet=$book.Worksheets.Item($saved.Sheet)
        $range=$sheet.Range($saved.Address)
        $actual=New-Object 'System.Collections.Generic.List[object]'
        foreach($value in $range.Formula) {$actual.Add($value)}
        if($actual.Count -ne $saved.Formula.Count) {throw "Flattened snapshot size differs: $($saved.Sheet)"}
        $columns=$range.Columns.Count
        for($index=0;$index -lt $actual.Count;$index++) {
            $row=$range.Row+[int][Math]::Floor($index/$columns)
            $column=$range.Column+($index%$columns)
            if($saved.Sheet -eq 'Results') {
                if($row -ge $geometry.Row-1) {continue}
                if($row -eq $crack.Row-3 -and $column -ge $crack.Column+62 -and $column -le $crack.Column+66) {continue}
            }
            $compared++
            if(-not (Test-ExactCell $saved.Formula[$index] $actual[$index])) {
                $differences+=@{Sheet=$saved.Sheet;Row=$row;Column=$column;Expected=$saved.Formula[$index];Actual=$actual[$index]}
            }
        }
        foreach($width in $saved.Widths) {
            $widths++
            $current=[double]$sheet.Columns.Item([int]$width.Column).ColumnWidth
            if($current -ne [double]$width.Width) {
                $widthDifferences+=@{Sheet=$saved.Sheet;Column=$width.Column;Expected=$width.Width;Actual=$current}
            }
        }
    }
} finally {
    if($book) {try {$book.Close($false)} catch {}; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}
    if($excel) {try {$excel.Quit()} catch {}; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
$after=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
$reportHash=(Get-FileHash -LiteralPath (Join-Path $root 'workbook/output/RC_Section_NDM_execution_report.txt')).Hash
$reportPreserved=$reportHash -eq '05303BCAFDF04378813F3EE6F30109A0258D608DF92C0E280E77E593F304AC79'
$passed=$differences.Count -eq 0 -and $widthDifferences.Count -eq 0 -and $hash -eq $after -and $reportPreserved
[ordered]@{Passed=$passed;ComparedCells=$compared;ComparedWidths=$widths;CellDifferences=$differences;
    WidthDifferences=$widthDifferences;SourceUnchanged=($hash -eq $after);SourceSHA256=$after;
    UserExecutionReportPreserved=$reportPreserved;UserExecutionReportSHA256=$reportHash;
    Exclusions='Results lower snapshot band and five approved crack-width heading cells only'} |
    ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $report -Encoding UTF8
Write-Output "PRESERVATION: passed=$passed; cells=$compared; widths=$widths; cellDifferences=$($differences.Count); widthDifferences=$($widthDifferences.Count); userReport=$reportPreserved"
if(-not $passed) {exit 1}
