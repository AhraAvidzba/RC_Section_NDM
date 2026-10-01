# Сохраняет результат обычного расчета текущих Config/LC в итоговой книге.
# Сначала работает с копией и сверяет Config. Исходник заменяется только после
# успешного макроса/save и проверки, что он не изменился параллельно.
param(
    [string]$WorkbookPath = 'workbook/output/RC_Section_NDM.xlsm',
    [string]$ReportPath = 'docs/regression/Audit02/final_saved_calculation_2026-10-01.txt'
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$source = (Resolve-Path -LiteralPath $WorkbookPath).Path
$hash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
$printAreas = @(Get-WorkbookPrintAreas $source)
$folder = Join-Path ([IO.Path]::GetTempPath()) ('RC_NDM_FinalSnapshot_' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $folder | Out-Null
$fixture = Join-Path $folder ([IO.Path]::GetFileName($source))
Copy-Item -LiteralPath $source -Destination $fixture
$lines = New-Object System.Collections.Generic.List[string]
$excel = $null
$book = $null
$failed = $false

# Сериализует все значения и формулы Config; сравнение не зависит от display.
function Get-ConfigSnapshot([object]$workbook) {
    $range = $workbook.Worksheets.Item('Config').UsedRange
    $values = $range.Value2
    $formulas = $range.Formula
    $snapshot = New-Object Text.StringBuilder
    [void]$snapshot.AppendLine($range.Address())
    for ($r=1; $r -le $values.GetLength(0); $r++) {
        for ($c=1; $c -le $values.GetLength(1); $c++) {
            foreach ($value in @($values[$r,$c],$formulas[$r,$c])) {
                $text = [string]$value
                [void]$snapshot.Append($text.Length).Append(':').Append($text).Append('|')
            }
        }
    }
    return $snapshot.ToString()
}

try {
    $lines.Add("SOURCE: $source; BEFORE_SHA256=$hash")
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($fixture)
    $before = Get-ConfigSnapshot $book
    $message = [string]$excel.Run("'$($book.Name)'!modWorkbookCalculation.RunSectionCalculationForWorkbook", $book, $false)
    $lines.Add("CALCULATION_MESSAGE: $message")
    if ($message -notmatch 'Расчет завершен') { throw $message }
    if ($before -cne (Get-ConfigSnapshot $book)) { throw 'Обычный расчет изменил Config.' }
    $lines.Add('CONFIG_VALUES_AND_FORMULAS_UNCHANGED: True')
    $book.Save()
    $book.Close($false)
    [Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null
    $book = $null
    Restore-WorkbookPrintAreas $fixture $printAreas
    if ((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -ne $hash) { throw 'Исходная книга изменилась параллельно.' }
    Copy-Item -LiteralPath $fixture -Destination $source -Force
    $executionReport = Join-Path $folder 'RC_Section_NDM_execution_report.txt'
    if (Test-Path -LiteralPath $executionReport) {
        Copy-Item -LiteralPath $executionReport -Destination (Join-Path (Split-Path $source -Parent) 'RC_Section_NDM_execution_report.txt') -Force
        $lines.Add('EXECUTION_REPORT_SAVED: True')
    }
    $lines.Add("AFTER_SHA256: $((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash)")
    $lines.Add('FINAL_SNAPSHOT_SAVED: True')
} catch {
    $failed = $true
    $lines.Add("SCRIPT ERROR: $($_.Exception.Message)")
} finally {
    if ($book) { try { $book.Close($false) } catch {}; [Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null }
    if ($excel) { try { $excel.Quit() } catch {}; [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    $lines | Set-Content -LiteralPath (Join-Path $root $ReportPath) -Encoding UTF8
    $lines
}
if ($failed) { exit 1 }
