# Сравнивает штатный расчет сохраненных пользовательских сочетаний при No/Yes.
# Работает только с двумя временными копиями, не заменяет геометрию, нагрузки,
# материалы или профили тестовыми значениями и не меняет исходную книгу.
param(
    [Parameter(Mandatory=$true)][string]$WorkbookPath,
    [Parameter(Mandatory=$true)][string]$ReportPath
)
$ErrorActionPreference = 'Stop'
$source = (Resolve-Path -LiteralPath $WorkbookPath).Path
$sourceHash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
$folder = Join-Path ([IO.Path]::GetTempPath()) ('RC_NDM_SavedPair_' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $folder | Out-Null
$report = [IO.Path]::GetFullPath($ReportPath)
$lines = New-Object System.Collections.Generic.List[string]
$excel = $null
$book = $null
$failed = $false
$results = @{}

try {
    $lines.Add("SOURCE: $source; SHA256=$sourceHash")
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    foreach ($mode in @('No', 'Yes')) {
        $fixture = Join-Path $folder ("RC_Section_NDM_$mode.xlsm")
        Copy-Item -LiteralPath $source -Destination $fixture
        $book = $excel.Workbooks.Open($fixture)
        $module = $book.VBProject.VBComponents.Item('modTestWorkbookInterface').CodeModule
        $module.AddFromString(@'
' ДЛЯ ТЕСТОВ: запускает обычный макрос с исходными Config и сочетаниями.
' Отключены лишь автоматическая схема и txt-report, не расчетные проверки.
Public Function Audit02RunSavedCombinationPair(ByVal mode As String) As String
    SetSystemSetting "General.DiagramExtension", mode
    SetSystemSetting "Plot.AutoUpdateAfterCalculation", "No"
    SetSystemSetting "General.ExecutionReportEnabled", "No"
    Audit02RunSavedCombinationPair = RunSectionCalculationForWorkbook(ThisWorkbook, False)
End Function
'@
        )
        $lines.Add("STARTED: $mode; $([DateTime]::Now.ToString('s')); $fixture")
        $lines | Set-Content -LiteralPath $report -Encoding UTF8
        $watch = [Diagnostics.Stopwatch]::StartNew()
        $message = [string]$excel.Run("'$($book.Name)'!modTestWorkbookInterface.Audit02RunSavedCombinationPair", $mode)
        $watch.Stop()
        $lines.Add("MESSAGE: $mode; $message")
        if ($message -notmatch 'Расчет завершен') { throw $message }
        $loadRows = $book.Names.Item('rngLoadCombinations').RefersToRange.Rows.Count
        $summary = $book.Names.Item('rngBatchSummary').RefersToRange
        $table = $summary.Resize($loadRows + 12, 26).Value2
        $anchor = $book.Names.Item('rngStrengthSummaryAnchor').RefersToRange
        $detail = $anchor.Resize($loadRows, 49).Value2
        $rows = @{}
        for ($r = 13; $r -le $table.GetLength(0); $r++) {
            $id = [string]$table[$r, 1]
            if (-not $id) { continue }
            $status = [string]$table[$r, 7]
            $detailStatus = $null
            for ($d = 1; $d -le $detail.GetLength(0); $d++) {
                if ([string]$detail[$d, 1] -eq $id) { $detailStatus = [string]$detail[$d, 49]; break }
            }
            if ($detailStatus -ne $status) { throw "LC $id ($mode): summary=$status, detail=$detailStatus." }
            $rows[$id] = [pscustomobject]@{ Status=$status; Overall=[string]$table[$r,4]; Comment=[string]$table[$r,3] }
            $lines.Add("LC|$mode|$id|Capacity=$status|Overall=$($table[$r,4])|SummaryDetailMatch=True|Comment=$($table[$r,3])")
        }
        $results[$mode] = $rows
        $lines.Add("FINISHED: $mode; elapsedSec=$($watch.Elapsed.TotalSeconds); rows=$($rows.Count)")
        $book.Close($false)
        [Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null
        $book = $null
    }
    $different = 0
    $baseFailPairs = 0
    foreach ($id in $results.No.Keys | Sort-Object) {
        if (-not $results.Yes.ContainsKey($id)) { throw "LC $id отсутствует в On." }
        $off = $results.No[$id].Status
        $on = $results.Yes[$id].Status
        if ($off -ne $on) { $different++ }
        if ($off -eq 'NumFail' -and $on -eq 'BaseFail') { $baseFailPairs++ }
        $lines.Add("PAIR|$id|Off=$off|On=$on")
    }
    $lines.Add("PAIRS: total=$($results.No.Count); different=$different; OffNumFailOnBaseFail=$baseFailPairs")
    if ((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -ne $sourceHash) { throw 'Исходная книга изменена.' }
    $lines.Add('SOURCE_UNCHANGED: True')
} catch {
    $failed = $true
    $lines.Add("SCRIPT ERROR: $($_.Exception.Message)")
} finally {
    if ($book) { try { $book.Close($false) } catch {}; [Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null }
    if ($excel) { try { $excel.Quit() } catch {}; [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    $lines | Set-Content -LiteralPath $report -Encoding UTF8
    $lines
}
if ($failed) { exit 1 }
