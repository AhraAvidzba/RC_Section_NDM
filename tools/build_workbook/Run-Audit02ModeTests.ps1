# Выполняет все штатные VBA suites на независимой копии заданной книги.
# Меняет только глобальную настройку копии: исходная baseline/output книга
# остается неизменной. Explicit-On setup внутри исторических тестов сохранен.
param(
    [Parameter(Mandatory=$true)][string]$SourceWorkbook,
    [string]$SettingKey = "General.DiagramExtension",
    [ValidateSet("Yes", "No")][string]$Mode = "No",
    [Parameter(Mandatory=$true)][string]$ReportPath,
    [string[]]$Macro = @(),
    [string]$MacroArgument1 = "",
    [string]$MacroArgument2 = "",
    [switch]$VerifyResultsReopen,
    [switch]$VerifyStatusReopen,
    [switch]$Visible
)
$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "SettingsCatalog.ps1")
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "../..")).Path
$sourcePath = (Resolve-Path -LiteralPath $SourceWorkbook).Path
$sourceHash = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash
$fixtureRoot = Join-Path ([IO.Path]::GetTempPath()) ("RC_NDM_ModeSuite_" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $fixtureRoot | Out-Null
$fixturePath = Join-Path $fixtureRoot "RC_Section_NDM.xlsm"
Copy-Item -LiteralPath $sourcePath -Destination $fixturePath
$printAreas = @(Get-WorkbookPrintAreas $fixturePath)
$lines = New-Object System.Collections.Generic.List[string]
$fullReport = Join-Path $root $ReportPath
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $fullReport) | Out-Null
$excel = $null
$workbook = $null
$failed = $false

# Определяет PID только созданного COM-экземпляра Excel по его окну.
# Watchdog использует этот идентификатор, не затрагивая пользовательские книги.
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class AuditExcelProcessIdentity {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
}
'@

# Сохраняет ход прогона до COM-вызова: при зависании видно последнюю suite,
# а уже завершенные проверки не теряются вместе с тестовым процессом Excel.
function Save-Progress {
    $lines | Set-Content -LiteralPath $fullReport -Encoding UTF8
}

# Диагностирует потерю COM-свойств после suite без повторного открытия книги
# и без признания прогона успешным. Прямой IDispatch-вызов позволяет отделить
# недоступность Excel от проблемы адаптера PowerShell при том же RCW-объекте.
function Save-ComPropertyDiagnostic {
    param([object]$Target, [string]$Label, [string[]]$Properties)
    if ($null -eq $Target) { $lines.Add("COM_DIAGNOSTIC: $Label is null"); return }
    $type = $Target.GetType()
    $lines.Add("COM_DIAGNOSTIC: $Label; type=$($type.FullName); isCom=$([Runtime.InteropServices.Marshal]::IsComObject($Target))")
    foreach ($property in $Properties) {
        try {
            $value = $type.InvokeMember($property, [Reflection.BindingFlags]::GetProperty, $null, $Target, $null)
            if ($null -eq $value) { $detail = 'null' }
            elseif ([Runtime.InteropServices.Marshal]::IsComObject($value)) {
                $detail = "COM object; type=$($value.GetType().FullName)"
                try {
                    $count = $value.GetType().InvokeMember('Count', [Reflection.BindingFlags]::GetProperty, $null, $value, $null)
                    $detail += "; rawCount=$count"
                } catch { $detail += "; countError=$($_.Exception.Message)" }
            } else { $detail = [string]$value }
            $lines.Add("COM_DIAGNOSTIC: $Label.$property; raw=$detail")
        } catch { $lines.Add("COM_DIAGNOSTIC: $Label.$property; rawError=$($_.Exception.Message)") }
    }
}

# Читает свойство того же COM-объекта напрямую через IDispatch. После VBA Run
# адаптер PowerShell может вернуть null для существующей коллекции, тогда как
# прямой getter видит открытую книгу. Здесь нет reopen, retry или подстановки
# данных; настоящий null остается ошибкой, COM-коллекция не разворачивается
# в pipeline и сохраняет собственные Count/Item.
function Get-RequiredComProperty {
    param([object]$Target, [string]$Property)
    if ($null -eq $Target) { throw "COM-объект недоступен при чтении $Property." }
    $value = $Target.GetType().InvokeMember($Property, [Reflection.BindingFlags]::GetProperty, $null, $Target, $null)
    if ($null -eq $value) {
        Save-ComPropertyDiagnostic $Target 'UnavailableTarget' @($Property)
        throw "COM-свойство $Property недоступно при прямом чтении."
    }
    return ,$value
}

# Находит единственную строку настройки через именованный Range листа Config.
# Не зависит от промежуточного COM-proxy Workbook.Names после долгого VBA-вызова;
# неизвестный формат не заменяет настройку другим default или режимом.
function Get-ModeSettingCell {
    param([object]$Book, [string]$Key)
    if ($null -eq $Book) { throw "Книга недоступна при чтении тестового режима $Key." }
    $worksheets = Get-RequiredComProperty $Book 'Worksheets'
    $config = $worksheets.Item('Config')
    if ($null -eq $config) { throw "Лист Config недоступен при чтении тестового режима $Key." }
    $range = $config.Range('rngSystemSettings')
    if ($null -eq $range) { throw "rngSystemSettings не ссылается на ячейки при чтении тестового режима $Key." }
    $data = $range.Value2
    $row = 0
    for ($r = 2; $r -le $range.Rows.Count; $r++) {
        if ([string]$data[$r, 1] -eq $Key) {
            if ($row) { throw "В тестовой книге несколько строк $Key." }
            $row = $r
        }
    }
    if (-not $row) { throw "В тестовой книге отсутствует $Key." }
    return $range.Cells.Item($row, 2)
}

# Сравнивает только сохраненные значения Results, включая все комментарии.
# Это gate сохранности данных, а не проверка цветов или внешнего вида.
function Get-ResultsValueHash([object]$Book) {
    $worksheets = Get-RequiredComProperty $Book 'Worksheets'
    $range = $worksheets.Item("Results").UsedRange
    $data = $range.Value2
    $payload = ConvertTo-Json -InputObject @{
        rows = $range.Rows.Count; columns = $range.Columns.Count; values = @($data)
    } -Depth 8 -Compress
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($payload))).Replace('-', '') }
    finally { $sha.Dispose() }
}

# Сохраняет фактическое оформление всех статусных колонок, включая пустые
# строки, легенду и направленную palette-fixture. DisplayFormat учитывает CF;
# это проверка Excel COM, а не утверждение о просмотренных пикселях экрана.
function Get-ResultsStatusStyleHash([object]$Book) {
    $names = Get-RequiredComProperty $Book 'Names'
    $rows = [int]$names.Item('rngLoadCombinations').RefersToRange.Rows.Count - 1
    $records = New-Object 'System.Collections.Generic.List[object]'
    $blocks = @(
        @{name='rngBatchSummary'; offset=12; columns=@(4,6,7,8,9,10,11,12,13,14,15)},
        @{name='rngStrengthSummaryAnchor'; offset=0; columns=@(3,30,49)},
        @{name='rngCrackSummaryAnchor'; offset=0; columns=@(3,19,20,21,23,45,49)},
        @{name='rngStabilitySummaryAnchor'; offset=0; columns=@(3,38,44,53,59,72,84)}
    )
    foreach ($block in $blocks) {
        $anchor = $names.Item($block.name).RefersToRange
        for ($r = 0; $r -lt $rows; $r++) {
            foreach ($column in $block.columns) {
                $cell = $anchor.Worksheet.Cells.Item($anchor.Row + $block.offset + $r, $anchor.Column + $column - 1)
                $records.Add(@{block=$block.name; address=$cell.Address(); value=[string]$cell.Value2;
                    fill=[int]$cell.Interior.Color; pattern=[int]$cell.Interior.Pattern;
                    display=[int]$cell.DisplayFormat.Interior.Color; displayPattern=[int]$cell.DisplayFormat.Interior.Pattern;
                    conditionalRules=[int]$cell.FormatConditions.Count})
            }
        }
    }
    $anchor = $names.Item('rngBatchSummary').RefersToRange
    for ($r = 7; $r -le 13; $r++) {
        $cell = $anchor.Worksheet.Cells.Item($anchor.Row + $r, $anchor.Column + 27)
        $records.Add(@{block='legend'; address=$cell.Address(); value=[string]$cell.Value2;
            fill=[int]$cell.Interior.Color; display=[int]$cell.DisplayFormat.Interior.Color;
            conditionalRules=[int]$cell.FormatConditions.Count})
    }
    $worksheets = Get-RequiredComProperty $Book 'Worksheets'
    foreach ($sheet in $worksheets) {
        if ($sheet.Name -ne '__Audit03Palette') { continue }
        for ($r = 1; $r -le 12; $r++) {
            $cell = $sheet.Cells.Item($r, 1)
            $records.Add(@{block='palette'; address=$cell.Address(); value=[string]$cell.Value2;
                fill=[int]$cell.Interior.Color; pattern=[int]$cell.Interior.Pattern;
                display=[int]$cell.DisplayFormat.Interior.Color; displayPattern=[int]$cell.DisplayFormat.Interior.Pattern;
                conditionalRules=[int]$cell.FormatConditions.Count})
        }
    }
    $payload = ConvertTo-Json -InputObject @($records.ToArray()) -Depth 6 -Compress
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($payload))).Replace('-', '') }
    finally { $sha.Dispose() }
}

try {
    $lines.Add("SOURCE: $sourcePath")
    $lines.Add("SOURCE_SHA256: $sourceHash")
    $lines.Add("FIXTURE: $fixturePath")
    $lines.Add("GLOBAL_MODE: $SettingKey=$Mode; explicit-On tests retain their setup")
    Save-Progress
    $excel = New-Object -ComObject Excel.Application
    [uint32]$testExcelProcessId = 0
    [void][AuditExcelProcessIdentity]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$testExcelProcessId)
    if ($testExcelProcessId -eq 0) { throw 'Не удалось определить собственный тестовый процесс Excel.' }
    $testExcelProcess = Get-Process -Id $testExcelProcessId
    $lines.Add("TEST_EXCEL_PROCESS: id=$testExcelProcessId; startTicks=$($testExcelProcess.StartTime.ToUniversalTime().Ticks)")
    Save-Progress
    $excel.Visible = [bool]$Visible
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $openBooks = Get-RequiredComProperty $excel 'Workbooks'
    $workbook = $openBooks.Open($fixturePath)
    (Get-ModeSettingCell $workbook $SettingKey).Value2 = $Mode
    $workbook.Save()
    $workbook.Close($false)
    $workbook = $null
    Restore-WorkbookPrintAreas $fixturePath $printAreas
    $macros = @(
        "modTestGeometry.RunGeometryTests",
        "modTestMaterialDiagrams.RunMaterialDiagramTests",
        "modTestSectionSolver.RunSectionSolverTests",
        "modTestCapacitySolver.RunCapacitySolverTests",
        "modTestCrackWidth.RunCrackWidthTests",
        "modTestBatchCalculation.RunBatchCalculationTests",
        "modTestWorkbookInterface.RunWorkbookInterfaceTests",
        "modTestRegressionBaseline.RunRegressionBaselineTests"
    )
    if ($Macro.Count -gt 0) { $macros = $Macro }
    if (($MacroArgument1 -or $MacroArgument2 -or $VerifyResultsReopen -or $VerifyStatusReopen) -and $macros.Count -ne 1) {
        throw "Аргументы macro/save-reopen допускаются только для одной явно выбранной проверки."
    }
    foreach ($macro in $macros) {
        $openBooks = Get-RequiredComProperty $excel 'Workbooks'
        $workbook = $openBooks.Open($fixturePath, $null, $true)
        $lines.Add("===== $macro =====")
        $lines.Add("SUITE_STARTED: $macro; $([DateTime]::Now.ToString('s'))")
        Save-Progress
        if ([string](Get-ModeSettingCell $workbook $SettingKey).Value2 -ne $Mode) {
            throw "Перед suite $macro не сохранено требуемое значение $Mode."
        }
        $macroName = "'RC_Section_NDM.xlsm'!$macro"
        # Фоновый тестовый Excel не должен прерывать VBA из-за Escape в другом
        # приложении. Excel сбрасывает этот режим после idle, поэтому задаем его
        # перед каждым Run; зависание по-прежнему ограничивает внешний watchdog.
        $excel.EnableCancelKey = 0
        $lines.Add('TEST_MACRO_CANCEL_KEY: Disabled; isolated Excel instance only')
        if ($MacroArgument2) { $result = [string]$excel.Run($macroName, $MacroArgument1, $MacroArgument2) }
        elseif ($MacroArgument1) { $result = [string]$excel.Run($macroName, $MacroArgument1) }
        else { $result = [string]$excel.Run($macroName) }
        foreach ($process in @(Get-Process EXCEL -ErrorAction SilentlyContinue)) {
            $lines.Add("EXCEL_AFTER_SUITE: pid=$($process.Id); workingSet=$($process.WorkingSet64); privateBytes=$($process.PrivateMemorySize64); cpu=$($process.CPU)")
        }
        foreach ($line in ($result -split "`r?`n")) {
            $lines.Add($line)
            if ($line -match "failed=([1-9][0-9]*)|^FAIL:|RUNTIME ERROR") { $failed = $true }
        }
        # Получаем заново ссылку на ту же открытую книгу после долгого VBA Run.
        # Файл не открывается повторно: несохраненные изменения suite остаются
        # доступными для post-mode и save/reopen проверок, без скрытого reset.
        $openBooks = Get-RequiredComProperty $excel 'Workbooks'
        $liveBook = $null
        for ($bookIndex = 1; $bookIndex -le $openBooks.Count; $bookIndex++) {
            $candidate = $openBooks.Item($bookIndex)
            if ([string]::Equals([string]$candidate.FullName, $fixturePath, [StringComparison]::OrdinalIgnoreCase)) {
                if ($null -ne $liveBook) { throw 'Найдено несколько открытых книг с путем тестовой копии.' }
                $liveBook = $candidate
            }
        }
        if ($null -eq $liveBook) { throw 'Тестовая книга закрыта или ее путь изменен во время suite.' }
        if (-not [bool]$liveBook.ReadOnly) { throw 'Тестовая книга перестала быть ReadOnly во время suite.' }
        $lines.Add("COM_WORKBOOK_AFTER_SUITE: path=$($liveBook.FullName); readOnly=$($liveBook.ReadOnly); sameProxy=$([object]::ReferenceEquals($workbook, $liveBook))")
        $workbook = $liveBook
        $modeAfter = [string](Get-ModeSettingCell $workbook $SettingKey).Value2
        $lines.Add("MODE_AFTER_SUITE: $modeAfter")
        $lines.Add("SUITE_FINISHED: $macro; $([DateTime]::Now.ToString('s'))")
        Save-Progress
        if ($modeAfter -ne $Mode) {
            $lines.Add("SUITE_MODE_OVERRIDE: explicit test setup left $modeAfter; changes discarded on close")
        }
        if ($VerifyResultsReopen -or $VerifyStatusReopen) {
            $beforeHash = Get-ResultsValueHash $workbook
            if ($VerifyStatusReopen) { $beforeStyleHash = Get-ResultsStatusStyleHash $workbook }
            $savedPath = Join-Path $fixtureRoot 'RC_Section_NDM_saved.xlsm'
            $workbook.SaveAs($savedPath, 52)
            $workbook.Close($false)
            $openBooks = Get-RequiredComProperty $excel 'Workbooks'
            $workbook = $openBooks.Open($savedPath, $null, $true)
            $afterHash = Get-ResultsValueHash $workbook
            $lines.Add("RESULTS_SAVE_REOPEN: before=$beforeHash; after=$afterHash; equal=$($beforeHash -eq $afterHash); file=$savedPath")
            if ($beforeHash -ne $afterHash) { throw "Значения или комментарии Results изменились после save/reopen." }
            if ($VerifyStatusReopen) {
                $afterStyleHash = Get-ResultsStatusStyleHash $workbook
                $lines.Add("STATUS_STYLE_SAVE_REOPEN: before=$beforeStyleHash; after=$afterStyleHash; equal=$($beforeStyleHash -eq $afterStyleHash)")
                if ($beforeStyleHash -ne $afterStyleHash) { throw "Отображаемая палитра/условное форматирование изменились после save/reopen." }
            }
        }
        $workbook.Close($false)
        $workbook = $null
    }
    if ((Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash -ne $sourceHash) {
        throw "Исходная книга изменилась во время проверки независимой копии."
    }
    $lines.Add("SOURCE_UNCHANGED: True")
} catch {
    $failed = $true
    $lines.Add("SCRIPT ERROR: $($_.Exception.Message)")
    $lines.Add("SCRIPT_ERROR_CONTEXT: macro=$macro; exceptionType=$($_.Exception.GetType().FullName)")
    $lines.Add("SCRIPT_ERROR_POSITION: $($_.InvocationInfo.PositionMessage)")
    $lines.Add("SCRIPT_ERROR_STACK: $($_.ScriptStackTrace)")
} finally {
    if ($workbook) {
        try { $workbook.Close($false) } catch { $lines.Add("CLEANUP ERROR: workbook; $($_.Exception.Message)") }
        [Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($excel) {
        try { $excel.Quit() } catch { $lines.Add("CLEANUP ERROR: Excel; $($_.Exception.Message)") }
        [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Save-Progress
    $lines | Where-Object { $_ -match "^SOURCE|^FIXTURE|^GLOBAL_MODE|^TOTAL|^FAIL:|^SCRIPT ERROR|^RUNTIME ERROR|^MODE_AFTER" }
}
if ($failed) { exit 1 }
