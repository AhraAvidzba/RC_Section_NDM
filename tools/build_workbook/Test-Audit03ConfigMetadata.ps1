# Проверяет служебные Config-поля на отдельной копии собранной книги.
# Формульные ID должны следовать строкам сочетаний и именованному диапазону;
# продолжения объединений не являются вторым вводом. Макросы не запускаются,
# исходная книга не сохраняется и не меняется. Это не проверка нормативных
# коэффициентов и не замена поведенческих тестов активных настроек.
param(
    [Parameter(Mandatory=$true)][string]$WorkbookPath,
    [Parameter(Mandatory=$true)][string]$RegistryPath,
    [Parameter(Mandatory=$true)][string]$ReportPath
)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$allowed = [IO.Path]::GetFullPath((Join-Path $root 'docs/regression/Audit03')) + [IO.Path]::DirectorySeparatorChar
$source = (Resolve-Path -LiteralPath (Join-Path $root $WorkbookPath)).Path
$report = [IO.Path]::GetFullPath((Join-Path $root $ReportPath))
foreach ($path in @($source, $report)) {
    if (-not $path.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase)) { throw 'Audit paths must stay in docs/regression/Audit03.' }
}
if (Test-Path -LiteralPath $report) { throw 'Evidence already exists; choose a new report path.' }
$registry = Get-Content -LiteralPath (Join-Path $root $RegistryPath) -Raw -Encoding UTF8 | ConvertFrom-Json
$sourceHash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
$fixtureDirectory = Join-Path ([IO.Path]::GetTempPath()) ('RC_NDM_Metadata_' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $fixtureDirectory | Out-Null
$fixture = Join-Path $fixtureDirectory 'RC_Section_NDM.xlsm'
Copy-Item -LiteralPath $source -Destination $fixture
$lines = [Collections.Generic.List[string]]::new()
$script:passed = 0
$script:failed = 0
$excel = $null
$book = $null

# Сохраняет точный test-ID каждого фактического адреса; не превращает ошибку
# проверки в успешную строку и не останавливает остальные независимые поля.
function Assert-Metadata([string]$Id, [bool]$Condition, [string]$Detail = '') {
    if ($Condition) { $script:passed++; $lines.Add('OK: ' + $Id + '; ' + $Detail) }
    else { $script:failed++; $lines.Add('FAIL: ' + $Id + '; ' + $Detail) }
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($fixture, 0, $false)
    $sheet = $book.Worksheets.Item('Config')
    $loads = $book.Names.Item('rngLoadCombinations').RefersToRange
    $duration = $book.Names.Item('rngStabilityDurationLoads').RefersToRange
    $loadName = $book.Names.Item('rngLoadCombinations')
    $originalReference = $loadName.RefersTo
    Assert-Metadata 'metadata.macrosDisabled' ($excel.AutomationSecurity -eq 3)
    Assert-Metadata 'metadata.slotCount' ($loads.Rows.Count -eq 31 -and $duration.Rows.Count -eq 31)

    # Смена ID и пустой ID проверяют вычисляемую зависимость всех 30 строк.
    # Ожидания не выводятся из текущей формулы: это явный контракт связи по slot.
    for ($slot = 1; $slot -le 30; $slot++) {
        $input = $loads.Cells.Item($slot + 1, 1)
        $derived = $duration.Cells.Item($slot + 1, 1)
        $prefix = 'metadata.duration.Slot' + $slot
        Assert-Metadata ($prefix + '.formula') ($derived.HasFormula -and ([string]$derived.Formula).Contains('rngLoadCombinations')) ([string]$derived.Address($false, $false))
        $input.Value2 = 'AUDIT03_META_' + $slot
        $sheet.Calculate()
        Assert-Metadata ($prefix + '.id') ([string]$derived.Value2 -ceq ('AUDIT03_META_' + $slot))
        [void]$input.ClearContents()
        $sheet.Calculate()
        Assert-Metadata ($prefix + '.blank') ([string]$derived.Value2 -ceq '')
    }

    # Переносим только controlled named range. INDEX должен читать новую
    # область, а не прежний жестко зашитый адрес исходной таблицы сочетаний.
    $relocated = $sheet.Range('CH1200:CN1230')
    [void]$loads.Copy($relocated)
    $loadName.RefersTo = '=' + $relocated.Address($true, $true, 1, $true)
    for ($slot = 1; $slot -le 30; $slot++) {
        $relocated.Cells.Item($slot + 1, 1).Value2 = 'AUDIT03_MOVED_' + $slot
    }
    $sheet.Calculate()
    for ($slot = 1; $slot -le 30; $slot++) {
        Assert-Metadata ('metadata.duration.Slot' + $slot + '.relocated') ([string]$duration.Cells.Item($slot + 1, 1).Value2 -ceq ('AUDIT03_MOVED_' + $slot))
    }
    $loadName.RefersTo = $originalReference

    # Поля-продолжения проверяются по фактическому MergeArea. Они не должны
    # иметь самостоятельное значение, формулу или второй несвязанный selector.
    $followers = @($registry.Fields | Where-Object Role -EQ 'NotEditable:MergedFollower')
    Assert-Metadata 'metadata.followers.count' ($followers.Count -eq 16)
    foreach ($field in $followers) {
        $cell = $sheet.Range(([string]$field.Address -split '!')[-1])
        $prefix = 'metadata.follower.' + ([string]$cell.Address($false, $false))
        Assert-Metadata ($prefix + '.merged') ([bool]$cell.MergeCells)
        Assert-Metadata ($prefix + '.notAnchor') ($cell.Address() -ne $cell.MergeArea.Cells.Item(1, 1).Address())
        Assert-Metadata ($prefix + '.noValue') ($null -eq $cell.Value2 -and -not $cell.HasFormula)
        Assert-Metadata ($prefix + '.singleDropdown') ($cell.MergeArea.Cells.Item(1, 1).Validation.Type -eq 3)
    }

    # Неприменимые клетки сохраняют явную подпись, а не ложный нулевой ввод.
    # Их consumer-активность остается предметом существующих профильных тестов.
    $notApplicable = @($registry.Fields | Where-Object Role -EQ 'NotEditable:NotApplicableCell')
    Assert-Metadata 'metadata.notApplicable.count' ($notApplicable.Count -eq 10)
    foreach ($field in $notApplicable) {
        $cell = $sheet.Range(([string]$field.Address -split '!')[-1])
        $prefix = 'metadata.notApplicable.' + ([string]$cell.Address($false, $false))
        Assert-Metadata ($prefix + '.marker') ([string]$cell.Value2 -ceq '-')
        Assert-Metadata ($prefix + '.notFormula') (-not $cell.HasFormula)
    }
    $lines.Add('METADATA_SCOPE: durationDerived=30; mergedFollowers=16; notApplicableLabels=10; normativeTable=NotTested; stateSolve=NotInvoked')
    $book.Close($false)
    $book = $null
} catch {
    $script:failed++
    $lines.Add('FAIL: metadata.runtime; ' + $_.Exception.Message + '; ' + $_.ScriptStackTrace)
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) {
        $excel.Quit()
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null
    }
}
$unchanged = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -eq $sourceHash
Assert-Metadata 'metadata.sourceUnchanged' $unchanged $sourceHash
$lines.Add("TOTAL_AUDIT03_CONFIG_METADATA: passed=$script:passed; failed=$script:failed")
$lines.Add('SOURCE_SHA256: ' + $sourceHash)
$lines.Add('SOURCE_UNCHANGED: ' + $unchanged)
[IO.File]::WriteAllLines($report, $lines, (New-Object Text.UTF8Encoding($true)))
Write-Output "CONFIG_METADATA_COMPLETED: passed=$script:passed; failed=$script:failed; report=$report"
if ($script:failed -gt 0) { exit 1 }
