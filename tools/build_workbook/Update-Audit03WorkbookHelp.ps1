# Обновляет справку только в изолированной Audit03-книге. Проверяет фактический
# лист, ссылки и неизменность input-значений/формул/validation. По явному флагу
# объединяет повторные общие селекторы RectSet с отчетом прежних значений;
# не подтверждает
# нормативную трассировку либо пиксельную визуальную приемку.
param(
    [Parameter(Mandatory=$true)][string]$WorkbookPath,
    [Parameter(Mandatory=$true)][string]$ReportPath,
    [string]$RegistryPath = 'docs/regression/Audit03/config_field_registry_2026-10-02.csv',
    [switch]$UpdateRectSetSharedSelectors
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$bookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
$allowed = [IO.Path]::GetFullPath((Join-Path $root 'docs/regression/Audit03')) + [IO.Path]::DirectorySeparatorChar
if (-not $bookPath.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Обновление справки допускается только в изолированной Audit03-книге.'
}
$fields = @(Import-Csv -LiteralPath (Join-Path $root $RegistryPath) | Where-Object Role -eq 'UserInput')
if ($fields.Count -eq 0) { throw 'Реестр не содержит пользовательских полей.' }
$lines = New-Object 'System.Collections.Generic.List[string]'
$script:failed = 0
$excel = $null
$book = $null

# Сравнивает все зарегистрированные input-ячейки, включая неактивные поля.
# Отсутствие validation - допустимое состояние, остальные свойства читаются
# без скрытой подстановки. Ссылки справки не являются input-значениями.
function Get-InputSignature([object]$Book) {
    $records = New-Object 'System.Collections.Generic.List[object]'
    foreach ($field in $fields) {
        $parts = $field.Address.Split('!')
        $cell = $Book.Worksheets.Item($parts[0]).Range($parts[1])
        $validation = $null
        try {
            $type = [int]$cell.Validation.Type
            $validation = @{type=$type; formula1=[string]$cell.Validation.Formula1; formula2=[string]$cell.Validation.Formula2;
                alert=[int]$cell.Validation.AlertStyle; ignoreBlank=[bool]$cell.Validation.IgnoreBlank}
        } catch [System.Runtime.InteropServices.COMException] {
            $validation = 'NoValidation'
        }
        $records.Add(@{id=$field.Id; address=$field.Address; value=$cell.Value2; formula=$cell.Formula;
            numberFormat=[string]$cell.NumberFormat; horizontal=[int]$cell.HorizontalAlignment;
            vertical=[int]$cell.VerticalAlignment; validation=$validation})
    }
    $script:lastInputRecords = $records.ToArray()
    return Get-InputRecordsSignature $script:lastInputRecords
}

# Хеширует зафиксированные input-records. Отдельный вызов нужен только для
# явно согласованного центрирования общих RectSet-селекторов; их значения,
# формулы, validation и number format остаются в неизменном строгом сравнении.
function Get-InputRecordsSignature([object[]]$Records) {
    $payload = ConvertTo-Json -InputObject $Records -Depth 8 -Compress
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($payload))).Replace('-', '') }
    finally { $sha.Dispose() }
}

# Сохраняет отдельное доказательство каждого утверждения; программная ошибка
# или отсутствие ожидаемой строки не превращается в общий успешный gate.
function Assert-Help([string]$Name, [bool]$Passed, [string]$Detail) {
    if (-not $Passed) { $script:failed++ }
    $lines.Add("HELP|$Name|passed=$Passed|$Detail")
}

# Проверяет реальные объединения и единственный dropdown каждой общей
# настройки. Нижняя ячейка больше не является отдельным вводом; все остальные
# input-поля остаются в строгой проверке сохранности значений и validation.
function Test-RectSetSharedLayout([object]$Book) {
    $target = $Book.Names.Item('rngRectSetGeometry').RefersToRange
    foreach ($firstRow in @(21, 23, 25, 27)) {
        foreach ($column in @(3, 4, 6, 7)) {
            $first = $target.Cells.Item($firstRow, $column)
            $second = $target.Cells.Item(($firstRow + 1), $column)
            $expected = $first.Resize(2, 1).Address()
            Assert-Help ("rectset.merge."+$first.Address()) ($first.MergeCells -and $first.MergeArea.Address() -eq $expected) $expected
            Assert-Help ("rectset.dropdown."+$first.Address()) ($first.Validation.Type -eq 3 -and $first.Validation.InCellDropdown) ([string]$first.Validation.Formula1)
            Assert-Help ("rectset.followerEmpty."+$second.Address()) ($null -eq $second.Value2) 'Shared selector has no second value'
            Assert-Help ("rectset.centered."+$first.Address()) ($first.HorizontalAlignment -eq -4108 -and $first.VerticalAlignment -eq -4108) 'Shared input is centered'
        }
    }
}

# Читает сохраненный лист и каждую прямую Config-ссылку. Проверка содержания
# целевая: словарь статусов и измененные input/extension/stability контракты.
# Остальные методические разделы требуют самостоятельной содержательной ревизии.
function Test-ActualHelp([object]$Book) {
    $guide = $Book.Worksheets.Item('Справка')
    $data = $guide.UsedRange.Value2
    $start = 0
    $text = New-Object Text.StringBuilder
    for ($r = 1; $r -le $data.GetLength(0); $r++) {
        if ([string]$data[$r,1] -eq 'Словарь пользовательских статусов') { $start = $r }
        [void]$text.AppendLine(([string]$data[$r,2]))
    }
    Assert-Help 'statusDictionaryPresent' ($start -gt 0) "row=$start"
    $statuses = @('OK','FAIL','BaseFail','NumFail','InputErr','CalcErr','N/A')
    if ($start -gt 0) {
        for ($i = 0; $i -lt $statuses.Count; $i++) {
            $description = [string]$data[($start+$i+1),2]
            Assert-Help ("status."+$statuses[$i]) ([string]$data[($start+$i+1),1] -eq $statuses[$i] -and $description -match '[А-Яа-я]') $description
        }
        Assert-Help 'baseFailIsPhysical' ([string]$data[($start+3),2] -match 'Физический отказ подтвержден' -and [string]$data[($start+3),2] -match 'не просто несходимость') 'lambda=0 requires physical evidence'
        Assert-Help 'notApplicableDependency' ([string]$data[($start+7),2] -match 'предыдущего результата') 'Blocked dependency is not an OK check'
    }
    $body = $text.ToString()
    Assert-Help 'extensionAllStates' ($body.Contains('PreCrackState/PostCrackState/CurrentCrackedState')) 'Global extension scope'
    Assert-Help 'extensionDoesNotChangePhysicalNodes' ($body.Contains('Исходные физические точки, сопротивления, касательные, плато и предельные деформации не меняются')) 'Physical diagram preserved'
    Assert-Help 'profileBooleanInputContract' ($body.Contains('Все четыре переключателя Calculation.* обязательны')) 'Explicit invalid value is not No'
    Assert-Help 'subdivisionsInputContract' ($body.Contains('Mesh.BoundarySubdivisions')) 'Mesh setting has help'
    Assert-Help 'worstCriterionInputContract' ($body.Contains('неизвестное значение этого критерия') -and $body.Contains('InputErr')) 'Invalid criterion cannot silently select another check'
    Assert-Help 'rectsetCommonSelectors' ($body.Contains('общие для двух сторон каждой грани H1, B1, H2 или B2')) 'One common selector, separate diameters'
    Assert-Help 'unitChoiceInputContract' ($body.Contains('Пустой выбор, TODO или неподдержанная единица являются ошибкой Config')) 'Explicit invalid units do not become defaults'
    Assert-Help 'signsDoNotFlipStressOrStrain' ($body.Contains('Пользовательские знаки N/Mx/My не меняют знак Stress и Strain')) 'Material tension/compression and strain plane keep internal signs'
    $config = $Book.Worksheets.Item('Config')
    $count = 0
    foreach ($link in $config.Hyperlinks) {
        $target = [string]$link.SubAddress
        if ($target -match "^'?Справка'?!([A-Z]+[0-9]+)$") {
            $count++
            $cell = $guide.Range($Matches[1])
            Assert-Help ("link."+$link.Range.Address()) (-not [string]::IsNullOrWhiteSpace([string]$cell.Value2)) $target
        } elseif ($target.Contains('Справка')) {
            Assert-Help ("linkSyntax."+$link.Range.Address()) $false $target
        }
    }
    Assert-Help 'directHelpLinksPresent' ($count -gt 100) "count=$count"
    $lines.Add("ACTUAL_HELP: rows=$($guide.UsedRange.Rows.Count); links=$count; shapes=$($guide.Shapes.Count)")
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($bookPath)
    if ($UpdateRectSetSharedSelectors) {
        $fullBefore = Get-InputSignature $book
        $target = $book.Names.Item('rngRectSetGeometry').RefersToRange
        $followers = @{}
        $selectorAnchors = @{}
        foreach ($firstRow in @(21, 23, 25, 27)) {
            foreach ($column in @(3, 4, 6, 7)) {
                $address = $target.Cells.Item(($firstRow + 1), $column).Address($false, $false)
                $followers[([string]$target.Worksheet.Name+'!'+$address)] = $true
                $anchorAddress = $target.Cells.Item($firstRow, $column).Address($false, $false)
                $selectorAnchors[([string]$target.Worksheet.Name+'!'+$anchorAddress)] = $true
            }
        }
        $originalCount = $fields.Count
        $fields = @($fields | Where-Object { -not $followers.ContainsKey($_.Address) })
        $lines.Add("RECTSET_PRESERVATION_SCOPE: originalFields=$originalCount; unchangedFields=$($fields.Count); sharedFollowers=16; fullBefore=$fullBefore")
        $before = Get-InputSignature $book
        $beforeRecords = @($script:lastInputRecords)
        $lines.Add("RECTSET_UNCHANGED_FIELDS_ORIGINAL: signature=$before")
        foreach ($record in $beforeRecords) {
            if ($selectorAnchors.ContainsKey($record.address)) {
                if ($record.horizontal -ne -4108 -or $record.vertical -ne -4108) {
                    $lines.Add("RECTSET_APPROVED_ALIGNMENT: address=$($record.address); horizontal=$($record.horizontal)->-4108; vertical=$($record.vertical)->-4108; values/formulas/validation/numberFormat unchanged")
                }
                $record.horizontal = -4108
                $record.vertical = -4108
            }
        }
        $before = Get-InputRecordsSignature $beforeRecords
        $migration = @(Set-RectSetSharedSelectorLayout $target)
        foreach ($record in $migration) {
            $lines.Add('RECTSET_SHARED_SELECTOR: '+(ConvertTo-Json $record -Compress))
        }
        Assert-Help 'rectset.migrationRecords' ($migration.Count -eq 16) "count=$($migration.Count)"
        Test-RectSetSharedLayout $book
        $layoutOnce = Get-InputSignature $book
        Set-RectSetSharedSelectorLayout $target | Out-Null
        Assert-Help 'rectset.layoutIdempotent' ($layoutOnce -eq (Get-InputSignature $book)) "signature=$layoutOnce"
    } else {
        $before = Get-InputSignature $book
        $beforeRecords = @($script:lastInputRecords)
    }
    Add-SettingsInstructions $book $book.Worksheets.Item('Config') $book.Worksheets.Item('Справка')
    $after = Get-InputSignature $book
    $afterRecords = @($script:lastInputRecords)
    if ($before -ne $after) {
        for ($i = 0; $i -lt $beforeRecords.Count; $i++) {
            $expectedRecord = ConvertTo-Json -InputObject $beforeRecords[$i] -Depth 8 -Compress
            $actualRecord = ConvertTo-Json -InputObject $afterRecords[$i] -Depth 8 -Compress
            if ($expectedRecord -ne $actualRecord) {
                $lines.Add("INPUT_CHANGED: before=$expectedRecord; after=$actualRecord")
            }
        }
    }
    Assert-Help 'inputsPreserved' ($before -eq $after) "fields=$($fields.Count); before=$before; after=$after"
    Test-ActualHelp $book
    $book.Save()
    $book.Close($false)
    $book = $excel.Workbooks.Open($bookPath)
    $reopened = Get-InputSignature $book
    Assert-Help 'inputsSaveReopen' ($before -eq $reopened) "before=$before; reopened=$reopened"
    if ($UpdateRectSetSharedSelectors) { Test-RectSetSharedLayout $book }
    Test-ActualHelp $book
    $lines.Add("TOTAL_AUDIT03_HELP: failed=$script:failed")
} catch {
    $lines.Add("HELP_RUNTIME_FAILURE: $($_.Exception.Message); $($_.ScriptStackTrace)")
    throw
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) {
        $excel.Quit()
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null
    }
    $lines | Set-Content -LiteralPath (Join-Path $root $ReportPath) -Encoding UTF8
}
if ($script:failed -gt 0) { exit 1 }
