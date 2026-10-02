# Проверяет правила выравнивания Config в реальной изолированной Audit03-книге.
# Ожидания берутся из реестра input-полей и границ отдельных таблиц, независимо
# от production helper. По явному флагу применяет только штатное форматирование,
# проверяя сохранность всех зарегистрированных данных и повторное открытие.
param(
    [Parameter(Mandatory=$true)][string]$WorkbookPath,
    [Parameter(Mandatory=$true)][string]$ReportPath,
    [string]$RegistryPath = 'docs/regression/Audit03/config_field_registry_v71_2026-10-02.csv',
    [switch]$ApplyAlignments
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$bookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
$allowed = [IO.Path]::GetFullPath((Join-Path $root 'docs/regression/Audit03')) + [IO.Path]::DirectorySeparatorChar
if (-not $bookPath.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Проверка/исправление допускается только в изолированной Audit03-книге.'
}
$fields = @(Import-Csv -LiteralPath (Join-Path $root $RegistryPath))
if ($fields.Count -ne 1065) { throw "Неожиданный реестр Config: $($fields.Count) адресов вместо 1065." }
$lines = New-Object 'System.Collections.Generic.List[string]'
$script:failed = 0
$script:passed = 0
$excel = $null
$book = $null
$beforeHash = (Get-FileHash -LiteralPath $bookPath -Algorithm SHA256).Hash

# Фиксирует каждое проверяемое свойство, не выдавая количество ячеек за
# количество физических сценариев и не скрывая исходный negative-run.
function Assert-Format([string]$Name, [bool]$Passed, [string]$Detail) {
    if ($Passed) { $script:passed++ } else { $script:failed++ }
    $lines.Add("FORMAT|$Name|passed=$Passed|$Detail")
}

# Хеширует снимок данных с фиксированным порядком полей и записей.
# Выравнивание исключается только у адресов с явно проверенным ожиданием.
function Get-RecordSignature([object]$Records) {
    $payload = ConvertTo-Json -InputObject @($Records) -Depth 9 -Compress
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($payload))).Replace('-', '') }
    finally { $sha.Dispose() }
}

# Сохраняет значения и формулы всего Config блочным чтением, включая поля,
# отсутствующие в реестре, и независимо фиксирует input-validation/форматы.
function Get-DataSignature([object]$Book) {
    $used = $Book.Worksheets.Item('Config').UsedRange
    return Get-RecordSignature ([ordered]@{ address=$used.Address(); values=$used.Value2; formulas=$used.Formula })
}

# Строит независимое ожидание INPUT и шапок/колонок реальных подтаблиц.
# Размеры задаются в относительных координатах именованного диапазона;
# комментарий последней подтаблицы не распространяется на соседние поля.
function Get-ExpectedAlignment([object]$Book) {
    $expected = @{}
    foreach ($field in $fields) {
        if ($field.Role -in @('UserInput','NormativeTableValue') -and $field.Id -notmatch '\.Comment$') {
            $expected[$field.Address] = @{ horizontal=-4108; vertical=-4108; role='Input' }
        }
    }
    # Opening не имеет вводимого n: четыре прочерка являются служебным
    # отображением автоматического количества, а не новыми INPUT-полями.
    $hollow = $Book.Names.Item('rngHollowRectangleGeometry').RefersToRange
    foreach ($row in @(15,16,17,18)) {
        $cell = $hollow.Cells.Item($row,4)
        if ([string]$cell.Value2 -ne '-') { throw "Неожиданный ввод количества Opening: $($cell.Address())." }
        $expected[('Config!'+$cell.Address($false,$false))] = @{ horizontal=-4108; vertical=-4108; role='AutomaticCountMarker' }
    }
    $sections = @{
        rngRoundedRectangleGeometry = @(@(2,3,5),@(6,10,5),@(12,16,6),@(18,22,9))
        rngHollowRectangleGeometry = @(@(2,4,4),@(6,7,7),@(10,18,6),@(20,28,9))
        rngRectSetGeometry = @(@(2,4,4),@(7,8,5),@(10,18,8),@(20,28,9))
    }
    $names = @('rngSystemSettings','rngUnitSettings','rngSignConventionSettings',
        'rngCircleGeometry','rngRectSetGeometry','rngRoundedRectangleGeometry','rngHollowRectangleGeometry',
        'rngConcreteMaterialParameters','rngSteelMaterialParameters','rngCalculationProfiles',
        'rngPlotAnnotationSettings','rngLoadCombinations','rngStabilityDurationLoads','rngSP35Table721')
    foreach ($name in $names) {
        $range = $Book.Names.Item($name).RefersToRange
        $blocks = @(@(1,$range.Rows.Count,$range.Columns.Count))
        if ($sections.ContainsKey($name)) { $blocks = $sections[$name] }
        foreach ($block in $blocks) {
            for ($column=1; $column -le $block[2]; $column++) {
                $headerRow = $block[0]
                $header = [string]$range.Cells.Item($headerRow,$column).Value2
                # Profiles содержит служебную первую строку и собственную шапку ниже.
                if ($header -notin @('Комментарий','Комментарии','Comment','Ед.','Справка','INTERNAL','Значение') -and
                    -not $sections.ContainsKey($name) -and $block[1] -ge 2) {
                    $headerRow++
                    $header = [string]$range.Cells.Item($headerRow,$column).Value2
                }
                if ($header -in @('Ед.','Справка','INTERNAL','Значение')) {
                    $lastRow = $headerRow
                    if ($header -ne 'Значение') { $lastRow=$block[1] }
                    for ($row=$headerRow; $row -le $lastRow; $row++) {
                        $cell = $range.Cells.Item($row,$column)
                        $address = 'Config!'+$cell.Address($false,$false)
                        $expected[$address] = @{ horizontal=-4108; vertical=-4108; role="Service:$header" }
                    }
                } elseif ($header -in @('Комментарий','Комментарии','Comment')) {
                    for ($row=$headerRow; $row -le $block[1]; $row++) {
                        $cell = $range.Cells.Item($row,$column)
                        $address = 'Config!'+$cell.Address($false,$false)
                        $horizontal = -4131
                        if ($row -gt $headerRow -and $column -eq $block[2]) { $horizontal=-4152 }
                        $expected[$address] = @{ horizontal=$horizontal; vertical=$null; role='Comment' }
                    }
                }
            }
        }
    }
    return $expected
}

# Проверяет фактические COM-свойства каждой ожидаемой ячейки. До исправления
# сохраняет подробный список отклонений; в Verify-режиме каждое является FAIL.
function Test-Alignment([object]$Book, [hashtable]$Expected, [string]$Phase, [bool]$Assert) {
    $config = $Book.Worksheets.Item('Config')
    $different = 0
    foreach ($address in ($Expected.Keys | Sort-Object)) {
        $cell = $config.Range($address.Split('!')[1])
        $wanted = $Expected[$address]
        $horizontal = [int]$cell.HorizontalAlignment
        $vertical = [int]$cell.VerticalAlignment
        $ok = $horizontal -eq $wanted.horizontal -and ($null -eq $wanted.vertical -or $vertical -eq $wanted.vertical)
        if (-not $ok) {
            $different++
            $lines.Add("ALIGNMENT_DEVIATION|phase=$Phase|address=$address|role=$($wanted.role)|horizontal=$horizontal->$($wanted.horizontal)|vertical=$vertical->$($wanted.vertical)")
        }
        if ($Assert) { Assert-Format "$Phase.$address" $ok "role=$($wanted.role); horizontal=$horizontal; vertical=$vertical" }
    }
    $lines.Add("ALIGNMENT_SCAN|phase=$Phase|addresses=$($Expected.Count)|deviations=$different")
}

# Сравнивает все 1065 зарегистрированных адресов без исключения данных общих
# selector anchors. Разрешены только ожидаемые H/V изменения; validation,
# значения, формулы, числовые форматы, заливки и шрифты сохраняются строго.
function Get-PreservationSignature([object]$Book, [hashtable]$Expected) {
    $records = New-Object 'System.Collections.Generic.List[object]'
    foreach ($field in $fields) {
        $parts = $field.Address.Split('!')
        $cell = $Book.Worksheets.Item($parts[0]).Range($parts[1])
        $validation = 'NoValidation'
        try {
            $validation = [ordered]@{ type=[int]$cell.Validation.Type; formula1=[string]$cell.Validation.Formula1;
                formula2=[string]$cell.Validation.Formula2; alert=[int]$cell.Validation.AlertStyle;
                ignoreBlank=[bool]$cell.Validation.IgnoreBlank; inCellDropdown=[bool]$cell.Validation.InCellDropdown }
        } catch [System.Runtime.InteropServices.COMException] {}
        $horizontal = [int]$cell.HorizontalAlignment
        $vertical = [int]$cell.VerticalAlignment
        if ($Expected.ContainsKey($field.Address)) {
            $horizontal = $Expected[$field.Address].horizontal
            if ($null -ne $Expected[$field.Address].vertical) { $vertical=$Expected[$field.Address].vertical }
        }
        $records.Add([ordered]@{ address=$field.Address; value=$cell.Value2; formula=$cell.Formula;
            numberFormat=[string]$cell.NumberFormat; merge=[string]$cell.MergeArea.Address();
            horizontal=$horizontal; vertical=$vertical; validation=$validation;
            fill=[int]$cell.Interior.Color; font=[string]$cell.Font.Name; size=[double]$cell.Font.Size;
            bold=[bool]$cell.Font.Bold; fontColor=[int]$cell.Font.Color })
    }
    $script:lastPreservationRecords = $records.ToArray()
    return Get-RecordSignature $script:lastPreservationRecords
}

# При несовпадении снимка показывает конкретные данные, а не только SHA.
# Это позволяет отличить разрешенное оформление от изменения исходных
# значений или validation; необъясненное различие остается failed gate.
function Assert-Preservation([object]$Book, [hashtable]$Expected, [string]$Phase,
    [string]$BeforeSignature, [object[]]$BeforeRecords) {
    $actual = Get-PreservationSignature $Book $Expected
    if ($BeforeSignature -ne $actual) {
        for ($i=0; $i -lt $BeforeRecords.Count; $i++) {
            $beforeJson = ConvertTo-Json -InputObject $BeforeRecords[$i] -Depth 9 -Compress
            $afterJson = ConvertTo-Json -InputObject $script:lastPreservationRecords[$i] -Depth 9 -Compress
            if ($beforeJson -ne $afterJson) {
                $lines.Add("PRESERVATION_DIFFERENCE|phase=$Phase|before=$beforeJson|after=$afterJson")
            }
        }
    }
    Assert-Format $Phase ($BeforeSignature -eq $actual) "fields=1065; before=$BeforeSignature; actual=$actual"
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($bookPath)
    $expected = Get-ExpectedAlignment $book
    $lines.Add("SOURCE|book=$bookPath|SHA256=$beforeHash|apply=$ApplyAlignments|registryFields=$($fields.Count)")
    $dataBefore = Get-DataSignature $book
    $preservedBefore = Get-PreservationSignature $book $expected
    $beforeRecords = @($script:lastPreservationRecords)
    Test-Alignment $book $expected 'before' (-not $ApplyAlignments)
    if ($ApplyAlignments) {
        Apply-ConfigUserInputAlignment $book
        Apply-ConfigCommentColumnAlignment $book
        Apply-ConfigShortColumnAlignment $book
        Test-Alignment $book $expected 'after' $true
        Assert-Format 'dataPreserved' ($dataBefore -eq (Get-DataSignature $book)) "signature=$dataBefore"
        Assert-Preservation $book $expected 'fieldsPreserved' $preservedBefore $beforeRecords
        Apply-ConfigUserInputAlignment $book
        Apply-ConfigCommentColumnAlignment $book
        Apply-ConfigShortColumnAlignment $book
        Assert-Format 'idempotentData' ($dataBefore -eq (Get-DataSignature $book)) "signature=$dataBefore"
        Assert-Preservation $book $expected 'idempotentFields' $preservedBefore $beforeRecords
        Test-Alignment $book $expected 'secondApply' $true
        $book.Save()
        $book.Close($false)
        $book = $excel.Workbooks.Open($bookPath)
        Test-Alignment $book $expected 'reopened' $true
        Assert-Format 'dataSaveReopen' ($dataBefore -eq (Get-DataSignature $book)) "signature=$dataBefore"
        Assert-Preservation $book $expected 'fieldsSaveReopen' $preservedBefore $beforeRecords
    }
} catch {
    $script:failed++
    $lines.Add("FORMATTING_RUNTIME_FAILURE: $($_.Exception.Message); $($_.ScriptStackTrace)")
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) {
        $excel.Quit()
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null
    }
    $afterHash = (Get-FileHash -LiteralPath $bookPath -Algorithm SHA256).Hash
    if (-not $ApplyAlignments) { Assert-Format 'sourceUnchanged' ($beforeHash -eq $afterHash) "before=$beforeHash; after=$afterHash" }
    $lines.Add("TOTAL_AUDIT03_CONFIG_FORMATTING: passed=$script:passed; failed=$script:failed; bookSHA256=$afterHash")
    $lines | Set-Content -LiteralPath (Join-Path $root $ReportPath) -Encoding UTF8
}
if ($script:failed -gt 0) { exit 1 }
