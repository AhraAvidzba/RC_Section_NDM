# Обновляет только согласованные блоки Config в изолированной сохраненной
# книге: LoadPath, контрольные формулы, общие селекторы RectSet,
# проектные комментарии и актуальные списки выбора из проверенного каталога.
# Не удаляет строки листа и не подставляет defaults вместо пользовательских
# нагрузок. Все остальные значения/формулы проверяются до save и после reopen.
param(
    [Parameter(Mandatory=$true)][string]$WorkbookPath,
    [Parameter(Mandatory=$true)][string]$ReportPath,
    [string]$TemplateCensusPath = 'docs/regression/Audit03/config_census_clean_v322_2026-10-05.json',
    [string]$TemplateRegistryPath = 'docs/regression/Audit03/config_field_registry_clean_v322_2026-10-05.csv',
    [switch]$VerifyIdempotence
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$path = (Resolve-Path -LiteralPath $WorkbookPath).Path
$allowedRoot = [IO.Path]::GetFullPath((Join-Path $root 'docs/regression/Audit03')) + [IO.Path]::DirectorySeparatorChar
if (-not $path.StartsWith($allowedRoot, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Миграция разрешена только для изолированной Audit03-книги.'
}
$report = [IO.Path]::GetFullPath((Join-Path $root $ReportPath))
$printAreas = @(Get-WorkbookPrintAreas $path)
$lines = New-Object 'System.Collections.Generic.List[string]'
$exceptions = @{}
$excel = $null
$book = $null
$failed = $false
$template = Get-Content -LiteralPath (Join-Path $root $TemplateCensusPath) -Raw -Encoding UTF8 | ConvertFrom-Json
$fields = @(Import-Csv -LiteralPath (Join-Path $root $TemplateRegistryPath))
$templateCells = @{}
foreach ($cell in $template.Cells) { $templateCells[$cell.Address] = $cell }
$templateRanges = @{}
foreach ($range in $template.NamedRanges) { $templateRanges[$range.Name] = $range.Rectangle }
$registeredCells = @{}
foreach ($field in $fields) { $registeredCells[$field.Address.Split('!')[1]] = $true }
$validationPlans = New-Object 'System.Collections.Generic.List[object]'
$commentPlans = New-Object 'System.Collections.Generic.List[object]'

# Регистрирует только ячейки явно утвержденного преобразования. Остальная
# область Config сравнивается целиком, а не по неполному списку настроек.
function Add-MigrationRange([object]$Range) {
    for ($r = 1; $r -le $Range.Rows.Count; $r++) {
        for ($c = 1; $c -le $Range.Columns.Count; $c++) {
            $exceptions[$Range.Cells.Item($r, $c).Address($false, $false)] = $true
        }
    }
}

# Formula содержит и исходную формулу, и значение константной ячейки.
# Поэтому пересчет кэша формул не считается изменением пользовательского ввода.
function Get-PreservedConfigSignature([object]$Sheet, [string]$Bounds, [bool]$IncludeMigrationCells = $false) {
    $range = $Sheet.Range($Bounds)
    $data = $range.Formula
    $rowCount = $range.Rows.Count
    $columnCount = $range.Columns.Count
    $firstRow = $range.Row
    $firstColumn = $range.Column
    $records = New-Object 'System.Collections.Generic.List[object]'
    for ($r = 1; $r -le $rowCount; $r++) {
        for ($c = 1; $c -le $columnCount; $c++) {
            $address = ConvertTo-ExcelColumn ($firstColumn + $c - 1)
            $address += [string]($firstRow + $r - 1)
            if ($IncludeMigrationCells -or -not $exceptions.ContainsKey($address)) {
                $records.Add(@{address=$address; input=($data[$r, $c])})
            }
        }
    }
    $payload = ConvertTo-Json -InputObject $records.ToArray() -Depth 5 -Compress
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($payload))).Replace('-', '') }
    finally { $sha.Dispose() }
}

# Фиксирует массив сочетаний без заголовка. Validation и имя колонки меняются,
# но все ID, числа, формулы, выбранные пути и пользовательские комментарии нет.
function Get-LoadInputSignature([object]$Range) {
    return ConvertTo-Json -InputObject ($Range.Offset(1, 0).Resize($Range.Rows.Count - 1, $Range.Columns.Count).Formula) -Depth 4 -Compress
}

# Сопоставляет ячейку каталога с сохраненной таблицей по именованному блоку.
# Для SystemSettings строка определяется ключом: удаленная настройка не должна
# сдвигать соседний пользовательский ввод или адрес комментария.
function Get-MigrationCell([string]$Block, [string]$Address, [string]$Key = '') {
    if ($Address -notmatch '^([A-Z]+)([0-9]+)$') { throw "Не распознан адрес каталога $Address." }
    $column = 0
    foreach ($letter in $Matches[1].ToCharArray()) { $column = 26 * $column + [int]$letter - [int][char]'A' + 1 }
    $row = [int]$Matches[2]
    $source = $templateRanges[$Block]
    if ($null -eq $source) { throw "Нет блока $Block в проверенном каталоге." }
    $target = $book.Names.Item($Block).RefersToRange
    if ($Block -eq 'rngSystemSettings') {
        if (-not $Key) { $Key = [string]$templateCells[(ConvertTo-ExcelColumn $source.Left) + $row].Value }
        $found = 0
        for ($r = 2; $r -le $target.Rows.Count; $r++) {
            if ([string]$target.Cells.Item($r, 1).Value2 -ceq $Key) {
                if ($found -ne 0) { throw "Повторный ключ $Key в сохраненной книге." }
                $found = $r
            }
        }
        if ($found -eq 0) { throw "В сохраненной книге отсутствует обязательная настройка $Key." }
        return $target.Cells.Item($found, $column - $source.Left + 1)
    }
    return $config.Cells.Item($target.Row + $row - $source.Top, $target.Column + $column - $source.Left)
}

# Читает допустимые значения из census реально собранной книги. Короткие
# неизменяемые enum-списки переносятся как literal validation; динамические
# списки ID/профилей сохраняют ссылки на исходные пользовательские таблицы.
function Prepare-CatalogValidation {
    foreach ($field in @($fields | Where-Object Role -eq 'UserInput')) {
        $address = $field.Address.Split('!')[1]
        $record = $templateCells[$address]
        $rules = @($template.Validations | Where-Object { $record.Validations -contains $_.Range })
        if ($rules.Count -eq 0) { continue }
        if ($rules.Count -ne 1 -or $rules[0].Type -ne 'list') { throw "Неоднозначная validation для $($field.Id)." }
        $rule = $rules[0]
        $reference = [string]$rule.Formula1
        if ($reference -notmatch '^\$?([A-Z]+)\$?([0-9]+):\$?\1\$?([0-9]+)$') { throw "Не распознан источник validation $reference." }
        $letters = $Matches[1]
        $first = [int]$Matches[2]
        $last = [int]$Matches[3]
        $options = New-Object 'System.Collections.Generic.List[string]'
        $dynamic = $false
        for ($r = $first; $r -le $last; $r++) {
            $source = $templateCells[$letters + $r]
            if ($null -eq $source) { throw "Нет ячейки списка $letters$r в census." }
            if ($source.Formula) { $dynamic = $true }
            $options.Add([string]$source.Value)
        }
        $cell = Get-MigrationCell $field.Block $address $field.Id
        $formula = '=' + $reference
        if ($field.Block -eq 'rngLoadCombinations' -and $field.Id -like '*LoadPath*') {
            $formula = '=' + $list.Address($true, $true)
        } elseif ($dynamic) {
            $formula = [string]$cell.Validation.Formula1
            if (-not $formula.StartsWith('=')) { throw "Потеряна динамическая validation $($field.Id)." }
            $source = $config.Range($formula.Substring(1))
            if ($source.Rows.Count -ne $options.Count -or $source.Columns.Count -ne 1) { throw "Некорректная длина динамического списка $($field.Id)." }
        } else {
            $formula = $options.ToArray() -join ','
            if ($formula.Length -gt 255 -or @($options | Where-Object { $_.Contains(',') }).Count -gt 0) {
                throw "Enum-список $($field.Id) нельзя перенести без изменения его смысла."
            }
        }
        $validationPlans.Add([pscustomobject]@{ Id=$field.Id; Cell=$cell; Address=$cell.Address(); Formula=$formula; Dynamic=$dynamic;
            Options=$options.ToArray(); IgnoreBlank=($field.Block -eq 'rngLoadCombinations' -and $field.Id -like '*LoadPath*') })
    }
}

# Обновляет проектные комментарии, но не пользовательский столбец Comment
# сочетаний. Вложенные геометрические таблицы определяются по своим шапкам;
# зарегистрированные input/formula/reference-ячейки не изменяются.
function Prepare-CatalogComments {
    foreach ($name in $templateRanges.Keys) {
        if ($name -eq 'rngLoadCombinations') { continue }
        $bounds = $templateRanges[$name]
        $headers = @($template.Cells | Where-Object {
            $_.NamedRanges -contains $name -and $_.Value -in @('Комментарий', 'Комментарии', 'Comment')
        } | Sort-Object { [int]($_.Address -replace '[A-Z]', '') })
        for ($i = 0; $i -lt $headers.Count; $i++) {
            $header = $headers[$i]
            $letters = $header.Address -replace '[0-9]', ''
            $first = [int]($header.Address -replace '[A-Z]', '') + 1
            $last = $bounds.Bottom
            if ($i + 1 -lt $headers.Count) { $last = [int]($headers[$i + 1].Address -replace '[A-Z]', '') - 1 }
            for ($r = $first; $r -le $last; $r++) {
                $address = $letters + $r
                $source = $templateCells[$address]
                if ($null -eq $source -or $registeredCells.ContainsKey($address) -or $source.Formula -or
                    [string]::IsNullOrWhiteSpace([string]$source.Value)) { continue }
                if ($name -eq 'rngSystemSettings' -and
                    [string]$templateCells[((ConvertTo-ExcelColumn $bounds.Left) + $r)].Value -notmatch '^[A-Za-z]+\.') { continue }
                $cell = Get-MigrationCell $name $address
                Add-MigrationRange $cell
                $commentPlans.Add([pscustomobject]@{ Cell=$cell; Value=[string]$source.Value })
            }
        }
    }
}

# Проверяет смысл всех перенесенных списков после записи и reopen. Для
# динамических списков проверяются сохраненные ссылки, а не чужие default ID.
function Test-MigratedValidation {
    foreach ($plan in $validationPlans) {
        $cell = $config.Range($plan.Address)
        if ([int]$cell.Validation.Type -ne 3 -or [string]$cell.Validation.Formula1 -cne $plan.Formula -or
            [bool]$cell.Validation.IgnoreBlank -ne $plan.IgnoreBlank -or -not [bool]$cell.Validation.InCellDropdown) {
            throw "После миграции изменилась validation $($plan.Id) в $($cell.Address())."
        }
    }
}

try {
    $lines.Add("SOURCE: $path; SHA256=$((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash)")
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($path)
    $config = $book.Worksheets.Item('Config')
    $bounds = $config.UsedRange.Address()
    $settings = $book.Names.Item('rngSystemSettings').RefersToRange
    $loads = $book.Names.Item('rngLoadCombinations').RefersToRange
    if ($loads.Columns.Count -lt 7 -or $loads.Rows.Count -lt 2) { throw 'Нет полного блока сочетаний с колонкой пути.' }
    if ([string]$loads.Cells.Item(1, 6).Value2 -notin @('CapacityLoadPath', 'LoadPath')) { throw 'Не распознан заголовок колонки пути.' }
    $loadBefore = Get-LoadInputSignature $loads
    Add-MigrationRange $loads.Cells.Item(1, 6)

    $obsolete = $null
    for ($r = 2; $r -le $settings.Rows.Count; $r++) {
        if ([string]$settings.Cells.Item($r, 1).Value2 -eq 'SLS.Crack.InitiationLoadPath') {
            if ($null -ne $obsolete) { throw 'Несколько obsolete-строк пути трещинообразования.' }
            $obsolete = $settings.Cells.Item($r, 1).Resize(1, $settings.Columns.Count)
            Add-MigrationRange $obsolete
            $lines.Add('REMOVED_SETTING: ' + (ConvertTo-Json @{address=$obsolete.Address(); previous=$obsolete.Formula} -Depth 4 -Compress))
        }
    }

    $title = $config.Cells.Find('Контрольные точки диаграмм*', $config.Cells.Item(1, 1), -4163, 1, 1, 1, $false, $false, $false)
    if ($null -eq $title) { throw 'Не найдена контрольная таблица диаграмм.' }
    Add-MigrationRange $title.Resize(54, 6)
    $rect = $book.Names.Item('rngRectSetGeometry').RefersToRange
    foreach ($r in @(21, 23, 25, 27)) {
        foreach ($c in @(3, 4, 6, 7)) { Add-MigrationRange $rect.Cells.Item(($r + 1), $c) }
    }

    $lambda = [char]0x03BB
    $options = @('Auto', "$lambda*Mx", "$lambda*My", "$lambda*Mxy", "$lambda*N", "$lambda*NMxy")
    $validationSource = [string]$loads.Cells.Item(2, 6).Validation.Formula1
    if (-not $validationSource.StartsWith('=')) { throw 'Колонка пути не использует распознаваемый список validation.' }
    $list = $config.Range($validationSource.Substring(1))
    if ($list.Columns.Count -ne 1 -or [string]$list.Worksheet.Name -ne 'Config') { throw 'Список пути должен быть вертикальным на Config.' }
    $list = $list.Cells.Item(1, 1).Resize($options.Count, 1)
    for ($r = 1; $r -le $list.Rows.Count; $r++) {
        $old = [string]$list.Cells.Item($r, 1).Value2
        if (-not [string]::IsNullOrWhiteSpace($old) -and $old -notin $options) { throw "Служебный список пути содержит постороннее значение в $($list.Cells.Item($r, 1).Address())." }
    }
    Add-MigrationRange $list
    Prepare-CatalogValidation
    Prepare-CatalogComments
    $before = Get-PreservedConfigSignature $config $bounds
    if ($VerifyIdempotence) { $allBefore = Get-PreservedConfigSignature $config $bounds $true }

    # Конфликт общих RectSet-настроек проверяет штатный helper до их изменения.
    foreach ($record in @(Set-RectSetSharedSelectorLayout $rect)) {
        $lines.Add('RECTSET_SHARED_SELECTOR: ' + (ConvertTo-Json $record -Compress))
    }
    if ($null -ne $obsolete) { $obsolete.Validation.Delete(); $obsolete.ClearContents() }
    $loads.Cells.Item(1, 6).Value2 = 'LoadPath'
    for ($r = 1; $r -le $options.Count; $r++) { $list.Cells.Item($r, 1).Value2 = $options[$r - 1] }
    $paths = $loads.Offset(1, 5).Resize($loads.Rows.Count - 1, 1)
    $paths.Validation.Delete()
    $paths.Validation.Add(3, 1, 1, ('=' + $list.Address($true, $true)))
    $paths.Validation.IgnoreBlank = $true
    $paths.Validation.InCellDropdown = $true
    Add-MaterialDiagramControlTables $config ($title.Row + 1) $title.Column
    foreach ($plan in $validationPlans) {
        $plan.Cell.Validation.Delete()
        $plan.Cell.Validation.Add(3, 1, 1, $plan.Formula)
        $plan.Cell.Validation.IgnoreBlank = $plan.IgnoreBlank
        $plan.Cell.Validation.InCellDropdown = $true
    }
    foreach ($plan in $commentPlans) { $plan.Cell.Value2 = $plan.Value }
    Test-MigratedValidation
    $lines.Add("CATALOG_VALIDATION: count=$($validationPlans.Count); static choices updated; dynamic references preserved")
    $lines.Add("CATALOG_COMMENTS: count=$($commentPlans.Count); load-case user comments preserved")
    Apply-ConfigUserInputAlignment $book
    Apply-ConfigCommentColumnAlignment $book
    Apply-ConfigShortColumnAlignment $book
    if ($loadBefore -cne (Get-LoadInputSignature $loads)) { throw 'Миграция изменила исходные сочетания.' }
    $after = Get-PreservedConfigSignature $config $bounds
    if ($before -ne $after) { throw 'Изменились значения или формулы вне разрешенных блоков миграции.' }
    if ($VerifyIdempotence -and $allBefore -ne (Get-PreservedConfigSignature $config $bounds $true)) {
        throw 'Повторная миграция изменила значения или формулы Config.'
    }
    $lines.Add("CONFIG_PRESERVED: True; bounds=$bounds; approvedCells=$($exceptions.Count); signature=$before")
    $lines.Add('LOAD_INPUTS_PRESERVED: True; blank/explicit paths retained; six validation choices')
    $book.Save()
    $book.Close($false)
    $book = $excel.Workbooks.Open($path, 0, $true)
    $config = $book.Worksheets.Item('Config')
    Test-MigratedValidation
    if ($before -ne (Get-PreservedConfigSignature $book.Worksheets.Item('Config') $bounds)) { throw 'Сохранение изменило защищенные исходные данные.' }
    if ($loadBefore -cne (Get-LoadInputSignature $book.Names.Item('rngLoadCombinations').RefersToRange)) { throw 'После reopen изменились сочетания.' }
    if ($VerifyIdempotence) {
        if ($allBefore -ne (Get-PreservedConfigSignature $config $bounds $true)) { throw 'Повторная миграция/save изменила Config.' }
        $lines.Add("CONFIG_IDEMPOTENT: True; entireConfigFormulaSignature=$allBefore")
    }
    $lines.Add('CONFIG_SAVE_REOPEN: True')
    $book.Close($false)
    $book = $null
} catch {
    $failed = $true
    $lines.Add('SCRIPT ERROR: ' + $_.Exception.Message)
    $lines.Add('SCRIPT ERROR POSITION: ' + $_.InvocationInfo.PositionMessage)
    $lines.Add('SCRIPT ERROR STACK: ' + $_.ScriptStackTrace)
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null }
    $lines | Set-Content -LiteralPath $report -Encoding UTF8
}
Restore-WorkbookPrintAreas $path $printAreas
$lines
if ($failed) { exit 1 }
