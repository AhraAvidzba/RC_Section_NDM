# Обновляет справку только в изолированной Audit03-книге. Проверяет фактический
# лист, ссылки и неизменность input-значений/формул/validation. Подписи единиц
# контрольных диаграмм связываются с INPUT Stress. По явному флагу
# разделяет настройки противоположных граней RectSet, сохраняя прежний выбор;
# не подтверждает
# нормативную трассировку либо пиксельную визуальную приемку.
param(
    [Parameter(Mandatory=$true)][string]$WorkbookPath,
    [Parameter(Mandatory=$true)][string]$ReportPath,
    [string]$RegistryPath = 'docs/regression/Audit03/config_field_registry_2026-10-02.csv',
    [switch]$UpdateRectSetIndependentSelectors
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$bookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
$allowed = [IO.Path]::GetFullPath((Join-Path $root 'docs/regression/Audit03')) + [IO.Path]::DirectorySeparatorChar
$postAuditAllowed = [IO.Path]::GetFullPath((Join-Path $root 'docs/regression/PostAudit03')) + [IO.Path]::DirectorySeparatorChar
if (-not $bookPath.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase) -and
    -not $bookPath.StartsWith($postAuditAllowed, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Обновление справки допускается только в изолированной Audit03-книге.'
}
$fields = @(Import-Csv -LiteralPath (Join-Path $root $RegistryPath) | Where-Object Role -eq 'UserInput')
if ($fields.Count -eq 0) { throw 'Реестр не содержит пользовательских полей.' }
$printAreas = @(Get-WorkbookPrintAreas $bookPath)
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
# явно согласованного центрирования RectSet-селекторов; их значения,
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

# Меняет только подпись размерности служебной таблицы и ссылки осей графиков.
# Строки материалов, численные формулы точек и пользовательский ввод не трогает;
# обе ячейки находятся по существующим заголовкам и именованной таблице единиц.
function Update-MaterialControlUnitCaptions([object]$Book) {
    $sheet = $Book.Worksheets.Item('Config')
    $title = $sheet.Cells.Find('Контрольные точки диаграмм*', $sheet.Cells.Item(1,1), -4163, 1, 1, 1, $false, $false, $false)
    if ($null -eq $title) { throw 'На Config не найдена контрольная таблица диаграмм.' }
    $units = $Book.Names.Item('rngUnitSettings').RefersToRange
    $stressCell = $null
    for ($row = 2; $row -le $units.Rows.Count; $row++) {
        if ([string]$units.Cells.Item($row,1).Value2 -eq 'Stress') { $stressCell = $units.Cells.Item($row,2); break }
    }
    if ($null -eq $stressCell) { throw 'В rngUnitSettings не найдена INPUT-единица Stress.' }
    $caption = $title.Offset(1,5)
    $caption.Formula = '="σ, "&' + $stressCell.Address($true,$true)
    $axisFormula = "='" + $sheet.Name.Replace("'", "''") + "'!" + $caption.Address($true,$true)
    foreach ($name in @('chMaterialConcreteDiagram','chMaterialSteelDiagram')) {
        $axis = $sheet.ChartObjects($name).Chart.Axes(2)
        $axis.HasTitle = $true
        $axis.AxisTitle.Formula = $axisFormula
        $lines.Add("MATERIAL_CONTROL_CAPTION: chart=$name; formula=$axisFormula")
    }
    $sheet.Calculate()
    Assert-Help 'materialControlCaption' ([string]$caption.Value2 -eq ('σ, '+[string]$stressCell.Value2)) "cell=$($caption.Address($false,$false)); unit=$($stressCell.Value2)"
}

# Сравнивает область по номеру листа и формуле XML. Канонизация имени
# Print_Area и удаление одинакового дубликата не меняют саму область печати.
function Get-PrintAreaSignature([object[]]$Areas) {
    $records = foreach ($area in $Areas) {
        [xml]$node = $area.Xml
        [string]$area.Sheet + '|' + $node.DocumentElement.InnerText
    }
    return (@($records | Sort-Object -Unique) -join "`n")
}

# Проверяет отсутствие объединений и отдельный dropdown каждой грани.
# Все 32 ячейки являются независимыми пользовательскими вводами.
function Test-RectSetIndependentLayout([object]$Book) {
    $target = $Book.Names.Item('rngRectSetGeometry').RefersToRange
    foreach ($firstRow in @(21, 23, 25, 27)) {
        foreach ($column in @(3, 4, 6, 7)) {
            $first = $target.Cells.Item($firstRow, $column)
            $second = $target.Cells.Item(($firstRow + 1), $column)
            foreach ($cell in @($first, $second)) {
                Assert-Help ("rectset.unmerged."+$cell.Address()) (-not $cell.MergeCells) 'Independent physical side'
                Assert-Help ("rectset.dropdown."+$cell.Address()) ($cell.Validation.Type -eq 3 -and $cell.Validation.InCellDropdown) ([string]$cell.Validation.Formula1)
                Assert-Help ("rectset.centered."+$cell.Address()) ($cell.HorizontalAlignment -eq -4108 -and $cell.VerticalAlignment -eq -4108) 'Independent input is centered'
            }
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
        # Общие абзацы объединены от A, а таблицы настроек содержат текст в B.
        [void]$text.AppendLine(([string]$data[$r,1]))
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
    # Генератор разбивает абзац на несколько объединенных строк. Для поиска
    # смысловой фразы восстанавливаем пробелы, не меняя слова или их порядок.
    $body = [regex]::Replace($text.ToString(), '\s+', ' ').Trim()
    Assert-Help 'extensionAllStates' ($body.Contains('PreCrackState/PostCrackState/CurrentCrackedState')) 'Global extension scope'
    Assert-Help 'extensionDoesNotChangePhysicalNodes' ($body.Contains('Исходные физические точки, сопротивления, касательные, плато и предельные деформации не меняются')) 'Physical diagram preserved'
    Assert-Help 'profileBooleanInputContract' ($body.Contains('Все четыре переключателя Calculation.* обязательны')) 'Explicit invalid value is not No'
    Assert-Help 'subdivisionsInputContract' ($body.Contains('Mesh.BoundarySubdivisions')) 'Mesh setting has help'
    Assert-Help 'worstCriterionInputContract' ($body.Contains('неизвестное значение этого критерия') -and $body.Contains('InputErr')) 'Invalid criterion cannot silently select another check'
    Assert-Help 'rectsetIndependentSelectors' ($body.Contains('задаются независимо в каждой строке физической грани') -and $body.Contains('Выбор одной грани не изменяет противоположную')) 'Separate selectors for each physical side'
    Assert-Help 'unitChoiceInputContract' ($body.Contains('Пустой выбор, TODO или неподдержанная единица являются ошибкой Config')) 'Explicit invalid units do not become defaults'
    Assert-Help 'autoCADExportAlwaysMillimetres' ($body.Contains('Геометрия экспортируется в AutoCAD всегда в миллиметрах') -and $body.Contains('одна единица AutoCAD соответствует 1 мм')) 'AutoCAD export scale is fixed, not OUTPUT length'
    Assert-Help 'autoCADExportUnitsException' ($body.Contains('экспорт в AutoCAD всегда выполняется в мм независимо от OUTPUT') -and $body.Contains('это не меняет единицы напряжений')) 'Units help distinguishes geometry scale and result labels'
    Assert-Help 'autoCADImportAreaThresholdUnits' ($body.Contains('INPUT-площадь применяется к порогу AutoCAD.Import.MinArea') -and $body.Contains('25 mm2, 0.25 cm2 и 0.000025 m2')) 'User threshold is converted; drawing remains mm'
    Assert-Help 'autoCADImportGeometryUnitPowers' ($body.Contains('координаты и размеры - мм, площади - мм2, инерции - мм4') -and $body.Contains('INPUT не задает коэффициент масштаба геометрии')) 'Drawing geometry has fixed mm/mm2/mm4 dimensions'
    Assert-Help 'autoCADRegionImportContract' ($body.Contains('Контракт импорта Region: координаты X/Y центра') -and $body.Contains('Area - как мм2, моменты инерции и произведение инерции - как мм4') -and $body.Contains('Программа не определяет масштаб по INSUNITS')) 'Region data have fixed units, not automatic drawing scale discovery'
    Assert-Help 'autoCADImportAreaThresholdInputContract' ($body.Contains('Нулевой порог отключает отбрасывание по малой площади') -and $body.Contains('скрытого значения по умолчанию нет')) 'Missing/invalid threshold is not replaced with default'
    Assert-Help 'importedSnapshotUnitChanges' ($body.Contains('Импортированная геометрия записывается в Results в выбранных на момент импорта OUTPUT-единицах') -and $body.Contains('восстановит геометрию по единицам старого снимка')) 'Import and recalculation use their own recorded OUTPUT units'
    Assert-Help 'inputChangeDoesNotRescaleImportedGeometry' ($body.Contains('смена INPUT без перевода уже введенных чисел меняет их физический смысл') -and $body.Contains('не отфильтровывает заново сохраненный снимок')) 'Current input values and preserved imported geometry are distinct'
    Assert-Help 'signsDoNotFlipStressOrStrain' ($body.Contains('Пользовательские знаки N/Mx/My не меняют знак Stress и Strain')) 'Material tension/compression and strain plane keep internal signs'
    Assert-Help 'crackSettingsRequiredInput' ($body.Contains('Для расчета по СП 63 ячейка все равно должна содержать положительное число') -and $body.Contains('пустота или ошибка ввода не заменяются значением по умолчанию')) 'Active SP63 coefficient input does not become a default'
    Assert-Help 'crackInactiveUserCoefficientContract' ($body.Contains('При SLS.Crack.Phi3Mode = Auto это значение не подставляется в формулу') -and $body.Contains('Ячейка все равно должна содержать положительный коэффициент') -and $body.Contains('При SLS.Crack.PsiMode = Auto или AlwaysCalc эта ячейка не задает итоговый коэффициент') -and $body.Contains('Для расчета по СП 63 ячейка все равно должна содержать положительное число')) 'Phi3 and PsiS remain valid inputs in automatic SP63 modes'
    Assert-Help 'mandatoryInputErrorNavigation' ($body.Contains('Сообщение называет настройку и фактическую ячейку Config') -and $body.Contains('Если строка удалена, адрес не угадывается') -and $body.Contains('в диспетчере имен Excel')) 'Input diagnostics explain actual cell, missing row and damaged named table repair'
    Assert-Help 'circleActiveRebarInput' ($body.Contains('Пустой отступ не заменяется значением шаблона') -and $body.Contains('Для каждого активного дополнительного ряда Loc2row/Loc3row обязателен')) 'Circle active cover and row position are required'
    Assert-Help 'inactiveGeometryFormulaErrors' ($body.Contains('ошибка в неактивном поле или невыбранной форме не мешает построению')) 'Inactive geometry formula error is not a failed active input'
    Assert-Help 'inactiveRebarFaceOffsets' ($body.Contains('as отключенной наружной грани не читается') -and $body.Contains('as выключенной наружной грани не читается') -and $body.Contains('ее as/t и параметры дополнительных рядов не читаются')) 'Rounded, Hollow and RectSet inactive face contracts are explicit'
    Assert-Help 'openingCoverStillRequired' ($body.Contains('Отступ as граней Opening нужен также для ограничения проекций соседних внутренних граней') -and $body.Contains('даже при отсутствии собственного ряда')) 'Opening cover retains a geometric consumer when its own row is disabled'
    Assert-Help 'notCrackedWidthIsNotComputedZero' ($body.Contains('Подтвержденный NotCracked исключает расчет ширины') -and $body.Contains('статус проверки N/A, численная ячейка a_crc остается пустой') -and $body.Contains('это отдельный случай, а не доказательство отсутствия трещины')) 'Uncomputed crack width is distinguished from a calculated zero'
    Assert-Help 'notCrackedCurrentLoadIsNotFailedFutureProbe' ($body.Contains('Если допустимое равновесие при λ = 1 не достигает критерия образования трещины') -and $body.Contains('текущее сочетание остается NotCracked')) 'A physically invalid larger probe does not fail a proven uncracked current load'
    Assert-Help 'fixedPathNotCrackedIsSuccess' ($body.Contains('является успешной проверкой NotCracked, в том числе для фиксированного пути') -and -not $body.Contains('либо для фиксированного пути доказана недостижимость нужного критерия')) 'Confirmed absence of a normal crack is not BaseFail'
    Assert-Help 'activeAnnotationInputContract' ($body.Contains('Enabled обязателен для каждой группы') -and $body.Contains('При No ее параметры не читаются, включая ошибки формул')) 'Only active annotation parameters are required'
    Assert-Help 'annotationBlackRgbContract' ($body.Contains('0,0,0 означает настоящий черный цвет, а не отсутствие настройки') -and $body.Contains('дробные компоненты не округляются')) 'Zero is a real color; invalid RGB is not silently repaired'
    Assert-Help 'annotationVisualLimits' ($body.Contains('4–180 pt у размеров и 2–120 pt у арматуры') -and $body.Contains('5–28 pt для читаемости')) 'Screen clamps are disclosed, not physical geometry'
    Assert-Help 'snapshotAnnotationInputContract' ($body.Contains('Отсутствующий или пустой необязательный блок аннотаций допустим') -and $body.Contains('ошибочное значение не превращается в ноль') -and $body.Contains('GeometryInterpretationStatus;') -and -not $body.Contains('GeometryInterpretationStatus либо ShapeType')) 'Existing damaged snapshot requires repair, not a silent zero'
    Assert-Help 'snapshotGeometryNumbersContract' ($body.Contains('координаты, площадь и диаметр арматуры должны содержать числа') -and $body.Contains('Текст с числовым началом, логическое значение и ошибка формулы не принимаются как число') -and $body.Contains('фактическую ячейку Results')) 'Geometry restoration rejects corrupted numbers without losing source address'
    Assert-Help 'axisFontInputContract' ($body.Contains('При Plot.AxisLabelsEnabled=Yes значение обязательно и больше нуля') -and $body.Contains('При No высота подписей не читается')) 'Active axis font is required; disabled font is not read'
    Assert-Help 'gradientDisabledStillUsesStateColor' ($body.Contains('No отключает градации интенсивности: работающие элементы получают один цвет')) 'No gradient does not mean no material color'
    Assert-Help 'materialControlStressUnits' ($body.Contains('Напряжения σ в контрольной таблице и на ее графиках показаны в текущей INPUT-единице Stress') -and $body.Contains('OUTPUT Stress на этот контрольный блок не влияет')) 'Formula control values and axis captions use INPUT Stress'
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
    if ($UpdateRectSetIndependentSelectors) {
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
        $lines.Add("RECTSET_PRESERVATION_SCOPE: originalFields=$originalCount; unchangedFields=$($fields.Count); secondSides=16; fullBefore=$fullBefore")
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
        $migration = @(Set-RectSetIndependentSelectorLayout $target)
        foreach ($record in $migration) {
            $lines.Add('RECTSET_INDEPENDENT_SELECTOR: '+(ConvertTo-Json $record -Compress))
        }
        Assert-Help 'rectset.migrationRecords' ($migration.Count -eq 16) "count=$($migration.Count)"
        Test-RectSetIndependentLayout $book
        $layoutOnce = Get-InputSignature $book
        Set-RectSetIndependentSelectorLayout $target | Out-Null
        Assert-Help 'rectset.layoutIdempotent' ($layoutOnce -eq (Get-InputSignature $book)) "signature=$layoutOnce"
    } else {
        $before = Get-InputSignature $book
        $beforeRecords = @($script:lastInputRecords)
    }
    foreach ($record in @(Update-SystemSettingUnitCaptions $book)) {
        $lines.Add('SYSTEM_UNIT_CAPTION: '+(ConvertTo-Json $record -Compress))
    }
    Add-SettingsInstructions $book $book.Worksheets.Item('Config') $book.Worksheets.Item('Справка')
    Apply-ConfigNamedRangeBorders $book
    Update-MaterialControlUnitCaptions $book
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
    if ($UpdateRectSetIndependentSelectors) { Test-RectSetIndependentLayout $book }
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
    Restore-WorkbookPrintAreas $bookPath $printAreas
    $restoredPrintAreas = @(Get-WorkbookPrintAreas $bookPath)
    $expectedPrintAreas = Get-PrintAreaSignature $printAreas
    $actualPrintAreas = Get-PrintAreaSignature $restoredPrintAreas
    Assert-Help 'printAreasPreserved' ($expectedPrintAreas -eq $actualPrintAreas) "count=$($restoredPrintAreas.Count)"
    $lines | Set-Content -LiteralPath (Join-Path $root $ReportPath) -Encoding UTF8
}
if ($script:failed -gt 0) { exit 1 }
