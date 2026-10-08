# Обновляет только справку в изолированной копии и проверяет сохранность книги.
# Не запускает VBA, расчеты, миграцию настроек или пересборку таблиц Config.
param(
    [Parameter(Mandatory=$true)][string]$WorkbookPath,
    [Parameter(Mandatory=$true)][string]$ReportDirectory,
    [switch]$Render,
    [switch]$VerifyOnly,
    [string]$ReferenceWorkbookPath
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$path = (Resolve-Path -LiteralPath $WorkbookPath).Path
$referencePath=$path
if ($VerifyOnly) {
    if (-not $ReferenceWorkbookPath) { throw 'Для проверки укажите исходную книгу ReferenceWorkbookPath.' }
    $referencePath=(Resolve-Path -LiteralPath $ReferenceWorkbookPath).Path
}
$allowed = [IO.Path]::GetFullPath((Join-Path $root 'docs/regression/HelpReadability')) + [IO.Path]::DirectorySeparatorChar
if (-not $path.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Обновление допускается только в копии внутри docs/regression/HelpReadability.'
}
$directory = [IO.Path]::GetFullPath($ReportDirectory)
if (-not $directory.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Отчет должен находиться внутри docs/regression/HelpReadability.'
}
New-Item -ItemType Directory -Path $directory -Force | Out-Null
$records = New-Object 'System.Collections.Generic.List[object]'
$script:failures = 0
$excel = $null
$book = $null

function Get-ObjectHash([object]$Value) {
    $payload = ConvertTo-Json -InputObject $Value -Depth 12 -Compress
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($payload))).Replace('-', '') }
    finally { $sha.Dispose() }
}

function Get-VbaHash([string]$File) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($File)
    try {
        $stream = $zip.GetEntry('xl/vbaProject.bin').Open()
        $sha = [Security.Cryptography.SHA256]::Create()
        try { return [BitConverter]::ToString($sha.ComputeHash($stream)).Replace('-', '') }
        finally { $stream.Dispose(); $sha.Dispose() }
    } finally { $zip.Dispose() }
}

# Excel может перепаковать VBA при сохранении даже без вызова макросов.
# Справка его не меняет: возвращаем исходные макро-части побайтно.
function Get-MacroParts([string]$File) {
    $zip=[IO.Compression.ZipFile]::OpenRead($File)
    $parts=@{}
    try {
        foreach ($entry in $zip.Entries) {
            if ($entry.FullName -notlike 'xl/vbaProject*') { continue }
            $stream=$entry.Open()
            $memory=New-Object IO.MemoryStream
            try { $stream.CopyTo($memory); $parts[$entry.FullName]=$memory.ToArray() }
            finally { $stream.Dispose(); $memory.Dispose() }
        }
    } finally { $zip.Dispose() }
    return $parts
}

function Get-CalculationProperties([string]$File) {
    $zip=[IO.Compression.ZipFile]::OpenRead($File)
    try {
        $reader=New-Object IO.StreamReader($zip.GetEntry('xl/workbook.xml').Open())
        try { [xml]$document=$reader.ReadToEnd() } finally { $reader.Dispose() }
        $node=$document.SelectSingleNode("//*[local-name()='calcPr']")
        if ($node) { return $node.OuterXml }
        return ''
    } finally { $zip.Dispose() }
}

# Временный ручной режим Excel нужен только для редактирования. Настройки
# пересчета опубликованной книги возвращаются без запуска ее вычислений.
function Restore-CalculationProperties([string]$File, [string]$Properties) {
    $zip=[IO.Compression.ZipFile]::Open($File,[IO.Compression.ZipArchiveMode]::Update)
    try {
        $entry=$zip.GetEntry('xl/workbook.xml')
        $reader=New-Object IO.StreamReader($entry.Open())
        try { [xml]$document=$reader.ReadToEnd() } finally { $reader.Dispose() }
        $old=$document.SelectSingleNode("//*[local-name()='calcPr']")
        if ($Properties.Length) {
            [xml]$source=$Properties
            $copy=$document.ImportNode($source.DocumentElement,$true)
            if ($old) { $document.DocumentElement.ReplaceChild($copy,$old) | Out-Null }
            else { $document.DocumentElement.AppendChild($copy) | Out-Null }
        } elseif ($old) { $document.DocumentElement.RemoveChild($old) | Out-Null }
        $stream=$entry.Open()
        $stream.SetLength(0)
        $writer=New-Object IO.StreamWriter($stream,(New-Object Text.UTF8Encoding($false)))
        try { $document.Save($writer) } finally { $writer.Dispose() }
    } finally { $zip.Dispose() }
}

function Restore-MacroParts([string]$File, [object]$Parts) {
    $zip=[IO.Compression.ZipFile]::Open($File,[IO.Compression.ZipArchiveMode]::Update)
    try {
        foreach ($name in $Parts.Keys) {
            $entry=$zip.GetEntry($name)
            if ($null -eq $entry) { throw "Макро-часть потеряна: $name" }
            $stream=$entry.Open()
            try { $stream.SetLength(0); $stream.Write($Parts[$name],0,$Parts[$name].Length) }
            finally { $stream.Dispose() }
        }
    } finally { $zip.Dispose() }
}

function Assert-Help([string]$Name, [bool]$Passed, [object]$Detail) {
    if (-not $Passed) { $script:failures++ }
    $records.Add([ordered]@{check=$Name; passed=$Passed; detail=$Detail})
    if (-not $Passed) { Write-Output "FAILED: $Name" }
}

# Полные значения и формулы всех листов, кроме обновляемой справки. Проверка
# не ограничивается старым реестром полей и захватывает новые настройки.
function Get-DataState([object]$Book) {
    $state = [ordered]@{}
    foreach ($sheet in $Book.Worksheets) {
        if ([string]$sheet.Name -eq 'Справка') { continue }
        $used = $sheet.UsedRange
        $widths = @()
        for ($column = $used.Column; $column -lt $used.Column + $used.Columns.Count; $column++) {
            $widths += [double]$sheet.Columns.Item($column).ColumnWidth
        }
        $state[[string]$sheet.Name] = [ordered]@{
            address=$used.Address(); rows=$used.Rows.Count; columns=$used.Columns.Count
            formulas=$used.Formula; values=$used.Value2; widths=$widths
        }
    }
    $names = @()
    foreach ($name in $Book.Names) {
        if ([string]$name.Name -match 'Print_Area$') { continue }
        $names += ([string]$name.Name + '|' + [string]$name.RefersTo)
    }
    $state['Names'] = @($names | Sort-Object)
    return $state
}

function Get-ValidationState([object]$Sheet) {
    $cells = $Sheet.UsedRange.SpecialCells(-4174)
    $state = @()
    foreach ($cell in $cells.Cells) {
        $v = $cell.Validation
        $source = $null
        if ([int]$v.Type -eq 3 -and [string]$v.Formula1 -like '=*') {
            $source = $Sheet.Range(([string]$v.Formula1).Substring(1)).Value2
        }
        $state += [ordered]@{
            address=$cell.Address(); type=[int]$v.Type; formula1=[string]$v.Formula1
            formula2=[string]$v.Formula2; dropdown=[bool]$v.InCellDropdown
            ignoreBlank=[bool]$v.IgnoreBlank; alert=[int]$v.AlertStyle; source=$source
            numberFormat=[string]$cell.NumberFormat; horizontal=$cell.HorizontalAlignment
            vertical=$cell.VerticalAlignment; merged=[bool]$cell.MergeCells
        }
    }
    return ,$state
}

function Get-GuideText([object]$Sheet, [int]$First, [int]$Last) {
    $data = $Sheet.Range($Sheet.Cells.Item($First,1),$Sheet.Cells.Item($Last,2)).Value2
    $builder = New-Object Text.StringBuilder
    for ($row = 1; $row -le $data.GetLength(0); $row++) {
        [void]$builder.Append(' ' + [string]$data[$row,1] + ' ' + [string]$data[$row,2])
    }
    return [regex]::Replace($builder.ToString(), '\s+', ' ').Trim()
}

function Test-ActualHelp([object]$Book, [string]$Stage) {
    $guide = $Book.Worksheets.Item('Справка')
    $data = $guide.UsedRange.Value2
    $titles = [ordered]@{}
    foreach ($title in @('Подготовка геометрии в AutoCAD и импорт','Поиск НДС текущего сочетания',
        'Поиск предельной несущей способности','Расчет ширины раскрытия нормальных трещин',
        'Расчет продольного изгиба и устойчивости','Справка по настройкам листа Config',
        'Справка по результатам расчета листа Results','СП 35: расчет ширины нормальных трещин')) {
        for ($row=1; $row -le $data.GetLength(0); $row++) {
            if ([string]$data[$row,1] -eq $title) { $titles[$title]=$row; break }
        }
        Assert-Help "$Stage.section.$title" ($titles.Contains($title)) $titles[$title]
    }
    Assert-Help "$Stage.autoCADFirst" ($titles['Подготовка геометрии в AutoCAD и импорт'] -eq 1) 'Первый раздел'
    $body = Get-GuideText $guide 1 $data.GetLength(0)
    $sp35 = Get-GuideText $guide $titles['СП 35: расчет ширины нормальных трещин'] ($titles['Расчет продольного изгиба и устойчивости']-1)
    foreach ($phrase in @('Группа неделима','Частичного включения нет','боковых границ площади взаимодействия',
        'Другие группы тоже будут рассмотрены','напряжение для этой проверки берется только из рассматриваемой группы',
        'Ряд 1 находится у растянутой бетонной границы','SLS.Crack.SP35.NeighborRatioLimit',
        'SLS.Crack.SP35.RowTolerance','SLS.Crack.SP35.RadiusDiameterMode','SLS.Crack.SP35.GroupGapTolerance',
        '[SP35] - правило СП 35','[METHOD] - пояснение','[NDM] - принятое в программе допущение',
        'D_1, D_2','Индекс s обозначает арматуру, cr - трещину','AutoCAD.Export.ExportCrackInteractionContour')) {
        Assert-Help "$Stage.sp35.$phrase" ($sp35.Contains($phrase)) 'Фактический лист'
    }
    Assert-Help "$Stage.sp35NoSP63" ($sp35 -notmatch 'СП 63|ψ_s|PsiS|Phi3') 'Без сравнения с другой методикой'
    $sp35CardCount=0
    for ($row=1; $row -le $data.GetLength(0); $row++) {
        $key=[string]$data[$row,1]
        if ($key -notmatch '^SLS\.Crack\.SP35\.[A-Za-z]+$') { continue }
        $sp35CardCount++
        $heading=$guide.Cells.Item($row,1)
        Assert-Help "$Stage.settingHeading.$key" ($heading.MergeArea.Columns.Count -eq 6 -and -not $heading.WrapText) 'Имя настройки без разрыва строки'
        $last=$row
        for ($next=$row+1; $next -le $data.GetLength(0); $next++) {
            $label=[string]$data[$next,1]
            if ($label.Length -gt 0 -and $label -notmatch '^\d+$') { break }
            $last=$next
        }
        $card=Get-GuideText $guide $row $last
        Assert-Help "$Stage.sp35Card.$key" ($card -notmatch 'СП 63|ψ_s|PsiS|Phi3') 'Только параметры СП 35'
    }
    Assert-Help "$Stage.sp35CardCoverage" ($sp35CardCount -eq 6) $sp35CardCount
    foreach ($phrase in @('1. Заданы только отверстия','2. Задан только наружный контур',
        '3. Заданы наружный контур и отверстия','4. Контуры не заданы','SUBTRACT',
        'AutoCAD.Export.ContourFormat','Фактически полученную площадь взаимодействия можно вывести в AutoCAD')) {
        Assert-Help "$Stage.contract.$phrase" ($body.Contains($phrase)) 'Фактический лист'
    }
    Assert-Help "$Stage.noInternalClasses" ($body -notmatch 'CSectionModel|CUnitSystem|CSectionSolver|shape-builder|snapshot|named-state') 'Пользовательская справка'
    Assert-Help "$Stage.worldCoordinates" ($body.Contains('_UCS') -and $body.Contains('_World') -and $body.Contains('глобальн')) 'Глобальная система координат'
    Assert-Help "$Stage.noOpaqueGroupSum" ($body -notmatch 'i∈g|iEg|gap_ij') 'Простая запись суммы'
    $count=0
    foreach ($link in $Book.Worksheets.Item('Config').Hyperlinks) {
        if ([string]$link.SubAddress -match "^'?Справка'?!([A-Z]+[0-9]+)$") {
            $target=$guide.Range($Matches[1])
            Assert-Help "$Stage.link.$($link.Range.Address())" (-not [string]::IsNullOrWhiteSpace([string]$target.Value2)) $link.SubAddress
            $count++
        }
    }
    Assert-Help "$Stage.linkCoverage" ($count -eq $script:originalHelpLinkCount -and $count -gt 100) @{before=$script:originalHelpLinkCount; after=$count}
    $borderCount=0
    foreach ($name in @('rngSystemSettings','rngUnitSettings','rngSignConventionSettings','rngCalculationProfiles',
        'rngConcreteMaterialParameters','rngSteelMaterialParameters','rngPlotAnnotationSettings','rngLoadCombinations',
        'rngCircleGeometry','rngRoundedRectangleGeometry','rngHollowRectangleGeometry','rngRectSetGeometry')) {
        $range=$Book.Names.Item($name).RefersToRange
        for ($row=1; $row -le $range.Rows.Count; $row++) {
            $cell=$range.Cells.Item($row,$range.Columns.Count)
            $edge=$cell
            if ($cell.MergeCells) { $edge=$cell.MergeArea }
            Assert-Help "$Stage.border.$name.$row" ([int]$edge.Borders.Item(10).LineStyle -ne -4142) $cell.Address()
            $borderCount++
        }
    }
    Assert-Help "$Stage.borderCoverage" ($borderCount -gt 250) $borderCount
    $rect=$Book.Names.Item('rngRectSetGeometry').RefersToRange
    foreach ($row in 21..28) {
        foreach ($column in @(3,4,6,7)) {
            $cell=$rect.Cells.Item($row,$column)
            Assert-Help "$Stage.rectset.$row.$column" (-not $cell.MergeCells -and $cell.Validation.Type -eq 3 -and $cell.Validation.InCellDropdown) $cell.Validation.Formula1
            $source=$cell.Worksheet.Range(([string]$cell.Validation.Formula1).Substring(1)).Value2
            $options=(@($source) | ForEach-Object { [string]$_ }) -join '|'
            $expected='Stacked|SideBySide'
            if ($column -in @(4,7)) { $expected='EachBar|EverySecondBar' }
            Assert-Help "$Stage.rectsetOptions.$row.$column" ($options -eq $expected) $options
        }
    }
    Assert-Help "$Stage.formulaPictures" ($guide.Shapes.Count -gt 50) $guide.Shapes.Count
    $guide.Outline.ShowLevels(8) | Out-Null
    foreach ($shape in $guide.Shapes) {
        if ([string]$shape.Name -notlike 'NDMHelpFormula_*') { continue }
        $cell=$shape.TopLeftCell
        $row=$guide.Rows.Item($cell.Row)
        $inside=([double]$shape.Top -ge [double]$row.Top) -and
            ([double]$shape.Top + [double]$shape.Height -le [double]$row.Top + [double]$row.Height + 0.5)
        Assert-Help "$Stage.pictureBounds.$($shape.Name)" $inside @{row=$cell.Row; top=$shape.Top; height=$shape.Height; rowHeight=$row.Height}
    }
    $guide.Outline.ShowLevels(1) | Out-Null
    $settingsFirst=$titles['Справка по настройкам листа Config']+1
    $settingsLast=$titles['Справка по результатам расчета листа Results']-1
    $guide.Range("A${settingsFirst}:A${settingsLast}").EntireRow.Hidden=$false
    return $titles
}

# Экспортирует только печатное представление справки. Настройки печати копии
# после рендера не сохраняются; рабочая книга и чертеж не затрагиваются.
function Export-GuidePages([object]$Book, [object]$Titles) {
    $sheet=$Book.Worksheets.Item('Справка')
    $sheet.Outline.ShowLevels(8)
    $starts=@('Подготовка геометрии в AutoCAD и импорт','Поиск НДС текущего сочетания',
        'Поиск предельной несущей способности','Расчет ширины раскрытия нормальных трещин',
        'Расчет продольного изгиба и устойчивости','Справка по настройкам листа Config',
        'Справка по результатам расчета листа Results')
    $names=@('AutoCAD','DirectState','Capacity','Cracks','Stability','Config','Results')
    for ($i=0; $i -lt $starts.Count; $i++) {
        $first=[int]$Titles[$starts[$i]]
        $last=$sheet.UsedRange.Rows.Count
        if ($i+1 -lt $starts.Count) { $last=[int]$Titles[$starts[$i+1]]-1 }
        $sheet.PageSetup.PrintArea=$sheet.Range("A${first}:F${last}").Address()
        $sheet.PageSetup.Orientation=2
        $sheet.PageSetup.PaperSize=8
        $sheet.PageSetup.Zoom=$false
        $sheet.PageSetup.FitToPagesWide=1
        $sheet.PageSetup.FitToPagesTall=$false
        $sheet.PageSetup.LeftMargin=15
        $sheet.PageSetup.RightMargin=15
        $sheet.PageSetup.TopMargin=15
        $sheet.PageSetup.BottomMargin=15
        $sheet.ExportAsFixedFormat(0,(Join-Path $directory ($names[$i]+'.pdf')),0,$true,$false)
    }
}

$beforeVba=Get-VbaHash $referencePath
$macroParts=Get-MacroParts $referencePath
$calculationProperties=Get-CalculationProperties $referencePath
$beforeFileHash=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
$printAreas=@(Get-WorkbookPrintAreas $path)
$script:FormulaImageDir=Join-Path $directory 'formula_images'
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false
    $excel.DisplayAlerts=$false
    $excel.EnableEvents=$false
    $excel.AutomationSecurity=3
    $book=$excel.Workbooks.Open($referencePath,0,[bool]$VerifyOnly)
    if (-not $VerifyOnly -and $book.ReadOnly) { throw 'Изолированная копия открыта только для чтения.' }
    $excel.Calculation=-4135
    $before=Get-DataState $book
    $beforeHash=Get-ObjectHash $before
    $script:originalHelpLinkCount=@($book.Worksheets.Item('Config').Hyperlinks | Where-Object { [string]$_.SubAddress -match "^'?Справка'?!([A-Z]+[0-9]+)$" }).Count
    $validation=Get-ValidationState $book.Worksheets.Item('Config')
    $validationHash=Get-ObjectHash $validation
    Assert-Help 'validationCoverage' ($validation.Count -gt 150) $validation.Count
    if ($VerifyOnly) {
        $book.Close($false)
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) | Out-Null
        $book=$excel.Workbooks.Open($path,0,$true)
    } else {
        Add-SettingsInstructions $book $book.Worksheets.Item('Config') $book.Worksheets.Item('Справка')
    }
    $afterHash=Get-ObjectHash (Get-DataState $book)
    Assert-Help 'allDataFormulasNamesWidthsPreserved' ($beforeHash -eq $afterHash) @{before=$beforeHash; after=$afterHash}
    Assert-Help 'allValidationPreserved' ($validationHash -eq (Get-ObjectHash (Get-ValidationState $book.Worksheets.Item('Config')))) $validationHash
    $titles=Test-ActualHelp $book 'updated'
    if ($script:failures) { throw 'Проверка обновления справки не пройдена; копия не сохранена.' }
    if (-not $VerifyOnly) { $book.Save() }
    $book.Close($false)
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) | Out-Null
    $book=$excel.Workbooks.Open($path,0,$true)
    Assert-Help 'allDataSaveReopen' ($beforeHash -eq (Get-ObjectHash (Get-DataState $book))) $beforeHash
    Assert-Help 'allValidationSaveReopen' ($validationHash -eq (Get-ObjectHash (Get-ValidationState $book.Worksheets.Item('Config')))) $validationHash
    $titles=Test-ActualHelp $book 'reopened'
    if ($Render) { Export-GuidePages $book $titles }
} catch {
    $records.Add([ordered]@{check='runtime'; passed=$false; detail=$_.Exception.Message; stack=$_.ScriptStackTrace})
    $script:failures++
    throw
} finally {
    if ($book) { $book.Close($false); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) | Out-Null }
    if ($excel) { $excel.Quit(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if (-not $VerifyOnly) {
        Restore-WorkbookPrintAreas $path $printAreas
        Restore-MacroParts $path $macroParts
        Restore-CalculationProperties $path $calculationProperties
    } else {
        Assert-Help 'readOnlyVerificationUnchanged' ($beforeFileHash -eq (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash) $beforeFileHash
    }
    Assert-Help 'vbaProjectUnchanged' ($beforeVba -eq (Get-VbaHash $path)) $beforeVba
    Assert-Help 'calculationPropertiesUnchanged' ($calculationProperties -eq (Get-CalculationProperties $path)) $calculationProperties
    [ordered]@{workbook=$path; failures=$script:failures; checks=$records.Count; evidence=$records.ToArray()} |
        ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $directory 'HelpVerification.json') -Encoding UTF8
}
if ($script:failures) { exit 1 }
Write-Output "HELP_UPDATE_OK checks=$($records.Count); workbook=$path"
