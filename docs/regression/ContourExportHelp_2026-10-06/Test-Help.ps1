# Проверяет адресность справки в каталоге и в сохраненной Excel-книге.
# Ссылки должны вести к пояснениям конкретного параметра, не к общему дампу.
param(
    [string]$WorkbookPath = 'docs/regression/ContourExportHelp_2026-10-06/Publication/RC_Section_NDM.xlsm',
    [string]$ReportDirectory = ''
)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path
. (Join-Path $root 'tools/build_workbook/SettingsCatalog.ps1')
$keys = @('GroupGapTolerance', 'RowTolerance', 'NeighborRatioLimit', 'RadiusDiameterMode', 'InteractionRadiusMode', 'RebarProfile')
$lines = New-Object 'System.Collections.Generic.List[string]'; $failed = 0; $passed = 0

# Сохраняет все независимые проверки до итогового протокола.
function Assert-Help([string]$Name, [bool]$Condition) {
    if ($Condition) { $script:passed++; $lines.Add("OK: $Name") }
    else { $script:failed++; $lines.Add("FAIL: $Name") }
}

$entries = @{}
foreach ($group in Get-SystemSettingsCatalog) { foreach ($entry in $group.Rows) { $entries[[string]$entry[0]] = $entry } }
foreach ($key in $keys) {
    $full = "SLS.Crack.SP35.$key"; $description = @(Get-SettingInstructionLines $full $entries[$full][3]) -join "`n"
    foreach ($other in $keys) {
        if ($other -ne $key) { Assert-Help "source.$key.noUnrelated.$other" (-not $description.Contains($other)) }
    }
    Assert-Help "source.$key.concise" (($description -split "`n").Count -le 5)
}
$description = @(Get-SettingInstructionLines 'SLS.Crack.SP35.NeighborRatioLimit' $entries['SLS.Crack.SP35.NeighborRatioLimit'][3]) -join "`n"
Assert-Help 'source.neighbor.rebarAndProjection' ($description.Contains('арматурных стержней') -and $description.Contains('проекциями на НЛ'))
$source = [IO.File]::ReadAllText((Join-Path $root 'tools/build_workbook/SettingsCatalog.ps1'))
$methodology = [regex]::Match($source, '(?ms)^function Add-SP35CrackWidthMethodologyGuide\b.*?^}').Value
Assert-Help 'source.noNestedGroupSum' ($methodology -notmatch 'Σ_\{i∈g\}|Σ_g\(')
Assert-Help 'source.noConflictingLocalNAndT' ($methodology -notmatch 'n_g =|t_g =|n_inner =')
Assert-Help 'source.plainDiameterAverage' ($methodology.Contains('d_1 + d_2 + ... + d_p') -and $methodology.Contains('число стержней p'))
Assert-Help 'source.noSP35WildcardBundle' (-not $source.Contains('"SLS.Crack.SP35.*"'))
Assert-Help 'source.sourceLabels' ($source.Contains('Метки источников') -and $source.Contains('[METHOD] - пояснения') -and $source.Contains('[NDM] - принятые'))
$excel = $null; $book = $null; $path = Join-Path $root $WorkbookPath
$hash = (Get-FileHash -LiteralPath $path).Hash
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false; $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($path, 0, $true)
    $settings = $book.Names.Item('rngSystemSettings').RefersToRange
    $guide = $book.Worksheets.Item('Справка'); $data = $guide.UsedRange.Value2
    foreach ($key in $keys) {
        $full = "SLS.Crack.SP35.$key"; $row = 0
        for ($r = 1; $r -le $settings.Rows.Count; $r++) { if ([string]$settings.Cells.Item($r,1).Value2 -eq $full) { $row = $r; break } }
        if ($row -eq 0) { throw "Missing key: $full" }
        $link = [string]$settings.Cells.Item($row,5).Hyperlinks.Item(1).SubAddress
        if ($link -notmatch '!\$?A\$?(\d+)$') { throw "Unexpected Help link: $link" }
        $start = [int]$Matches[1]; $text = ''
        for ($r = $start + 1; $r -le $data.GetLength(0); $r++) {
            if ([string]$data[$r,1] -notmatch '^\d+$') { break }
            $text += [string]$data[$r,2] + "`n"
        }
        Assert-Help "saved.$key.linkTarget" ([string]$data[$start,1] -like "*$full*")
        foreach ($other in $keys) { if ($other -ne $key) { Assert-Help "saved.$key.noUnrelated.$other" (-not $text.Contains($other)) } }
        Assert-Help "saved.$key.textMatchesCatalog" ($text.Trim() -ceq (@(Get-SettingInstructionLines $full $entries[$full][3]) -join "`n").Trim())
    }
    $joined = ($data | ForEach-Object { [string]$_ }) -join "`n"
    # Абзацы справки могут занимать несколько строк объединенных ячеек.
    $plainJoined = [regex]::Replace($joined, '\s+', ' ')
    $formulaText = ($guide.Shapes | ForEach-Object { [string]$_.AlternativeText }) -join "`n"
    Assert-Help 'saved.expandedGroupSum' ($formulaText.Contains('Σβnd = β_1·D_1 + β_2·D_2 + ...') -and -not $formulaText.Contains('Σ_{i∈g}'))
    Assert-Help 'saved.groupDiameterDefinitions' ($joined.Contains('D_1 - сумма диаметров всех стержней первой группы') -and $joined.Contains('β_1 и β_2 - коэффициенты этих же групп'))
    Assert-Help 'saved.numericGroupSumExample' ($formulaText.Contains('0.85·45 + 1.00·16 = 54.25 мм') -and $joined.Contains('деление на 54.25 мм дает R_r в мм'))
    Assert-Help 'saved.centerDefinitions' ($joined.Contains('Индекс «ц» означает центр всей группы') -and $formulaText.Contains('A_1·x_1 + A_2·x_2 + ...'))
    Assert-Help 'saved.localDirectionDefinitions' ($joined.Contains('e_n - единичное направление поперек НЛ') -and $joined.Contains('u_ц - координата центра группы поперек НЛ') -and $formulaText.Contains('u_ц = e_{n,x}·x_ц'))
    Assert-Help 'saved.radiusAndAreaDefinitions' ($joined.Contains('d_ряд - выбранный диаметр опорного ряда') -and $joined.Contains('A_бет,1, A_бет,2, ... - площади наружных частей уже построенного участка'))
    Assert-Help 'saved.internalGroup.concreteBoundary' ($plainJoined.Contains('между рассматриваемой группой и растянутой границей бетона есть другие группы арматуры') -and $plainJoined.Contains('По вышележащим стержням новая граница не проводится'))
    Assert-Help 'saved.internalGroup.membersAndStress' ($plainJoined.Contains('В знаменатель Σβnd входят все стержни целых растянутых групп, чьи центры попали в конечную область') -and $plainJoined.Contains('σ_s наиболее растянутого стержня только рассматриваемой группы, из текущего НДС') -and $plainJoined.Contains('Напряжения других включенных групп в эту σ_s не подставляются'))
    Assert-Help 'saved.internalGroup.separateChecksAndMaximum' ($plainJoined.Contains('Вышележащие группы тоже проверяются поочередно: для каждой строится собственная область и берется ее собственное напряжение') -and $plainJoined.Contains('В итог идет наибольшее раскрытие из всех пригодных проверок'))
    Assert-Help 'saved.sourceLabels' ($joined.Contains('Метки источников') -and $joined.Contains('[METHOD] - пояснения'))
    Assert-Help 'saved.additiveContract' ($joined.Contains('без поиска совпадений и без автоматической очистки'))
    Assert-Help 'saved.colors' ($joined.Contains('ACI 30') -and $joined.Contains('ACI 4') -and $joined.Contains('ACI 31'))
    Assert-Help 'saved.exactDimensions' ($joined.Contains('габаритные размеры импортированного сечения берутся по точному наружному контуру'))
    Assert-Help 'saved.openingsIndependent' ($joined.Contains('Наружный контур и отверстия независимы') -and $joined.Contains('пересечение ее приближенной границы допустимо'))
    Assert-Help 'saved.indivisibleGroupMembership' ($joined.Contains('Группа неделима:') -and $joined.Contains('Если центр снаружи, исключается вся группа') -and $joined.Contains('даже когда отдельный стержень внутри участка'))
    Assert-Help 'saved.fullGroupFormulaData' ($joined.Contains('β и n относятся к полному составу каждой принятой группы') -and $joined.Contains('без коэффициента β') -and $joined.Contains('Центр на общей границе участков включает всю группу') -and -not $joined.Contains('учитывается вошедшая часть:'))
    $full = 'AutoCAD.Common.OpeningContourLayer'; $row = 0
    for ($r = 1; $r -le $settings.Rows.Count; $r++) { if ([string]$settings.Cells.Item($r,1).Value2 -eq $full) { $row = $r; break } }
    if ($row -eq 0) { throw "Missing key: $full" }
    $link = [string]$settings.Cells.Item($row,5).Hyperlinks.Item(1).SubAddress
    if ($link -notmatch '!\$?A\$?(\d+)$') { throw "Unexpected Help link: $link" }
    $start = [int]$Matches[1]; $text = ''
    for ($r = $start + 1; $r -le $data.GetLength(0); $r++) {
        if ([string]$data[$r,1] -notmatch '^\d+$') { break }
        $text += [string]$data[$r,2] + "`n"
    }
    Assert-Help 'saved.opening.linkTarget' ([string]$data[$start,1] -like "*$full*")
    Assert-Help 'saved.opening.textMatchesCatalog' ($text.Trim() -ceq (@(Get-SettingInstructionLines $full $entries[$full][3]) -join "`n").Trim())
    Assert-Help 'saved.opening.commentMatchesCatalog' ([string]$settings.Cells.Item($row,4).Value2 -ceq [string]$entries[$full][3])
} finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
Assert-Help 'saved.readOnly' ((Get-FileHash -LiteralPath $path).Hash -eq $hash)
$lines.Add("TOTAL_HELP: passed=$passed; failed=$failed")
$directory = $PSScriptRoot
if ($ReportDirectory) { $directory = Join-Path $root $ReportDirectory; New-Item -ItemType Directory -Path $directory -Force | Out-Null }
$lines | Set-Content -LiteralPath (Join-Path $directory 'HelpChecks.txt') -Encoding UTF8
Write-Output $lines[$lines.Count-1]
if ($failed -gt 0) { throw 'Help checks failed.' }
