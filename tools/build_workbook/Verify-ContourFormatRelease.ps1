# Проверяет узкий выпуск формата контуров и сохраняет доказательства.
# Публикация разрешена только после направленных проверок и визуальной приемки.
param([string]$Directory = 'docs/regression/Performance/RegionContourExport_2026-10-08',
      [string]$CandidateDirectory = 'Candidate',
      [string]$NativeDirectory = 'Native',
      [string]$UnitDirectory = 'Unit',
      [switch]$Publish)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$directoryPath = [IO.Path]::GetFullPath((Join-Path $root $Directory))
$allowed = [IO.Path]::GetFullPath((Join-Path $root 'docs/regression/Performance')) + [IO.Path]::DirectorySeparatorChar
if (-not $directoryPath.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase)) { throw 'Evidence must remain inside Performance.' }
$baseline = Join-Path $directoryPath 'Baseline/RC_Section_NDM.xlsm'
$candidate = Join-Path $directoryPath "$CandidateDirectory/RC_Section_NDM.xlsm"
$export = Join-Path $directoryPath "$CandidateDirectory/VBA_All_Code.txt"
$outputBook = Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$outputExport = Join-Path $root 'workbook/output/VBA_All_Code.txt'
$outputReport = Join-Path $root 'workbook/output/RC_Section_NDM_execution_report.txt'

# Отдельные файлы проверки должны принадлежать именно этой версии книги.
function Assert-Hash([string]$Path, [string]$Expected) {
    if ((Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash -cne $Expected) { throw "Artifact changed: $Path" }
}
function Text-Hash([string]$Text) {
    $sha = [Security.Cryptography.SHA256]::Create()
    try { ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Text)))).Replace('-', '') }
    finally { $sha.Dispose() }
}

# После правки только справки повтор CAD не нужен, если все рабочие
# компоненты и сам CAD-тест поблочно совпадают с успешно проверенной книгой.
function Assert-NativeCodeUnchanged([string]$NativeExport, [string]$CurrentExport) {
    $tokens=$null; $errors=$null
    $ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'Compare-PerformanceVba.ps1'),[ref]$tokens,[ref]$errors)
    $node=$ast.Find({param($item) $item -is [Management.Automation.Language.FunctionDefinitionAst] -and $item.Name -eq 'Normalize-Code'},$true)
    if ($null -eq $node) { throw 'VBA normalizer is missing.' }
    . ([scriptblock]::Create($node.Extent.Text))
    $pattern='(?ms)^={100}\r?\nCOMPONENT: ([^\r\n]+)\r?\nTYPE: (\d+)\r?\nLINES: (\d+)\r?\n={100}\r?\n(.*?)(?=^={100}\r?\nCOMPONENT: |\z)'
    $snapshots=@()
    foreach ($path in @($NativeExport,$CurrentExport)) {
        $components=@{}
        foreach ($match in [regex]::Matches([IO.File]::ReadAllText($path,[Text.Encoding]::UTF8),$pattern)) {
            $body=$match.Groups[4].Value
            if (-not [string]::IsNullOrWhiteSpace($body)) { $body=Normalize-Code $body } else { $body='' }
            $components[$match.Groups[1].Value]=@{Type=[int]$match.Groups[2].Value; Body=$body}
        }
        $snapshots += ,$components
    }
    if ($snapshots[0].Count -ne 128 -or $snapshots[1].Count -ne 128) { throw 'Native comparison is incomplete.' }
    $changes=@(); $same=@()
    foreach ($name in $snapshots[0].Keys) {
        $a=$snapshots[0][$name]; $b=$snapshots[1][$name]
        if ($null -eq $b -or $a.Type -ne $b.Type) { throw "Native component type changed: $name" }
        if ($a.Body -cne $b.Body) { $changes += $name } else { $same += $name }
    }
    if (@($changes | Where-Object { $_ -ne 'modTestWorkbookInterface' }).Count) { throw 'CAD code changed after the native test.' }
    [ordered]@{Passed=$true; ComparedComponents=128; MatchedComponents=$same.Count;
        ChangedComponents=$changes; AllowedChange='modTestWorkbookInterface (help assertions only)';
        NativeExportSHA256=(Get-FileHash -LiteralPath $NativeExport -Algorithm SHA256).Hash;
        CurrentExportSHA256=(Get-FileHash -LiteralPath $CurrentExport -Algorithm SHA256).Hash}
}

# Добавленная строка реестра не меняет прежние настройки. Остальные ячейки
# Config сравниваются по исходным адресам, а ссылка справки - по смысловой цели.
function Book-Fingerprint([object]$Book, [int]$ConfigLastColumn, [int]$ConfigLastRow, [int]$NewSourceColumn) {
    $sheets = @()
    foreach ($sheet in $Book.Worksheets) {
        if ([string]$sheet.Name -in @('Config', 'Справка')) { continue }
        $range = $sheet.UsedRange
        $sheets += [ordered]@{Name = [string]$sheet.Name; Address = $range.Address();
            FormulaHash = (Text-Hash ($range.Formula | ConvertTo-Json -Depth 6 -Compress));
            ValueHash = (Text-Hash ($range.Value2 | ConvertTo-Json -Depth 6 -Compress))}
    }
    $config = $Book.Worksheets.Item('Config')
    $right = $config.Range($config.Cells.Item(1, 6), $config.Cells.Item($ConfigLastRow, $ConfigLastColumn))
    $settings = $Book.Names.Item('rngSystemSettings').RefersToRange
    $rows = @(); $values = $settings.Formula
    for ($r = 1; $r -le $settings.Rows.Count; $r++) {
        if ([string]$values[$r, 1] -eq 'AutoCAD.Export.ContourFormat') { continue }
        $rows += ,@($values[$r, 1], $values[$r, 2], $values[$r, 3], $values[$r, 4], $values[$r, 5])
    }
    $names = @()
    foreach ($name in $Book.Names) {
        if ([string]$name.Name -eq 'rngSystemSettings') { continue }
        $names += ([string]$name.Name + '|' + [string]$name.RefersTo)
    }
    $widths = @(); for ($c = 1; $c -le 260; $c++) { $widths += [double]$Book.Worksheets.Item('Results').Columns.Item($c).ColumnWidth }
    $targets = @(); $help = $Book.Worksheets.Item('Справка')
    foreach ($link in $config.Hyperlinks) {
        $address = [string]$link.Range.Address()
        $key = $address
        if ($link.Range.Column -eq $settings.Column + 4 -and $link.Range.Row -ge $settings.Row -and $link.Range.Row -lt $settings.Row + $settings.Rows.Count) {
            $key = [string]$config.Cells.Item($link.Range.Row, $settings.Column).Value2
        }
        if ($key -eq 'AutoCAD.Export.ContourFormat') { continue }
        $subAddress = [string]$link.SubAddress
        if ($subAddress -notmatch '!\$?([A-Z]+)\$?(\d+)$') { throw "Unexpected help link: $subAddress" }
        $target = [string]$help.Range($Matches[1] + $Matches[2]).Value2
        $targets += ($key + '|' + [string]$link.TextToDisplay + '|' + $target)
    }
    $rightFormula=$right.Formula; $rightValues=$right.Value2
    if ($NewSourceColumn -le $ConfigLastColumn) {
        $index=$NewSourceColumn - 5
        $rightFormula[1,$index]=$null; $rightFormula[2,$index]=$null
        $rightValues[1,$index]=$null; $rightValues[2,$index]=$null
    }
    [ordered]@{Sheets=$sheets; ConfigOtherFormula=(Text-Hash ($rightFormula | ConvertTo-Json -Depth 6 -Compress));
        ConfigOtherValues=(Text-Hash ($rightValues | ConvertTo-Json -Depth 6 -Compress));
        Settings=(Text-Hash ($rows | ConvertTo-Json -Depth 6 -Compress)); Names=(Text-Hash ((@($names | Sort-Object)) -join "`n"));
        ResultsWidths=(Text-Hash ($widths | ConvertTo-Json -Compress)); HelpTargets=(Text-Hash ((@($targets | Sort-Object)) -join "`n"));
        OriginalHelpLinkCount=$targets.Count}
}

$baselineHash = (Get-FileHash -LiteralPath $baseline -Algorithm SHA256).Hash
$candidateHash = (Get-FileHash -LiteralPath $candidate -Algorithm SHA256).Hash
Assert-Hash $outputBook $baselineHash
$preparation = Get-Content -LiteralPath (Join-Path $directoryPath "$CandidateDirectory/Preparation.json") -Raw | ConvertFrom-Json
if (-not $preparation.VBACompileCompleted) { throw 'Preparation is incomplete.' }
Assert-Hash $candidate $preparation.CandidateSHA256
$equality = Get-Content -LiteralPath (Join-Path $directoryPath "$CandidateDirectory/SourceEquality.json") -Raw | ConvertFrom-Json
if (-not $equality.AllPassed) { throw 'Actual VBA/source equality failed.' }
$structure = Get-Content -LiteralPath (Join-Path $directoryPath "$CandidateDirectory/Structure.json") -Raw | ConvertFrom-Json
if (@($structure | Where-Object { -not $_.Passed }).Count) { throw 'Workbook structure failed.' }
foreach ($case in @(@("$UnitDirectory/Contours.txt", 'TOTAL_AUTOCAD_CONTOURS'), @("$UnitDirectory/Presentation.txt", 'TOTAL_CONFIG_PRESENTATION'))) {
    $text = Get-Content -LiteralPath (Join-Path $directoryPath $case[0]) -Raw
    if ($text -match '(?m)^FAIL:' -or $text -notmatch ($case[1] + ': passed=[1-9][0-9]*; failed=0') -or
        -not $text.Contains($candidateHash)) { throw "Directed gate failed: $($case[0])" }
}
$native = Get-Content -LiteralPath (Join-Path $directoryPath "$NativeDirectory/NativeCAD.txt") -Raw
$identity = Get-Content -LiteralPath (Join-Path $directoryPath "$NativeDirectory/NativeCADIdentity.json") -Raw | ConvertFrom-Json
if ($native -match '(?m)^FAIL:' -or $native -notmatch 'TOTAL_REAL_AUTOCAD_CONTOURS: passed=[1-9][0-9]*; failed=0' -or
    $identity.UserDocumentsUsed -or $identity.Executable -cne 'C:\Program Files\Autodesk\AutoCAD 2023\acad.exe') { throw 'Native CAD gate failed.' }
foreach ($format in @('Region', 'Polyline')) {
    foreach ($mode in @(0, 1)) {
        foreach ($suffix in @('count', 'area', 'loops')) {
            if ($native -notmatch ('(?m)^OK: ' + [regex]::Escape("native.format.$format.$mode.$suffix") + '\b')) { throw 'Native format coverage is incomplete.' }
        }
    }
}
$nativeBook=Join-Path $directoryPath "$NativeDirectory/RC_Section_NDM.xlsm"
$nativeHash=(Get-FileHash -LiteralPath $nativeBook -Algorithm SHA256).Hash
$nativeReused=$false
if ($nativeHash -cne $candidateHash) {
    $nativeProof=Assert-NativeCodeUnchanged (Join-Path $directoryPath "$NativeDirectory/VBA_All_Code.txt") $export
    $nativeProof | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directoryPath 'NativeCodeReuse.json') -Encoding UTF8
    $nativeReused=$true
}
$reportHash = (Get-FileHash -LiteralPath $outputReport -Algorithm SHA256).Hash
$oldExportHash = (Get-FileHash -LiteralPath $outputExport -Algorithm SHA256).Hash
$exportHash = (Get-FileHash -LiteralPath $export -Algorithm SHA256).Hash
$excel = $null; $book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false; $excel.AutomationSecurity=3
    $book = $excel.Workbooks.Open($candidate, 0, $true)
    $table=$book.Names.Item('rngSystemSettings').RefersToRange
    $formatRow=0
    for ($r=2; $r -le $table.Rows.Count; $r++) { if ([string]$table.Cells.Item($r,1).Value2 -eq 'AutoCAD.Export.ContourFormat') { $formatRow=$r } }
    if (-not $formatRow) { throw 'ContourFormat is missing.' }
    $source=$table.Worksheet.Range(([string]$table.Cells.Item($formatRow,2).Validation.Formula1).Substring(1))
    $sourceColumn=$source.Column
    $book.Close($false); $book=$null
    $book = $excel.Workbooks.Open($baseline, 0, $true)
    if ($excel.WorksheetFunction.CountA($book.Worksheets.Item('Config').Range($book.Worksheets.Item('Config').Cells.Item(1,$sourceColumn),$book.Worksheets.Item('Config').Cells.Item(501,$sourceColumn))) -ne 0) { throw 'New validation source overwrote existing content.' }
    $range = $book.Worksheets.Item('Config').UsedRange
    $lastColumn = $range.Column + $range.Columns.Count - 1; $lastRow = $range.Row + $range.Rows.Count - 1
    $before = Book-Fingerprint $book $lastColumn $lastRow $sourceColumn
    $oldSettings = $book.Names.Item('rngSystemSettings').RefersToRange
    $oldTop = $oldSettings.Row; $oldLeft = $oldSettings.Column; $oldRows = $oldSettings.Rows.Count
    $book.Close($false); $book=$null
    $book = $excel.Workbooks.Open($candidate, 0, $true)
    $after = Book-Fingerprint $book $lastColumn $lastRow $sourceColumn
    $before | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $directoryPath 'BeforeFingerprint.json') -Encoding UTF8
    $after | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $directoryPath 'AfterFingerprint.json') -Encoding UTF8
    if (($before | ConvertTo-Json -Depth 8 -Compress) -cne ($after | ConvertTo-Json -Depth 8 -Compress)) {
        $differences=@($before.Keys | Where-Object { ($before[$_] | ConvertTo-Json -Depth 8 -Compress) -cne ($after[$_] | ConvertTo-Json -Depth 8 -Compress) })
        throw ('Unexpected user-data change: ' + ($differences -join ', '))
    }
    $settings = $book.Names.Item('rngSystemSettings').RefersToRange
    if ($settings.Row -ne $oldTop -or $settings.Column -ne $oldLeft -or $settings.Rows.Count -ne $oldRows + 1) { throw 'Unexpected settings range shift.' }
    $settingRow = 0
    for ($r = 2; $r -le $settings.Rows.Count; $r++) { if ([string]$settings.Cells.Item($r, 1).Value2 -eq 'AutoCAD.Export.ContourFormat') { $settingRow=$r } }
    if (-not $settingRow) { throw 'ContourFormat setting is missing.' }
    $cell = $settings.Cells.Item($settingRow, 2)
    if ([string]$cell.Value2 -cne 'Region' -or $cell.Validation.Type -ne 3 -or -not $cell.Validation.InCellDropdown) { throw 'ContourFormat dropdown/default is wrong.' }
    $options = @($cell.Worksheet.Range(([string]$cell.Validation.Formula1).Substring(1)).Value2 | ForEach-Object { $_ })
    if (($options -join '|') -cne 'Region|Polyline') { throw 'ContourFormat options are wrong.' }
    $link = $settings.Cells.Item($settingRow, 5)
    if ($link.Hyperlinks.Count -ne 1 -or $link.Borders.Item(10).LineStyle -ne -4115) { throw 'New help link or right border is missing.' }
    $help = $book.Worksheets.Item('Справка')
    $newTarget=[string]$link.Hyperlinks.Item(1).SubAddress
    if ($newTarget -notmatch '!\$?([A-Z]+)\$?(\d+)$') { throw 'New help link target is invalid.' }
    if ([string]$help.Range($Matches[1]+$Matches[2]).Value2 -cne 'AutoCAD.Export.ContourFormat') { throw 'New setting points to another help block.' }
    if ([string]$help.Cells.Item(1, 1).Value2 -cne 'Подготовка геометрии в AutoCAD и импорт') { throw 'Import guide is not first.' }
    $text = ($help.UsedRange.Value2 | ForEach-Object { [string]$_ }) -join "`n"
    foreach ($phrase in @('автоматически получает наружную границу', 'SUBTRACT', 'AutoCAD.Export.ContourFormat', 'каждое вырезанное отверстие')) {
        if ($text -notmatch [regex]::Escape($phrase)) { throw "Help explanation is missing: $phrase" }
    }
    $guideValues=$help.UsedRange.Value2; $lastGuideRow=0
    for ($r=1; $r -le $guideValues.GetLength(0); $r++) {
        if ([string]$guideValues[$r,1] -ceq 'Поиск НДС текущего сочетания') { $lastGuideRow=$help.UsedRange.Row+$r-2; break }
    }
    if ($lastGuideRow -le 1) { throw 'End of geometry guide is missing.' }
    $help.Outline.ShowLevels(8)
    $help.PageSetup.PrintArea=$help.Range($help.Cells.Item(1,1),$help.Cells.Item($lastGuideRow,6)).Address()
    $help.PageSetup.Orientation=2; $help.PageSetup.PaperSize=8
    $help.PageSetup.Zoom=$false; $help.PageSetup.FitToPagesWide=1; $help.PageSetup.FitToPagesTall=2
    $help.ExportAsFixedFormat(0, (Join-Path $directoryPath 'GeometryGuide.pdf'), 0, $true, $false)
    $config=$book.Worksheets.Item('Config'); $row=$cell.Row
    $config.PageSetup.PrintArea=$config.Range($config.Cells.Item($row - 2, 1), $config.Cells.Item($row + 3, 5)).Address()
    $config.PageSetup.Orientation=2; $config.PageSetup.PaperSize=8; $config.PageSetup.Zoom=$false
    $config.PageSetup.FitToPagesWide=1; $config.PageSetup.FitToPagesTall=1
    $config.ExportAsFixedFormat(0, (Join-Path $directoryPath 'ExportSettings.pdf'), 0, $true, $false)
    $book.Close($false); $book=$null
} finally {
    if ($book) { try {$book.Close($false)} catch {} }
    if ($excel) { try {$excel.Quit()} catch {}; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
Assert-Hash $candidate $candidateHash
Assert-Hash $baseline $baselineHash
Assert-Hash $outputBook $baselineHash
Assert-Hash $outputReport $reportHash
Assert-Hash $outputExport $oldExportHash
$verification=[ordered]@{Passed=$true; WorkbookSHA256=$candidateHash; ExportSHA256=$exportHash; BaselineWorkbookSHA256=$baselineHash;
    UserReportSHA256=$reportHash; UserReportPreserved=$true; UserDataNamesWidthsPreserved=$true; Fingerprint=$after;
    NewSetting='AutoCAD.Export.ContourFormat'; Default='Region'; NativeBothFormatsPassed=$true;
    NativeInputWorkbookSHA256=$nativeHash; NativeReusedAfterHelpOnlyChange=$nativeReused; FullOnOffRepeated=$false}
$verification | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $directoryPath 'Verification.json') -Encoding UTF8
if ($Publish) {
    $visual = Get-Content -LiteralPath (Join-Path $directoryPath 'VisualAcceptance.json') -Raw | ConvertFrom-Json
    if (-not $visual.Passed -or $visual.WorkbookSHA256 -cne $candidateHash) { throw 'Visual acceptance is missing.' }
    $guard=[IO.File]::Open($outputBook,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::None); $guard.Dispose()
    Copy-Item -LiteralPath $candidate -Destination $outputBook
    Copy-Item -LiteralPath $export -Destination $outputExport
    Assert-Hash $outputBook $candidateHash; Assert-Hash $outputExport $exportHash; Assert-Hash $outputReport $reportHash
    $verification | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $directoryPath 'Publication.json') -Encoding UTF8
}
Write-Output ('CONTOUR_FORMAT_RELEASE_VERIFIED: published=' + [bool]$Publish + '; SHA256=' + $candidateHash)
