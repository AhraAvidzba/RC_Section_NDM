# Imports only the selected source set into a private copy of the accepted book.
# No migration, calculation, output replacement or user-process cleanup occurs here.
param(
    [string]$Directory = 'docs/regression/Performance/StorageCandidate',
    [switch]$FrozenProduction,
    [string]$SourceWorkbook = 'docs/regression/Performance/Baseline/RC_Section_NDM.xlsm',
    [switch]$TestModuleOnly,
    [string[]]$SourceModules = @(),
    [switch]$RefreshHelp,
    [switch]$AddContourFormatSetting,
    [switch]$RefreshAutoCADWarning,
    [string]$RestoreConfigFormulaWorkbook = '',
    [switch]$PreserveVba,
    [switch]$SkipSmoke,
    [string]$SmokeMacro = 'modTestPerformance.RunPerformanceStorageTests'
)
$ErrorActionPreference = 'Stop'
if ($PreserveVba -and -not $SkipSmoke) { throw 'PreserveVba requires SkipSmoke: workbook macros remain disabled.' }
$SourceModules = @($SourceModules | ForEach-Object { $_ -split ',' } | Where-Object { $_ })
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
$base = [IO.Path]::GetFullPath((Join-Path $root $SourceWorkbook))
$directoryPath = [IO.Path]::GetFullPath((Join-Path $root $Directory))
$allowed = [IO.Path]::GetFullPath((Join-Path $root 'docs/regression/Performance')) + [IO.Path]::DirectorySeparatorChar
if (-not $directoryPath.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase)) { throw 'Fixture must remain inside Performance evidence.' }
if (-not $base.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase)) { throw 'Source fixture is outside Performance evidence.' }
New-Item -ItemType Directory -Path $directoryPath -Force | Out-Null
$target = Join-Path $directoryPath 'RC_Section_NDM.xlsm'
if (Test-Path -LiteralPath $target) { throw 'Fixture exists; use a fresh evidence directory.' }
Copy-Item -LiteralPath $base -Destination $target
$printAreas = @(Get-WorkbookPrintAreas $target)
$baseHash = (Get-FileHash -LiteralPath $base -Algorithm SHA256).Hash
$excel = $null; $book = $null
function Get-ComProperty([object]$Target, [string]$Property) {
    if ($null -eq $Target) { throw "Missing COM target: $Property" }
    try { $value = $Target.GetType().InvokeMember($Property, [Reflection.BindingFlags]::GetProperty, $null, $Target, $null) }
    catch { throw "COM property $Property failed: $($_.Exception.Message)" }
    if ($null -eq $value) { throw "Missing COM property: $Property" }
    return ,$value
}

# Вставка ячеек A:E не должна менять счетчик ROW(A1) скрытых списков.
# Восстанавливаются только известные формулы сборочного списка сочетаний;
# любые другие различия формул требуют отдельного разбирательства.
function Restore-ConfigCombinationListFormulas([object]$Sheet, [object]$Original) {
    $range=$Sheet.Range($Sheet.Cells.Item(1,6),$Sheet.Cells.Item($Original.GetLength(0),$Original.GetLength(1)+5))
    $actual=$range.Formula; $restored=0
    for ($r=1; $r -le $Original.GetLength(0); $r++) {
        for ($c=1; $c -le $Original.GetLength(1); $c++) {
            $before=[string]$Original[$r,$c]; $after=[string]$actual[$r,$c]
            if ($before -ceq $after) { continue }
            if (-not $before.StartsWith('=IF(ROW(A') -or -not $before.Contains('ROWS(rngLoadCombinations)') -or
                -not $after.StartsWith('=IF(ROW(A') -or -not $after.Contains('ROWS(rngLoadCombinations)')) {
                # Новая служебная колонка допускается, но не перезаписывается.
                if (-not $before -and $after -in @('Region','Polyline')) { continue }
                throw "Unexpected Config formula change at row=$r, column=$($c+5)."
            }
            $range.Cells.Item($r,$c).Formula=$before; $restored++
        }
    }
    Write-Output "CONFIG_COMBINATION_LIST_FORMULAS_RESTORED: $restored"
}
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false; $excel.AutomationSecurity = 1
    if ($PreserveVba) { $excel.AutomationSecurity = 3 }
    $books = Get-ComProperty $excel 'Workbooks'
    if ($RestoreConfigFormulaWorkbook) {
        $referencePath=[IO.Path]::GetFullPath((Join-Path $root $RestoreConfigFormulaWorkbook))
        if (-not $referencePath.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)) { throw 'Formula reference is outside Performance evidence.' }
        $referenceHash=(Get-FileHash -LiteralPath $referencePath -Algorithm SHA256).Hash
        $reference=$null
        # Excel не открывает две книги с одинаковым именем в одном instance.
        # Считываем эталон до открытия целевой копии и сразу закрываем его.
        try {
            $reference=$books.Open($referencePath,0,$true)
            $config=$reference.Worksheets.Item('Config'); $used=$config.UsedRange
            $formulaReference=$config.Range($config.Cells.Item(1,6),$config.Cells.Item($used.Row+$used.Rows.Count-1,$used.Column+$used.Columns.Count-1)).Formula
        } finally { if ($reference) { $reference.Close($false) } }
        if ((Get-FileHash -LiteralPath $referencePath -Algorithm SHA256).Hash -cne $referenceHash) { throw 'Formula reference changed.' }
    }
    $book = $books.Open($target)
    $components = Get-ComProperty (Get-ComProperty $book 'VBProject') 'VBComponents'
    $files = @(Get-ChildItem -LiteralPath (Join-Path $root 'src'), (Join-Path $root 'tests') -Recurse -File | Where-Object Extension -in '.cls','.bas')
    foreach ($name in $SourceModules) {
        if (@($files | Where-Object BaseName -eq $name).Count -ne 1) { throw "Source component not found uniquely: $name" }
    }
    if ($PreserveVba) { $files = @() }
    elseif ($FrozenProduction -or $TestModuleOnly) { $files = @($files | Where-Object BaseName -eq 'modTestPerformance') }
    elseif ($SourceModules.Count -gt 0) { $files = @($files | Where-Object {$_.BaseName -eq 'modTestPerformance' -or $_.BaseName -in $SourceModules}) }
    foreach ($file in $files) {
        $component = $null
        for ($i = 1; $i -le $components.Count; $i++) {
            $candidate = $components.Item($i)
            if ([string]$candidate.Name -eq $file.BaseName) { $component = $candidate; break }
        }
        $type = 1; if ($file.Extension -eq '.cls') { $type = 2 }
        if ($null -eq $component) { $component = $components.Add($type); $component.Name = $file.BaseName }
        if ([int]$component.Type -ne $type) { throw "Unexpected component type: $($file.BaseName)" }
        $body = [regex]::Match([IO.File]::ReadAllText($file.FullName, [Text.Encoding]::UTF8), '(?ms)^Option Explicit.*').Value
        if (-not $body) { throw "Missing VBA body: $($file.FullName)" }
        if ($FrozenProduction -and $file.BaseName -eq 'modTestPerformance') {
            $body = $body.Replace('#Const PERFORMANCE_CURRENT = True', '#Const PERFORMANCE_CURRENT = False')
        }
        $code = Get-ComProperty $component 'CodeModule'
        if ($code.CountOfLines -gt 0) { $code.DeleteLines(1, $code.CountOfLines) }
        $code.AddFromString(($body -split "`r?`n") -join "`r`n")
    }
    $project = Get-ComProperty $book 'VBProject'
    $excel.VBE.ActiveVBProject = $project
    $compile = $excel.VBE.CommandBars.FindControl(1, 578)
    if ($null -eq $compile) { throw 'VBA compile command is unavailable.' }
    if ($compile.Enabled) { $compile.Execute() }
    if ($compile.Enabled) {
        $pane = $excel.VBE.ActiveCodePane
        $startLine=0; $startColumn=0; $endLine=0; $endColumn=0
        $pane.GetSelection([ref]$startLine,[ref]$startColumn,[ref]$endLine,[ref]$endColumn)
        throw "VBA compile failed at $($pane.CodeModule.Parent.Name):${startLine}: $($pane.CodeModule.Lines($startLine,1))"
    }
    if ($AddContourFormatSetting) {
        # Дополняем только левый реестр. Остальные таблицы, их списки и
        # пользовательские значения не пересобираются и не сбрасываются.
        $config = $book.Worksheets.Item('Config')
        $table = $book.Names.Item('rngSystemSettings').RefersToRange
        $key = 'AutoCAD.Export.ContourFormat'
        $existing = 0; $after = 0
        $previousSettings = @{}
        for ($r = 2; $r -le $table.Rows.Count; $r++) {
            $name = [string]$table.Cells.Item($r, 1).Value2
            if ($name -and -not $name.StartsWith('[')) {
                $previousSettings[$name] = ConvertTo-Json -InputObject @($table.Cells.Item($r, 2).Formula, $table.Cells.Item($r, 3).Formula, $table.Cells.Item($r, 4).Formula) -Compress
            }
            if ($name -eq $key) { $existing = $r }
            if ($name -eq 'AutoCAD.Export.ContourEnabled') { $after = $r }
        }
        if (-not $after) { throw 'ContourEnabled setting is missing.' }
        if (-not $existing) {
            $used=$config.UsedRange
            $right=$config.Range($config.Cells.Item(1,6),$config.Cells.Item($used.Row+$used.Rows.Count-1,$used.Column+$used.Columns.Count-1))
            $originalRightFormulas=$right.Formula
            $firstRow = $table.Row; $lastRow = $table.Row + $table.Rows.Count - 1
            $firstColumn = $table.Column; $lastColumn = $table.Column + $table.Columns.Count - 1
            $row = $firstRow + $after
            $insert = $config.Range($config.Cells.Item($row, $firstColumn), $config.Cells.Item($row, $lastColumn))
            $insert.Insert(-4121) | Out-Null
            Restore-ConfigCombinationListFormulas $config $originalRightFormulas
            # Excel сдвигает исходный Range вместе с прежними ячейками.
            # Повторно получаем именно новую строку, а не строку под ней.
            $insert = $config.Range($config.Cells.Item($row, $firstColumn), $config.Cells.Item($row, $lastColumn))
            $config.Range($config.Cells.Item($row - 1, $firstColumn), $config.Cells.Item($row - 1, $lastColumn)).Copy($insert)
            $insert.Hyperlinks.Delete(); $insert.Validation.Delete(); $insert.ClearContents()
            $setting = @(Get-SystemSettingsCatalog | ForEach-Object { $_.Rows } | Where-Object { $_[0] -eq $key })
            if ($setting.Count -ne 1) { throw 'ContourFormat is not unique in the catalog.' }
            for ($c = 0; $c -lt 4; $c++) { $insert.Cells.Item(1, $c + 1).Value2 = [string]$setting[0][$c] }
            Set-WorkbookNameByBounds $book 'rngSystemSettings' $config $firstRow $firstColumn ($lastRow + 1) $lastColumn
            # Находим свободную служебную колонку без сброса прежнего реестра.
            Reset-ConfigValidationSources $config
            while ($config.Application.WorksheetFunction.CountA($config.Range($config.Cells.Item(1, $script:ConfigValidationNextColumn), $config.Cells.Item(501, $script:ConfigValidationNextColumn))) -gt 0) {
                $script:ConfigValidationNextColumn++
            }
            $source = Get-ConfigValidationListAddress $config @('Region', 'Polyline')
            $valueCell = $insert.Cells.Item(1, 2)
            $valueCell.Validation.Add(3, 1, 1, $source)
            $valueCell.Validation.IgnoreBlank = $false; $valueCell.Validation.InCellDropdown = $true
            $valueCell.HorizontalAlignment = -4108; $valueCell.VerticalAlignment = -4108
            $config.Application.CutCopyMode = $false
        }
        $table = $book.Names.Item('rngSystemSettings').RefersToRange
        $remaining = @{} + $previousSettings
        for ($r = 2; $r -le $table.Rows.Count; $r++) {
            $name = [string]$table.Cells.Item($r, 1).Value2
            if ($previousSettings.ContainsKey($name)) {
                $actual = ConvertTo-Json -InputObject @($table.Cells.Item($r, 2).Formula, $table.Cells.Item($r, 3).Formula, $table.Cells.Item($r, 4).Formula) -Compress
                if ($actual -cne $previousSettings[$name]) { throw "Existing setting changed while adding ContourFormat: $name" }
                $remaining.Remove($name)
            }
        }
        if ($remaining.Count) { throw "Existing settings lost while adding ContourFormat: $($remaining.Keys -join ', ')" }
    }
    if ($RestoreConfigFormulaWorkbook) {
        Restore-ConfigCombinationListFormulas $book.Worksheets.Item('Config') $formulaReference
        if ((Get-FileHash -LiteralPath $referencePath -Algorithm SHA256).Hash -cne $referenceHash) { throw 'Formula reference changed.' }
    }
    if ($RefreshAutoCADWarning -and -not $RefreshHelp) {
        $guide = $book.Worksheets.Item('Справка')
        $values = $guide.UsedRange.Value2; $warningRows = @()
        for ($r = 1; $r -le $values.GetLength(0); $r++) {
            if ([string]$values[$r,1] -like 'ВНИМАНИЕ: для расчета раскрытия трещин*') {
                $row = $guide.UsedRange.Row + $r - 1
                while ($guide.Cells.Item($row,1).Font.Color -eq 255 -and $guide.Cells.Item($row,1).Font.Bold) {
                    $warningRows += $row; $row++
                }
                break
            }
        }
        $parts = @(Split-GuideParagraphText (Get-AutoCADContourWarningText))
        if ($parts.Count -ne $warningRows.Count) { throw 'Warning row count changed; refresh the full guide instead.' }
        for ($i=0; $i -lt $parts.Count; $i++) { $guide.Cells.Item($warningRows[$i],1).Value2 = [string]$parts[$i] }
    }
    if ($RefreshHelp) {
        Add-SettingsInstructions $book $book.Worksheets.Item('Config') $book.Worksheets.Item(2)
    }
    $book.Save()
    if ($SmokeMacro -and -not $SkipSmoke) {
        $result = [string]$excel.Run("'$($book.Name)'!$SmokeMacro")
        $result | Set-Content -LiteralPath (Join-Path $directoryPath 'Smoke.txt') -Encoding UTF8
        Write-Output $result
        if ($result -match '(?m)^FAIL:' -or $result -notmatch 'failed=0') { throw 'Smoke test failed.' }
    }
    $book.Close($false); $book = $null
    Restore-WorkbookPrintAreas $target $printAreas
    [ordered]@{BaselineSHA256=$baseHash; FrozenProduction=[bool]$FrozenProduction; ImportedComponents=@($files | ForEach-Object { $_.BaseName }); VbaPreserved=[bool]$PreserveVba; HelpRefreshed=[bool]$RefreshHelp; VBACompileCompleted=$true; SourceGitSHA=(& git -C $root rev-parse HEAD); CandidateSHA256=(Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash} |
        ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directoryPath 'Preparation.json') -Encoding UTF8
    if ((Get-FileHash -LiteralPath $base -Algorithm SHA256).Hash -ne $baseHash) { throw 'Baseline changed.' }
}
finally {
    if ($null -ne $book) { try {$book.Close($false)} catch {Write-Warning $_}; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($null -ne $excel) { try {$excel.Quit()} catch {Write-Warning $_}; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
