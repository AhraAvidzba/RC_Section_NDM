# Собирает возможность полного снимка в отдельной книге, сохраняя пользовательский ввод.
param([string]$SourceWorkbook='workbook/output/RC_Section_NDM.xlsm',
      [string]$EvidenceDirectory='docs/regression/CompleteSnapshot_2026-10-24/Release01')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
$script:WorkbookLayout=Read-WorkbookLayout
$directory=[IO.Path]::GetFullPath((Join-Path $root $EvidenceDirectory))
$allowed=(Join-Path $root 'docs/regression')+[IO.Path]::DirectorySeparatorChar
if(-not $directory.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)){throw 'Кандидат должен находиться в docs/regression.'}
if(Test-Path -LiteralPath $directory){throw 'Каталог доказательств уже существует.'}
[void][IO.Directory]::CreateDirectory($directory)
$source=(Resolve-Path -LiteralPath (Join-Path $root $SourceWorkbook)).Path
$sourceHash=(Get-FileHash -LiteralPath $source).Hash
$path=Join-Path $directory 'RC_Section_NDM.xlsm'
Copy-Item -LiteralPath $source -Destination $path
$printAreas=@(Get-WorkbookPrintAreas $path)
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class SnapshotBuildWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
$app=$null;$book=$null;$owned=0;$ticks='';$existing=@(Get-Process EXCEL -ErrorAction SilentlyContinue|ForEach-Object Id)

# Проверяет принадлежность Excel перед использованием и закрытием.
function Test-OwnedExcel {
    if(-not $owned -or $owned -in $existing){return $false}
    $process=Get-Process -Id $owned -ErrorAction SilentlyContinue
    return $process -and $process.ProcessName -eq 'EXCEL' -and $process.StartTime.ToUniversalTime().Ticks.ToString() -ceq $ticks
}

# Находит ключ, не привязывая обновление Config и Справки к фиксированным строкам.
function Find-KeyRow($Sheet,[string]$Key) {
    $last=[int]$Sheet.UsedRange.Row+[int]$Sheet.UsedRange.Rows.Count-1
    $values=$Sheet.Range('A1:A'+$last).Value2
    for($r=1;$r -le $last;$r++){if([string]$values[$r,1] -ceq $Key){return $r}}
    throw ('Отсутствует ключ: '+$Sheet.Name+'! '+$Key)
}

try {
    $app=New-Object -ComObject Excel.Application
    [uint32]$id=0;[void][SnapshotBuildWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id)
    $owned=[int]$id;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks.ToString()
    if(-not(Test-OwnedExcel)){throw 'Не подтверждена принадлежность нового Excel.'}
    $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=1
    [ordered]@{PID=$owned;StartTicks=$ticks;Source=$source;SourceSHA256=$sourceHash}|ConvertTo-Json|
        Set-Content -LiteralPath (Join-Path $directory 'Ownership.json') -Encoding UTF8
    Write-Output ('OWNED_EXCEL: id='+$owned+'; startTicks='+$ticks)
    $book=$app.Workbooks.Open($path,0,$false)
    if($book.ReadOnly){throw 'Кандидат открыт только для чтения.'}
    $config=$book.Worksheets.Item('Config');$help=$book.Worksheets.Item('Справка');$results=$book.Worksheets.Item('Results')
    $oldNames=@($book.Names|ForEach-Object Name)
    $definitions=@((Get-SystemSettingsCatalog|Where-Object Name -eq 'GeneralSettings').Rows|Where-Object {$_[0] -like 'Results.Snapshot*'})
    $row=Find-KeyRow $config 'General.WorstCombinationCriterion'
    $config.Range('A'+$row+':E'+($row+1)).Insert(-4121)|Out-Null
    for($i=0;$i -lt 2;$i++) {
        $r=$row+$i;$definition=$definitions[$i]
        $config.Range('A'+($row+2)+':E'+($row+2)).Copy($config.Range('A'+$r+':E'+$r))
        $config.Cells.Item($r,1).Value2=[string]$definition[0]
        $config.Cells.Item($r,2).Value2=[string]$definition[1]
        $config.Cells.Item($r,3).Value2=[string]$definition[2]
        $config.Cells.Item($r,4).Value2=[string]$definition[3]
        $config.Cells.Item($r,4).HorizontalAlignment=-4131;$config.Cells.Item($r,4).WrapText=$false
        $cell=$config.Cells.Item($r,2);$cell.HorizontalAlignment=-4108;$cell.Interior.Color=15135487
        $cell.Validation.Delete()
        if($i -eq 0){$cell.Validation.Add(3,1,1,(Get-ConfigValidationListAddress $config @('Excel','File')));$cell.Validation.IgnoreBlank=$false;$cell.Validation.ShowError=$true;$cell.Validation.InCellDropdown=$true}
        $config.Cells.Item($r,5).Hyperlinks.Delete()
    }
    $help.Outline.ShowLevels(8)
    foreach($definition in $definitions) {
        $key=[string]$definition[0];$helpRow=Find-KeyRow $help 'General.WorstCombinationCriterion'
        $instruction=@(Get-SettingsInstructionCatalog|Where-Object Key -ceq $key)[0]
        $lines=@($instruction.Lines|Where-Object {-not[string]::IsNullOrWhiteSpace([string]$_)})
        $inserted=$lines.Count+2
        $links=@(foreach($sheet in $book.Worksheets){foreach($link in $sheet.Hyperlinks){
            $target=[string]$link.SubAddress
            if($target -match '^''?Справка''?!(\$?)([A-Z]+)(\$?)(\d+)$' -and [int]$Matches[4] -ge $helpRow){
                [pscustomobject]@{Link=$link;Target=("'Справка'!"+$Matches[1]+$Matches[2]+$Matches[3]+([int]$Matches[4]+$inserted))}
            }
        }})
        $help.Range([string]$helpRow+':'+($helpRow+$inserted-1)).Insert(-4121)|Out-Null
        foreach($record in $links){$record.Link.SubAddress=$record.Target}
        $template=$helpRow+$inserted
        $help.Range('A'+$template+':F'+$template).Copy($help.Range('A'+$helpRow+':F'+$helpRow))
        $help.Cells.Item($helpRow,1).Value2=$key
        for($i=0;$i -lt $lines.Count;$i++) {
            $r=$helpRow+$i+1
            $help.Range('A'+($template+1)+':F'+($template+1)).Copy($help.Range('A'+$r+':F'+$r))
            Set-InstructionParagraphNumber $help.Cells.Item($r,1) ($i+1)
            $help.Cells.Item($r,2).Value2=[string]$lines[$i]
            $help.Rows.Item($r).AutoFit()|Out-Null
        }
        $configRow=Find-KeyRow $config $key
        Add-InstructionHyperlink $config.Cells.Item($configRow,5) 'Справка' $helpRow
    }
    # Создание Hyperlink может убрать левую и правую грань ячейки Подробнее.
    # Восстанавливаем каждую рамку новых строк после создания всех ссылок.
    foreach($definition in $definitions){
        $r=Find-KeyRow $config ([string]$definition[0])
        $referenceRow=Find-KeyRow $config 'General.WorstCombinationCriterion'
        for($column=1;$column -le 5;$column++){
            $cell=$config.Cells.Item($r,$column);$reference=$config.Cells.Item($referenceRow,$column)
            $cell.HorizontalAlignment=$reference.HorizontalAlignment
            $cell.VerticalAlignment=$reference.VerticalAlignment
            $cell.WrapText=$reference.WrapText
            $cell.NumberFormat=$reference.NumberFormat
        }
        $band=$config.Range('A'+$r+':E'+$r)
        $band.Borders.Weight=2;$band.Borders.Color=0;$band.Borders.LineStyle=-4115
        $right=$config.Cells.Item($r,5).Borders.Item(10)
        $right.Weight=2;$right.Color=0;$right.LineStyle=-4115
    }
    # Все прежние таблицы сначала переносятся на временный лист, чтобы при
    # перестановке перекрывающихся областей не потерять значения или стили.
    $layout=@(
        @{Name='rngNDMSectionGeometry';Column=1;Width=16},
        @{Name='rngNDMSectionProperties';Column=19;Width=6},
        @{Name='rngNDMMaterialDiagrams';Column=27;Width=11},
        @{Name='rngNDMElementResults';Column=40;Width=9},
        @{Name='rngNDMSectionContours';Column=55;Width=17},
        @{Name='rngNDMSectionAnnotations';Column=74;Width=24}
    )
    $temp=$book.Worksheets.Add();$temp.Name='NDM_SnapshotBuild_Temp'
    $targetRow=[int]$book.Names.Item('rngNDMSectionGeometry').RefersToRange.Row
    foreach($block in $layout) {
        $anchor=$book.Names.Item($block.Name).RefersToRange
        $last=[int]$results.Cells.Item($results.Rows.Count,$anchor.Column).End(-4162).Row
        $block.Rows=[Math]::Max(1,$last-$anchor.Row+1)
        $block.OldColumn=[int]$anchor.Column
        $anchor.Offset(-1,0).Resize($block.Rows+1,$block.Width).Copy($temp.Cells.Item(1,$block.Column))
    }
    foreach($block in $layout){$results.Cells.Item($targetRow-2,$block.OldColumn).Resize($block.Rows+2,$block.Width).UnMerge();$results.Cells.Item($targetRow-2,$block.OldColumn).Resize($block.Rows+2,$block.Width).Clear()}
    foreach($block in $layout) {
        $temp.Cells.Item(1,$block.Column).Resize($block.Rows+1,$block.Width).Copy($results.Cells.Item($targetRow-1,$block.Column))
        $book.Names.Item($block.Name).RefersTo="='Results'!"+$results.Cells.Item($targetRow,$block.Column).Address()
    }
    $temp.Delete()
    [void]$book.Names.Add('rngNDMLoadCaseResults',"='Results'!"+$results.Cells.Item($targetRow,51).Address())
    $calculation=$book.Worksheets.Item('Расчет')
    $calculation.Shapes.Item('btnUpdateSectionPlot').TextFrame.Characters().Text='Обновить схему и результаты'
    # Короткая подпись помещается в прежний размер кнопки без перестройки Config.
    $calculation.Shapes.Item('btnUpdateSectionPlot').TextFrame.Characters().Font.Size=8
    $helperColumn=Get-WorkbookLayoutHelperColumn 'LoadCases'
    $helperRows=$book.Names.Item('rngLoadCombinations').RefersToRange.Rows.Count-1
    $config.Range($config.Cells.Item(2,$helperColumn),$config.Cells.Item($helperRows+1,$helperColumn)).Formula='=IF(ROW()-1>ROWS(rngLoadCombinations)-1,"",INDEX(rngLoadCombinations,ROW(),1))'
    # Полный импорт исходников гарантирует единую ревизию всех ускорений.
    $files=@(Get-ChildItem -LiteralPath (Join-Path $root 'src'),(Join-Path $root 'tests') -Recurse -File|Where-Object Extension -in '.bas','.cls')
    foreach($file in $files) {
        try{$component=$book.VBProject.VBComponents.Item($file.BaseName)}catch{$component=$book.VBProject.VBComponents.Add($(if($file.Extension -eq '.cls'){2}else{1}));$component.Name=$file.BaseName}
        $body=[regex]::Match([IO.File]::ReadAllText($file.FullName,[Text.Encoding]::UTF8),'(?ms)^Option Explicit.*').Value
        if(-not $body -and $file.BaseName -eq 'modSectionRecipeApiGuide'){$body=([IO.File]::ReadAllLines($file.FullName,[Text.Encoding]::UTF8)|Where-Object {$_ -notmatch '^Attribute VB_Name'}) -join "`r`n"}
        if(-not $body){throw ('Не найден код модуля: '+$file.Name)}
        $module=$component.CodeModule
        if($module.CountOfLines){$module.DeleteLines(1,$module.CountOfLines)}
        $module.AddFromString(($body -split '\r?\n') -join "`r`n")
    }
    $events=$book.VBProject.VBComponents.Item($book.CodeName).CodeModule
    if($events.CountOfLines){$events.DeleteLines(1,$events.CountOfLines)}
    $events.AddFromString(([IO.File]::ReadAllLines((Join-Path $PSScriptRoot 'WorkbookHelpEvents.vba'),[Text.Encoding]::UTF8)) -join "`r`n")
    # Одноразовая миграция старого сохранённого расчёта существует только
    # в сборщике. В рабочем VBA и опубликованной книге этого модуля нет.
    $migration=$book.VBProject.VBComponents.Add(1);$migration.Name='modTestSnapshotBuild'
    $migration.CodeModule.AddFromString(@'
Option Explicit
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Переносит уже сохранённые значения в новый контракт без нового решения.
Public Function TestMigrateSavedSnapshot() As Boolean
    Dim s As CCalculationSnapshot: Set s = New CCalculationSnapshot
    Dim key As Variant, values As Variant
    values = ReadPhysicalResultTable(ThisWorkbook, "rngNDMElementResults")
    If UBound(values, 1) < 2 Then Exit Function
    s.RunID = CDbl(values(2, 1))
    Dim elapsed As Variant: elapsed = ThisWorkbook.Names("rngBatchSummary").RefersToRange.Offset(3, 0).Value2
    If IsNumeric(elapsed) Then s.ElapsedSeconds = CDbl(elapsed)
    For Each key In s.BaseKeys()
        values = ReadPhysicalResultTable(ThisWorkbook, CStr(key))
        If CStr(key) = "rngNDMSectionProperties" Or CStr(key) = "rngNDMSectionAnnotations" Then
            s.AddSeparatedTable CStr(key), values
        Else
            s.AddTable CStr(key), values
        End If
    Next key
    Dim writer As CBatchResultWriter: Set writer = New CBatchResultWriter
    Dim rows As Long: rows = writer.RequiredDataRowsForWorkbook(ThisWorkbook, 0)
    s.AddTable "Summary", ThisWorkbook.Names("rngBatchSummary").RefersToRange.Resize(writer.RequiredSummaryRowsForWorkbook(ThisWorkbook, 0), 26).Value2
    s.AddTable "Strength.Headers", ThisWorkbook.Names("rngStrengthSummaryAnchor").RefersToRange.Offset(-4).Resize(4, 49).Value2
    s.AddTable "Strength.Data", ThisWorkbook.Names("rngStrengthSummaryAnchor").RefersToRange.Resize(rows, 49).Value2
    s.AddTable "Cracks.Headers", ThisWorkbook.Names("rngCrackSummaryAnchor").RefersToRange.Offset(-4).Resize(4, 72).Value2
    s.AddTable "Cracks.Data", ThisWorkbook.Names("rngCrackSummaryAnchor").RefersToRange.Resize(rows, 72).Value2
    s.AddTable "Stability.Labels", ThisWorkbook.Names("rngStabilitySummaryAnchor").RefersToRange.Offset(-1).Resize(1, 84).Value2
    s.AddTable "Stability.Data", ThisWorkbook.Names("rngStabilitySummaryAnchor").RefersToRange.Resize(rows, 84).Value2
    s.AddTable "Calculation.Profiles", ThisWorkbook.Names("rngCalculationProfiles").RefersToRange.Value2
    s.Validate
    Dim store As CCalculationSnapshotStore: Set store = New CCalculationSnapshotStore
    store.SaveSheetExtra ThisWorkbook, s
    Dim base As CNDMResultsWriter: Set base = New CNDMResultsWriter
    base.RestoreSnapshotBase ThisWorkbook, s
    TestMigrateSavedSnapshot = True
End Function
'@)
    $app.VBE.ActiveVBProject=$book.VBProject
    $compile=$app.VBE.CommandBars.FindControl(1,578)
    if($null -eq $compile){throw 'Не найдена команда компиляции VBA.'}
    if($compile.Enabled){$compile.Execute()}
    if($compile.Enabled){
        $pane=$app.VBE.ActiveCodePane
        [int]$line=0;[int]$column=0;[int]$lastLine=0;[int]$lastColumn=0
        $pane.GetSelection([ref]$line,[ref]$column,[ref]$lastLine,[ref]$lastColumn)
        $diagnostic=[ordered]@{Module=$pane.CodeModule.Parent.Name;Line=$line;Column=$column;Code=$pane.CodeModule.Lines([Math]::Max(1,$line-2),5)}
        $diagnostic|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $directory 'CompileFailure.json') -Encoding UTF8
        Write-Output ($diagnostic|ConvertTo-Json -Compress)
        throw 'Кандидат не скомпилирован.'
    }
    $app.AutomationSecurity=1
    $migrated=$app.Run("'"+$book.Name+"'!modTestSnapshotBuild.TestMigrateSavedSnapshot")
    $book.VBProject.VBComponents.Remove($migration)
    $newNames=@($book.Names|ForEach-Object Name|Where-Object {$_ -notin $oldNames})
    if($newNames.Count -ne 1 -or $newNames[0] -ne 'rngNDMLoadCaseResults'){throw 'Добавлено более одного диапазона или постороннее имя.'}
    $app.Calculate();$book.Save()
    [ordered]@{SourceSHA256=$sourceHash;Compiled=$true;Settings=@('Results.SnapshotStorage','Results.SnapshotName');Default='Excel';Migrated=$migrated;NewNames=$newNames;LowerOrder=$layout;Published=$false}|
        ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $directory 'Preparation.json') -Encoding UTF8
    Write-Output 'COMPLETE_SNAPSHOT_CANDIDATE_READY'
} finally {
    if(Test-OwnedExcel){try{if($book){$book.Close($false)}}catch{};try{if($app){$app.Quit()}}catch{}}
    if($app){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if(Test-OwnedExcel){Stop-Process -Id $owned -Force -ErrorAction SilentlyContinue}
}
Restore-WorkbookPrintAreas $path $printAreas
if((Get-FileHash -LiteralPath $source).Hash -cne $sourceHash){throw 'Исходная книга изменилась во время подготовки.'}
