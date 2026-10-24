# ======================================================================
# ДЛЯ ТЕСТОВ: фазы снимка в частных книгах, сравнение памяти и повторной загрузки.
# ======================================================================
# Значения, схема и исходная книга проверяются вне таймера. Полные suites не запускаются.
param(
    [Parameter(Mandatory=$true)][string]$SourceWorkbook,
    [Parameter(Mandatory=$true)][string]$EvidenceDirectory,
    [ValidateSet('Standard','Real')][string]$Fixture='Standard',
    [string]$FixtureWorkbook='',
    [int[]]$Counts=@(1,10,30),
    [ValidateRange(1,20)][int]$Repeats=5,
    [switch]$CompareReload,
    [switch]$PoisonResults
)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$source=(Resolve-Path -LiteralPath (Join-Path $root $SourceWorkbook)).Path
$directory=[IO.Path]::GetFullPath((Join-Path $root $EvidenceDirectory))
$allowed=(Join-Path $root 'docs/regression')+[IO.Path]::DirectorySeparatorChar
if(-not $directory.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)){throw 'Private evidence directory required.'}
if(Test-Path -LiteralPath (Join-Path $directory 'Measurements.json')){throw 'A new evidence directory is required.'}
[void][IO.Directory]::CreateDirectory($directory)
$sourceHash=(Get-FileHash -LiteralPath $source).Hash
$fixtureHash='';if($Fixture -eq 'Real'){$fixturePath=(Resolve-Path -LiteralPath (Join-Path $root $FixtureWorkbook)).Path;$fixtureHash=(Get-FileHash -LiteralPath $fixturePath).Hash}
$copy=Join-Path $directory 'Profile.xlsm';Copy-Item -LiteralPath $source -Destination $copy
Add-Type 'using System;using System.Runtime.InteropServices;public static class SnapshotPhaseOwner{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr w,out uint p);}'
$app=$null;$book=$null;$other=$null;$owned=0;$ticks=0;$previous=@(Get-Process EXCEL -ErrorAction SilentlyContinue|ForEach-Object Id)
$records=New-Object 'System.Collections.Generic.List[object]'

# Проверяет время старта единственного принадлежащего тесту Excel.
function Test-OwnedSnapshotPhaseExcel {
    if(-not $owned -or $owned -in $previous){return $false}
    $process=Get-Process -Id $owned -ErrorAction SilentlyContinue
    return $process -and $process.ProcessName -eq 'EXCEL' -and $process.StartTime.ToUniversalTime().Ticks -eq $ticks
}

# Точечно меняет значение настройки; неизвестный ключ не создаётся.
function Set-TestSnapshotPhaseSetting([string]$Key,[string]$Value){
    $table=$book.Names.Item('rngSystemSettings').RefersToRange;$values=$table.Value2
    for($r=1;$r -le $values.GetLength(0);$r++){if([string]$values[$r,1] -ceq $Key){$table.Cells.Item($r,2).Value2=$Value;return}}
    throw ('Missing setting: '+$Key)
}

# Добавляет измерение вокруг одной существующей строки, не подменяя её результат.
function Add-TestSnapshotPhaseLine([string]$Module,[string]$Method,[string]$Needle,[string]$Phase){
    $code=$book.VBProject.VBComponents.Item($Module).CodeModule
    $first=$code.ProcStartLine($Method,0);$count=$code.ProcCountLines($Method,0);$body=[string]$code.Lines($first,$count)
    $matches=[regex]::Matches($body,[regex]::Escape($Needle),'IgnoreCase')
    if($matches.Count -ne 1){throw ('Phase anchor mismatch: '+$Module+'.'+$Method+': '+$Needle)}
    $actual=$matches[0].Value
    $replacement='    modTestSnapshotIOProfiler.TestBegin "'+$Phase+'"'+"`r`n"+$actual+"`r`n"+'    modTestSnapshotIOProfiler.TestEnd "'+$Phase+'"'
    $body=$body.Replace($actual,$replacement);$code.DeleteLines($first,$count);$code.InsertLines($first,$body)
}

# Вставляет диагностическую строку в начало тела процедуры после всех продолжений сигнатуры.
function Add-TestSnapshotPhaseEntry([string]$Module,[string]$Method,[string]$Text){
    $code=$book.VBProject.VBComponents.Item($Module).CodeModule
    $first=$code.ProcStartLine($Method,0);$count=$code.ProcCountLines($Method,0);$body=[string]$code.Lines($first,$count)
    $signature=[regex]::Match($body,'(?ms)^(?:Public |Private |Friend )?(?:Sub|Function) [^\r\n]+(?:_\r?\n[^\r\n]+)*\r?\n')
    if(-not $signature.Success){throw ('Procedure entry mismatch: '+$Module+'.'+$Method)}
    $offset=$signature.Index+$signature.Length;$body=$body.Insert($offset,$Text+"`r`n")
    $code.DeleteLines($first,$count);$code.InsertLines($first,$body)
}

# Считает чтения сохранённых данных; контрольные суммы файла учитываются своим этапом.
function Get-TestSnapshotPhaseDump {
    $text=[string]$app.Run("'$($book.Name)'!modTestSnapshotIOProfiler.TestDump")
    $result=[ordered]@{}
    foreach($line in $text -split "`n"){if(-not $line.Trim()){continue};$parts=$line.Trim() -split "`t";$result[$parts[0]]=[double]::Parse($parts[1],[Globalization.CultureInfo]::InvariantCulture)}
    return $result
}

try{
    $app=New-Object -ComObject Excel.Application
    [uint32]$id=0;[void][SnapshotPhaseOwner]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id)
    $owned=[int]$id;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
    if(-not(Test-OwnedSnapshotPhaseExcel)){throw 'Excel ownership not established.'}
    $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=1
    $book=$app.Workbooks.Open($copy,0,$false)
    [ordered]@{Source=$SourceWorkbook;SourceSHA256=$sourceHash;Fixture=$Fixture;FixtureWorkbook=$FixtureWorkbook;FixtureSHA256=$fixtureHash;PID=$owned;StartTicks=[string]$ticks;CompareReload=[bool]$CompareReload;PoisonResults=[bool]$PoisonResults}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $directory 'Environment.json') -Encoding UTF8
    $helper=@'
Option Explicit
' ==================== ДЛЯ ТЕСТОВ ====================
' Короткие счётчики и QPC принадлежат только частной измерительной книге.
Private Declare PtrSafe Function TestQpc Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef value As Currency) As Long
Private Declare PtrSafe Function TestQpf Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef value As Currency) As Long
Private mTestStarted As Object, mTestValues As Object
Private mTestSolved As Boolean, mTestSaved As Boolean, mTestPoison As Boolean
Public TestForceReload As Boolean
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает монотонное время без обращения к листам и файлам.
Private Function TestNow() As Double
    Dim ticks As Currency, frequency As Currency
    TestQpc ticks: TestQpf frequency: TestNow = ticks / frequency
End Function
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Начинает независимый запуск без остатка счётчиков предыдущего опыта.
Public Sub TestReset(ByVal forceReload As Boolean, ByVal poison As Boolean)
    Set mTestStarted = CreateObject("Scripting.Dictionary"): Set mTestValues = CreateObject("Scripting.Dictionary")
    mTestSolved = False: mTestSaved = False: mTestPoison = poison: TestForceReload = forceReload
End Sub
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Открывает именованный интервал; вложенные разные этапы не мешают друг другу.
Public Sub TestBegin(ByVal name As String)
    mTestStarted(name) = TestNow()
End Sub
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Суммирует реальные вызовы одного этапа, включая несколько нижних таблиц.
Public Sub TestEnd(ByVal name As String)
    If Not mTestValues.Exists(name) Then mTestValues.Add name, 0#
    mTestValues(name) = mTestValues(name) + TestNow() - CDbl(mTestStarted(name))
End Sub
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Отмечает логическое чтение полного снимка, отдельно от хеширования файла.
Public Sub TestRead(ByVal name As String)
    If Not mTestValues.Exists(name) Then mTestValues.Add name, 0#
    mTestValues(name) = mTestValues(name) + 1#
    If name = "ReadPhysicalAfterSolve" And mTestPoison And mTestSolved And Not mTestSaved Then Err.Raise 5, , "Результаты читаются с листа после расчёта до записи снимка."
End Sub
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Уничтожает только результаты частной книги; готовый расчёт обязан остаться в памяти.
Public Sub TestAfterSolve()
    mTestSolved = True
    If mTestPoison Then ThisWorkbook.Worksheets("Results").UsedRange.ClearContents
End Sub
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Учитывает только чтения между окончанием решения и сохранением снимка.
Public Sub TestPhysicalRead()
    If mTestSolved And Not mTestSaved Then TestRead "ReadPhysicalAfterSolve"
End Sub
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Отмечает границу, после которой выполняется лишь показ уже сохранённого снимка.
Public Sub TestAfterSave()
    mTestSaved = True
End Sub
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает компактные времена и счётчики без таблиц пользовательских нагрузок.
Public Function TestDump() As String
    Dim key As Variant
    For Each key In mTestValues.Keys
        TestDump = TestDump & CStr(key) & vbTab & Replace$(CStr(mTestValues(key)), ",", ".") & vbLf
    Next key
End Function
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Даёт ручную кнопку обновления для таймера, не вызывая расчёт или формирование геометрии.
Public Function TestManualRefresh() As String
    On Error GoTo Failed
    UpdateSectionResultsForWorkbook ThisWorkbook
    Exit Function
Failed:
    TestManualRefresh = "TEST_CALCULATION_ERROR: " & Err.Description
End Function
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает ожидаемую ошибку контроля в runner, не оставляя Excel в окне отладчика.
Public Function TestCalculate(ByVal workbook As Object) As String
    On Error GoTo Failed
    TestCalculate = RunSectionCalculationForWorkbook(workbook, False)
    Exit Function
Failed:
    TestCalculate = "TEST_CALCULATION_ERROR: " & Err.Description
End Function
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет общий контекст памяти после выхода штатной команды.
Public Function TestReleased() As Boolean
    TestReleased = TestCalculationSnapshotReadReleased()
End Function
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Меняет только частный путь после сохранения для честного сравнения повторной загрузки.
Public Sub TestDisplay(ByVal workbook As Object, ByVal snapshot As CCalculationSnapshot, ByRef warning As String)
    TestAfterSave
    If TestForceReload Then
        UpdateSectionResultsForWorkbook workbook
    Else
        UpdateSectionResultsForWorkbook workbook, snapshot, warning
    End If
End Sub
'@
    $component=$book.VBProject.VBComponents.Add(1);$component.Name='modTestSnapshotIOProfiler';$component.CodeModule.AddFromString($helper)
    Add-TestSnapshotPhaseLine 'modWorkbookCalculation' 'RunSectionCalculationForWorkbook' '    batch.Execute' 'Solve'
    $code=$book.VBProject.VBComponents.Item('modWorkbookCalculation').CodeModule
    $body=[string]$code.Lines(1,$code.CountOfLines);$body=$body.Replace('    modTestSnapshotIOProfiler.TestEnd "Solve"','    modTestSnapshotIOProfiler.TestEnd "Solve"'+"`r`n"+'    modTestSnapshotIOProfiler.TestAfterSolve');$code.DeleteLines(1,$code.CountOfLines);$code.AddFromString($body)
    Add-TestSnapshotPhaseLine 'modWorkbookCalculation' 'RunSectionCalculationForWorkbook' '    summaryWriter.WriteSummary workbook, batch, units, section, fullFormatting' 'PrepareSummary'
    Add-TestSnapshotPhaseLine 'modWorkbookCalculation' 'RunSectionCalculationForWorkbook' '    ndmWriter.WriteResults workbook, section, PrepareSectionSnapshot(section, materialProvider), batch, units, True, fullFormatting' 'PrepareTechnical'
    Add-TestSnapshotPhaseLine 'modCalculationSnapshot' 'SaveAndDisplayCalculationSnapshot' '        store.SaveFile workbook, snapshot' 'SaveFile'
    $code=$book.VBProject.VBComponents.Item('modCalculationSnapshot').CodeModule
    $body=[string]$code.Lines(1,$code.CountOfLines);$needle='    UpdateSectionResultsForWorkbook workbook, snapshot, warning'
    $matches=[regex]::Matches($body,[regex]::Escape($needle),'IgnoreCase');if($matches.Count -ne 1){throw 'Display anchor mismatch.'}
    $body=$body.Replace($matches[0].Value,'    modTestSnapshotIOProfiler.TestDisplay workbook, snapshot, warning');$code.DeleteLines(1,$code.CountOfLines);$code.AddFromString($body)
    Add-TestSnapshotPhaseLine 'modCalculationSnapshot' 'UpdateSectionResultsForWorkbook' '    writer.RestoreSnapshotBase workbook, snapshot, fullFormatting, mode = "Excel"' 'WriteBase'
    Add-TestSnapshotPhaseLine 'modCalculationSnapshot' 'UpdateSectionResultsForWorkbook' '    summary.RestoreSummary workbook, snapshot, fullFormatting' 'WriteSummary'
    Add-TestSnapshotPhaseLine 'modCalculationSnapshot' 'UpdateSectionResultsForWorkbook' '            Set snapshot = store.LoadFile(workbook)' 'LoadFile'
    Add-TestSnapshotPhaseLine 'modCalculationSnapshot' 'UpdateSectionResultsForWorkbook' '        store.LinkFile workbook' 'BindFile'
    $noticeCode=$book.VBProject.VBComponents.Item('modCalculationSnapshot').CodeModule
    if(([string]$noticeCode.Lines(1,$noticeCode.CountOfLines)) -match '(?i)store\.WriteFileNotice workbook'){
        Add-TestSnapshotPhaseLine 'modCalculationSnapshot' 'UpdateSectionResultsForWorkbook' '        store.WriteFileNotice workbook' 'WriteFileNotice'
    }
    Add-TestSnapshotPhaseLine 'modCalculationSnapshot' 'UpdateSectionResultsForWorkbook' '        plotWarning = TryUpdateCalculatedSnapshotPlot(workbook, snapshot)' 'Plot'
    Add-TestSnapshotPhaseLine 'modCalculationSnapshot' 'UpdateSectionResultsForWorkbook' '        UpdateSectionPlotForWorkbook workbook, False' 'Plot'
    Add-TestSnapshotPhaseLine 'modWorkbookCalculation' 'RunSectionCalculationForWorkbook' '    snapshotStore.UpdateSavedElapsed workbook, savedSnapshot' 'FinalizeFileOrSheet'
    Add-TestSnapshotPhaseEntry 'CCalculationSnapshotStore' 'ReadBinary' '    modTestSnapshotIOProfiler.TestRead "ReadBinary"'
    Add-TestSnapshotPhaseEntry 'modSectionSnapshot' 'ReadPhysicalResultTable' '    modTestSnapshotIOProfiler.TestPhysicalRead'
    $app.VBE.ActiveVBProject=$book.VBProject;$compile=$app.VBE.CommandBars.FindControl(1,578);if($compile.Enabled){$compile.Execute()};if($compile.Enabled){throw 'Instrumented project does not compile.'}
    $book.Save()
    if($Fixture -eq 'Real'){
        $other=$app.Workbooks.Open($fixturePath,0,$true)
        $old=$other.Names.Item('rngSystemSettings').RefersToRange.Value2;$target=$book.Names.Item('rngSystemSettings').RefersToRange;$current=$target.Value2
        $settings=@{};for($r=1;$r -le $old.GetLength(0);$r++){$settings[[string]$old[$r,1]]=$old[$r,2]}
        for($r=1;$r -le $current.GetLength(0);$r++){$key=[string]$current[$r,1];if($settings.ContainsKey($key)){$value=$settings[$key];if($value -is [double]){$value=$value.ToString('R',[Globalization.CultureInfo]::InvariantCulture)};$target.Cells.Item($r,2).Value2=[string]$value}}
        foreach($name in @('rngUnitSettings','rngSignConventionSettings','rngCalculationProfiles','rngLoadCombinations','rngStabilityDurationLoads','rngSP35Table721')){$book.Names.Item($name).RefersToRange.Value2=$other.Names.Item($name).RefersToRange.Value2}
        foreach($pair in @(@('rngNDMSectionGeometry',16),@('rngNDMSectionContours',17),@('rngNDMSectionProperties',6),@('rngNDMSectionAnnotations',24))){
            $anchor=$other.Names.Item($pair[0]).RefersToRange;$last=[int]$anchor.Worksheet.Cells.Item($anchor.Worksheet.Rows.Count,$anchor.Column).End(-4162).Row
            $rows=[Math]::Max(1,$last-$anchor.Row+1);$book.Names.Item($pair[0]).RefersToRange.Resize($rows,[int]$pair[1]).Value2=$anchor.Resize($rows,[int]$pair[1]).Value2
        }
        $other.Close($false);$other=$null
    }
    $allLoads=$book.Names.Item('rngLoadCombinations').RefersToRange.Value2;$allDuration=$book.Names.Item('rngStabilityDurationLoads').RefersToRange.Value2
    foreach($count in $Counts){
        if($Fixture -eq 'Standard'){
            [void]$app.Run("'$($book.Name)'!modTestPerformance.TestConfigurePerformanceFixture",'CAPACITY_CRACK_STABILITY_SP63',[int]$count,'Newton','Rectangles','Normal','Yes')
            Set-TestSnapshotPhaseSetting 'Geometry.Source' 'Generated';Set-TestSnapshotPhaseSetting 'Geometry.Type' 'Rectangles';Set-TestSnapshotPhaseSetting 'Mesh.StepMode' 'Auto 0.1';Set-TestSnapshotPhaseSetting 'Mesh.BoundarySubdivisions' '2'
            $table=$book.Names.Item('rngRectanglesGeometry').RefersToRange;$data=$table.Value2
            for($r=1;$r -le $data.GetLength(0);$r++){switch -CaseSensitive ([string]$data[$r,1]){'Shape'{$table.Cells.Item($r,3).Value2='OneRect'}'H1'{$table.Cells.Item($r,3).Value2='550'}'B1'{$table.Cells.Item($r,3).Value2='250'}}}
        }else{
            $book.Names.Item('rngLoadCombinations').RefersToRange.Value2=$allLoads;$book.Names.Item('rngStabilityDurationLoads').RefersToRange.Value2=$allDuration
            foreach($name in @('rngLoadCombinations','rngStabilityDurationLoads')){$table=$book.Names.Item($name).RefersToRange;if($count+1 -lt $table.Rows.Count){$table.Offset($count+1,0).Resize($table.Rows.Count-$count-1,$table.Columns.Count).ClearContents()}}
            Set-TestSnapshotPhaseSetting 'Plot.LoadCase' ([string]$allLoads[2,1])
        }
        Set-TestSnapshotPhaseSetting 'General.ExecutionReportEnabled' 'Yes';Set-TestSnapshotPhaseSetting 'General.ReformatResultsEnabled' 'Yes';Set-TestSnapshotPhaseSetting 'Plot.Enabled' 'Yes';Set-TestSnapshotPhaseSetting 'Plot.AutoUpdateAfterCalculation' 'Yes';Set-TestSnapshotPhaseSetting 'Results.SnapshotName' 'TestSnapshotIO'
        if($Fixture -eq 'Standard'){[void]$app.Run("'$($book.Name)'!modWorkbookCalculation.FormSectionForWorkbook",$book)}
        $anchor=$book.Names.Item('rngNDMSectionGeometry').RefersToRange;$last=[int]$anchor.Worksheet.Cells.Item($anchor.Worksheet.Rows.Count,$anchor.Column).End(-4162).Row
        $geometryRows=$anchor.Resize($last-$anchor.Row+1,3).Value2;$concrete=0;$rebar=0
        for($row=2;$row -le $geometryRows.GetLength(0);$row++){if([string]$geometryRows[$row,3] -ceq 'Concrete'){$concrete++}elseif([string]$geometryRows[$row,3] -ceq 'Rebar'){$rebar++}}
        $paths=@('Memory');if($CompareReload){$paths+=,'Reload'}
        for($repeat=0;$repeat -le $Repeats;$repeat++){
            foreach($path in $paths){foreach($mode in @('File','Excel')){
                if($mode -eq 'Excel' -and $path -eq 'Reload'){continue} # В Excel нет архивного пути повторного восстановления сводок.
                Set-TestSnapshotPhaseSetting 'Results.SnapshotStorage' $mode
                [void]$app.Run("'$($book.Name)'!modTestSnapshotIOProfiler.TestReset",[bool]($path -eq 'Reload'),[bool]$PoisonResults)
                $watch=[Diagnostics.Stopwatch]::StartNew();$message=[string]$app.Run("'$($book.Name)'!modTestSnapshotIOProfiler.TestCalculate",$book);$watch.Stop()
                if($message.StartsWith('TEST_CALCULATION_ERROR:')){throw $message}
                $phases=Get-TestSnapshotPhaseDump
                $released=[bool]$app.Run("'$($book.Name)'!modTestSnapshotIOProfiler.TestReleased")
                if(-not $released){throw 'Calculation read context leaked.'}
                $fingerprint=[string]$app.Run("'$($book.Name)'!modTestPerformance.TestPerformanceResultsFingerprint")
                $fingerprint+="`n"+[string]$app.Run("'$($book.Name)'!modTestPerformance.TestPerformancePlotFingerprint")
                $fingerprint=[regex]::Replace($fingerprint,'(?m)(^|\|)D:0000000000000080(?=\|)','$1D:0000000000000000')
                $digest=[Security.Cryptography.SHA256]::Create();try{$fingerprintHash=[BitConverter]::ToString($digest.ComputeHash([Text.Encoding]::UTF8.GetBytes($fingerprint))).Replace('-','')}finally{$digest.Dispose()}
                [void]$app.Run("'$($book.Name)'!modTestSnapshotIOProfiler.TestReset",$false,$false)
                $manual=[Diagnostics.Stopwatch]::StartNew();$manualMessage=[string]$app.Run("'$($book.Name)'!modTestSnapshotIOProfiler.TestManualRefresh");$manual.Stop()
                if($manualMessage.StartsWith('TEST_CALCULATION_ERROR:')){throw $manualMessage}
                $manualPhases=Get-TestSnapshotPhaseDump
                $manualReleased=[bool]$app.Run("'$($book.Name)'!modTestSnapshotIOProfiler.TestReleased")
                if(-not $manualReleased){throw 'Manual refresh read context leaked.'}
                $after=[string]$app.Run("'$($book.Name)'!modTestPerformance.TestPerformanceResultsFingerprint")
                $after+="`n"+[string]$app.Run("'$($book.Name)'!modTestPerformance.TestPerformancePlotFingerprint")
                $after=[regex]::Replace($after,'(?m)(^|\|)D:0000000000000080(?=\|)','$1D:0000000000000000')
                if($after -cne $fingerprint){throw 'Manual refresh changed results or plot.'}
                $record=[ordered]@{Count=$count;Mode=$mode;Path=$path;Repeat=$repeat;Warmup=($repeat -eq 0);Seconds=$watch.Elapsed.TotalSeconds;Phases=$phases;Fingerprint=$fingerprintHash;ReadContextReleased=$released;ManualRefreshSeconds=$manual.Elapsed.TotalSeconds;ManualPhases=$manualPhases;ManualContextReleased=$manualReleased;PoisonResults=[bool]$PoisonResults;Geometry=@{Concrete=$concrete;Rebar=$rebar}}
                $records.Add($record);$records|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $directory 'Measurements.json') -Encoding UTF8
                Write-Output ($record|ConvertTo-Json -Depth 6 -Compress)
            }}
        }
        $book.SaveCopyAs((Join-Path $directory ('Results_'+$count+'.xlsm'))) # Вне таймера: независимая последующая сверка оформления Results.
        Copy-Item -LiteralPath (Join-Path $directory 'TestSnapshotIO.ndm') -Destination (Join-Path $directory ('Results_'+$count+'.ndm'))
    }
}catch{
    $reason=$_.Exception.Message
    $failurePhases=$null
    try{$failurePhases=Get-TestSnapshotPhaseDump}catch{}
    [ordered]@{Status='FAIL';Reason=$reason;PoisonResults=[bool]$PoisonResults;Phases=$failurePhases;SourceSHA256=$sourceHash}|
        ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $directory 'Failure.json') -Encoding UTF8
    throw
}finally{
    if(Test-OwnedSnapshotPhaseExcel){if($other){$other.Close($false)};if($book){$book.Close($false)};if($app){$app.Quit()}}
    if($app){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
if((Get-FileHash -LiteralPath $source).Hash -cne $sourceHash){throw 'Source workbook changed.'}
if($Fixture -eq 'Real' -and (Get-FileHash -LiteralPath $fixturePath).Hash -cne $fixtureHash){throw 'Source fixture changed.'}
Write-Output 'SNAPSHOT_IO_PHASES_COMPLETED'
