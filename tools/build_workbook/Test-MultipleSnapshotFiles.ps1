# ==========================================================================
# ДЛЯ ТЕСТОВ: реальные расчёты разных сечений и переключение файловых снимков.
# ==========================================================================
# Работает только с новой частной копией. VBA не импортируется. Перед каждой
# загрузкой сохраняет книгу для независимого сравнения всех остальных листов.
param([Parameter(Mandatory=$true)][string]$WorkbookPath,
      [Parameter(Mandatory=$true)][string]$EvidenceDirectory)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$source=(Resolve-Path -LiteralPath $WorkbookPath).Path
$sourceHash=(Get-FileHash -LiteralPath $source).Hash
$directory=[IO.Path]::GetFullPath((Join-Path $root $EvidenceDirectory))
$allowed=(Join-Path $root 'docs/regression')+[IO.Path]::DirectorySeparatorChar
if(-not $directory.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $directory)){throw 'A fresh private evidence directory is required.'}
New-Item -ItemType Directory -Path $directory | Out-Null
$path=Join-Path $directory 'RC_Section_NDM.xlsm'
Copy-Item -LiteralPath $source -Destination $path
$previous=@(Get-Process EXCEL -ErrorAction SilentlyContinue|ForEach-Object Id)
Add-Type 'using System;using System.Runtime.InteropServices;public static class MultipleSnapshotTestOwner{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
$app=$null;$book=$null;$owned=0;$ticks=0;$cases=@();$restores=@();$expected=@{}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Закрывает только созданный здесь Excel с тем же временем старта.
function Test-OwnedMultipleSnapshotExcel {
    if(-not $owned -or $owned -in $previous){return $false}
    $process=Get-Process -Id $owned -ErrorAction SilentlyContinue
    return $process -and $process.ProcessName -eq 'EXCEL' -and $process.StartTime.ToUniversalTime().Ticks -eq $ticks
}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Изменяет существующее поле частного Config по его ключу, без фиксированного адреса.
function Test-SetSnapshotInput([string]$Key,[string]$Value) {
    $table=$book.Names.Item('rngSystemSettings').RefersToRange;$values=$table.Value2
    for($row=1;$row -le $values.GetLength(0);$row++){
        if([string]$values[$row,1] -ceq $Key){$table.Cells.Item($row,2).Value2=$Value;return}
    }
    throw ('Missing setting: '+$Key)
}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Хеширует точные строки существующего VBA-теста, включая все 16 таблиц и потребителей.
function Test-SnapshotDataHash {
    $text=[string]$app.Run("'$($book.Name)'!modTestCalculationSnapshot.TestSnapshotReopenFingerprint")
    $digest=[Security.Cryptography.SHA256]::Create()
    try{return [BitConverter]::ToString($digest.ComputeHash([Text.Encoding]::UTF8.GetBytes($text))).Replace('-','')}
    finally{$digest.Dispose()}
}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Сравнивает серии и заголовок ChartObject; фигуры сверяет отдельный OOXML-проверяющий.
function Test-SnapshotChartHash {
    $chart=$book.Worksheets.Item('Расчет').ChartObjects().Item(1).Chart
    $series=@(foreach($item in $chart.SeriesCollection()){
        [ordered]@{Name=[string]$item.Name;X=$item.XValues;Y=$item.Values}
    })
    $value=[ordered]@{Title=[string]$chart.ChartTitle.Text;Series=$series}|ConvertTo-Json -Depth 12 -Compress
    $digest=[Security.Cryptography.SHA256]::Create()
    try{return [BitConverter]::ToString($digest.ComputeHash([Text.Encoding]::UTF8.GetBytes($value))).Replace('-','')}
    finally{$digest.Dispose()}
}

try {
    $app=New-Object -ComObject Excel.Application
    [uint32]$id=0;[void][MultipleSnapshotTestOwner]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id)
    $owned=[int]$id;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
    if(-not(Test-OwnedMultipleSnapshotExcel)){throw 'Excel ownership not established.'}
    [ordered]@{PID=$owned;StartTicks=[string]$ticks;SourceSHA256=$sourceHash}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $directory 'Ownership.json') -Encoding UTF8
    Write-Output ('OWNED_EXCEL: id='+$owned+'; startTicks='+$ticks)
    $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$true;$app.AutomationSecurity=1
    $book=$app.Workbooks.Open($path,0,$false)
    $definitions=@(
        @{Name='Круг';Shape='Circle';Count=1},
        @{Name='Клин';Shape='PointedSection';Count=2},
        @{Name='Полый';Shape='HollowRectangle';Count=3},
        @{Name='Прямоугольники';Shape='Rectangles';Count=4},
        @{Name='Круг_трещины';Shape='Circle';Count=2;Family='CRACK_SP35'}
    )
    foreach($definition in $definitions){
        $family='MIXED';if($definition.Family){$family=[string]$definition.Family}
        $app.Run("'$($book.Name)'!modTestPerformance.TestConfigurePerformanceFixture",$family,[int]$definition.Count,'Newton',[string]$definition.Shape,'Normal','Yes')
        Test-SetSnapshotInput 'Results.SnapshotStorage' 'File'
        Test-SetSnapshotInput 'Results.SnapshotName' $definition.Name
        Test-SetSnapshotInput 'General.ExecutionReportEnabled' 'No'
        Test-SetSnapshotInput 'General.NonCriticalMessagesEnabled' 'No'
        Test-SetSnapshotInput 'Plot.Enabled' 'Yes'
        Test-SetSnapshotInput 'Plot.AutoUpdateAfterCalculation' 'Yes'
        Test-SetSnapshotInput 'Plot.LoadCase' 'Worst'
        Test-SetSnapshotInput 'AutoCAD.Export.CombinationID' 'Worst'
        Test-SetSnapshotInput 'Stability.ElementLength' '1000'
        Test-SetSnapshotInput 'General.WorstCombinationCriterion' $(if($family -eq 'MIXED'){'StrengthCapacity'}else{'Cracks'})
        foreach($rangeName in @('rngLoadCombinations','rngStabilityDurationLoads')){
            $range=$book.Names.Item($rangeName).RefersToRange;$values=$range.Value2
            for($row=2;$row -le $values.GetLength(0);$row++){
                if($values[$row,1]){
                    $values[$row,1]=$definition.Name+'_'+[string]$values[$row,1]
                    if($family -eq 'CRACK_SP35' -and $rangeName -eq 'rngLoadCombinations'){
                        $values[$row,3]=[double]$values[$row,3]*0.25
                        $values[$row,4]=[double]$values[$row,4]*0.25 # Для штатного круга момент остаётся выше трещинообразования, но ниже физического предела.
                    }
                }
            }
            $range.Value2=$values
        }
        $formation=[string]$app.Run("'$($book.Name)'!modWorkbookCalculation.FormSectionForWorkbook",$book)
        $timer=[Diagnostics.Stopwatch]::StartNew()
        $message=[string]$app.Run("'$($book.Name)'!modWorkbookCalculation.RunSectionCalculationForWorkbook",$book)
        $timer.Stop()
        if($message -match 'InputErr|NumFail|Расчёт не запускался'){throw ('Calculation fixture failed: '+$definition.Name+' '+$message)}
        $file=Join-Path $directory ($definition.Name+'.ndm')
        if(-not(Test-Path -LiteralPath $file)){throw 'Snapshot was not saved.'}
        $save=$definition.Name+'_Expected.xlsm';$book.SaveCopyAs((Join-Path $directory $save))
        $item=[ordered]@{Name=$definition.Name;Shape=$definition.Shape;Combinations=$definition.Count;Seconds=$timer.Elapsed.TotalSeconds;DataHash=(Test-SnapshotDataHash);ChartHash=(Test-SnapshotChartHash);FileSHA256=(Get-FileHash -LiteralPath $file).Hash;ExpectedBook=$save;Message=$message;Formation=$formation}
        $cases+=,$item;$expected[$definition.Name]=$item
        Write-Output ('CALCULATED: '+$definition.Name+'; combinations='+$definition.Count)
    }
    # Имя выбирает пользователь; остальные входы намеренно относятся к другой задаче.
    Test-SetSnapshotInput 'Geometry.Type' 'Circle'
    Test-SetSnapshotInput 'Solver.MaxIterations' 'not_a_number'
    $loads=$book.Names.Item('rngLoadCombinations').RefersToRange
    $loads.Cells.Item(2,1).Value2='CURRENT_INPUT_ONLY';$loads.Cells.Item(2,2).Value2=777.125
    $index=0
    foreach($name in @('Клин','Круг_трещины','Круг','Прямоугольники','Полый','Круг_трещины','Клин')){
        $index++;Test-SetSnapshotInput 'Results.SnapshotName' $name
        Test-SetSnapshotInput 'General.ReformatResultsEnabled' $(if($index%2){'No'}else{'Yes'})
        $before='Restore'+$index+'_Before.xlsm';$after='Restore'+$index+'_After.xlsm'
        $book.SaveCopyAs((Join-Path $directory $before))
        $app.Run("'$($book.Name)'!modSolverWorkStats.ResetSectionEquilibriumSolveCount")
        $app.Run("'$($book.Name)'!modCalculationSnapshot.UpdateSectionResultsForWorkbook",$book)
        $solves=[int]$app.Run("'$($book.Name)'!modSolverWorkStats.SectionEquilibriumSolveCount")
        if($solves -ne 0){throw 'Restoration invoked the solver.'}
        if(-not [bool]$app.Run("'$($book.Name)'!modCalculationSnapshot.TestCalculationSnapshotReadReleased")){throw 'Read context was not released.'}
        $dataHash=Test-SnapshotDataHash;$chartHash=Test-SnapshotChartHash
        if($dataHash -cne $expected[$name].DataHash -or $chartHash -cne $expected[$name].ChartHash){throw ('Wrong snapshot restored: '+$name)}
        $book.SaveCopyAs((Join-Path $directory $after))
        $restores+=,[ordered]@{Name=$name;Before=$before;After=$after;ExpectedBook=$expected[$name].ExpectedBook;DataHash=$dataHash;ChartHash=$chartHash;SolverCalls=$solves;ReadContextReleased=$true}
        Write-Output ('RESTORED: '+$name+'; solver=0; payload/header/series exact; shapes await independent review')
    }
    # После открытия заново используем именно макрос штатной кнопки.
    $book.Save();$book.Close($false);$book=$null
    $book=$app.Workbooks.Open($path,0,$false)
    $action=[string]$book.Worksheets.Item('Расчет').Shapes.Item('btnUpdateSectionPlot').OnAction
    if($action -notmatch '(^|!)UpdateSectionPlot$'){throw 'Unexpected update button macro.'}
    foreach($definition in $definitions){
        $name=[string]$definition.Name;$index++
        Test-SetSnapshotInput 'Results.SnapshotName' $name
        $before='Restore'+$index+'_Before.xlsm';$after='Restore'+$index+'_After.xlsm'
        $book.SaveCopyAs((Join-Path $directory $before))
        $app.Run("'$($book.Name)'!modSolverWorkStats.ResetSectionEquilibriumSolveCount")
        $app.Run("'$($book.Name)'!UpdateSectionPlot")
        $solves=[int]$app.Run("'$($book.Name)'!modSolverWorkStats.SectionEquilibriumSolveCount")
        if($solves -ne 0){throw 'Button restoration invoked the solver.'}
        if(-not [bool]$app.Run("'$($book.Name)'!modCalculationSnapshot.TestCalculationSnapshotReadReleased")){throw 'Read context was not released.'}
        $dataHash=Test-SnapshotDataHash;$chartHash=Test-SnapshotChartHash
        if($dataHash -cne $expected[$name].DataHash -or $chartHash -cne $expected[$name].ChartHash){throw ('Wrong snapshot after reopen: '+$name)}
        $book.SaveCopyAs((Join-Path $directory $after))
        $restores+=,[ordered]@{Name=$name;Before=$before;After=$after;ExpectedBook=$expected[$name].ExpectedBook;DataHash=$dataHash;ChartHash=$chartHash;SolverCalls=$solves;ReadContextReleased=$true;AfterReopen=$true;ButtonMacro=$action}
        Write-Output ('REOPEN_BUTTON_RESTORED: '+$name+'; solver=0')
    }
    $proof=[ordered]@{Status='NATIVE_PASS_PENDING_INDEPENDENT_SHEET_REVIEW';SourceSHA256=$sourceHash;ImportedVBA=0;Cases=$cases;Restores=$restores;OwnedPID=$owned;StartTicks=[string]$ticks}
    $proof|ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $directory 'NativeMultipleSnapshots.json') -Encoding UTF8
} finally {
    if(Test-OwnedMultipleSnapshotExcel){if($book){$book.Close($false)};if($app){$app.Quit()}}
    if($app){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
if((Get-FileHash -LiteralPath $source).Hash -cne $sourceHash){throw 'The published source workbook changed.'}
