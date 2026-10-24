# ==========================================================================
# ДЛЯ ТЕСТОВ: реальные нагрузки верхней сводки в Excel и файловых снимках.
# ==========================================================================
# Все расчёты и изменения ввода выполняются только в новой частной копии.
param([Parameter(Mandatory=$true)][string]$WorkbookPath,
      [Parameter(Mandatory=$true)][string]$EvidenceDirectory,
      [ValidateSet('No','Yes')][string]$Mode='No',
      [switch]$Baseline,
      [string]$LegacySnapshot='')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$source=(Resolve-Path -LiteralPath $WorkbookPath).Path
$sourceHash=(Get-FileHash -LiteralPath $source).Hash
$directory=[IO.Path]::GetFullPath((Join-Path $root $EvidenceDirectory))
$allowed=(Join-Path $root 'docs/regression')+[IO.Path]::DirectorySeparatorChar
if(-not $directory.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $directory)){throw 'A fresh private evidence directory is required.'}
[void][IO.Directory]::CreateDirectory($directory)
$path=Join-Path $directory 'RC_Section_NDM.xlsm'
Copy-Item -LiteralPath $source -Destination $path
$previous=@(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
Add-Type 'using System;using System.Runtime.InteropServices;public static class BatchLoadsTestOwner{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
$app=$null;$book=$null;$owned=0;$ticks=0;$cases=@();$restores=@();$checks=0

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Исключает пользовательские экземпляры и повторно использованные PID.
function Test-OwnedBatchLoadsExcel {
    if(-not $owned -or $owned -in $previous){return $false}
    $process=Get-Process -Id $owned -ErrorAction SilentlyContinue
    return $process -and $process.ProcessName -eq 'EXCEL' -and $process.StartTime.ToUniversalTime().Ticks -eq $ticks
}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Читает/меняет только значение найденного ключа частного Config.
function Test-BatchLoadsSetting([string]$Key,[object]$Value) {
    $column=2;$field=$Key;$name='rngSystemSettings'
    if($Key.StartsWith('Units.')){$name='rngUnitSettings';$field=$Key.Split('.')[1];$column=4;if($Key.EndsWith('.Input')){$column=2}}
    $table=$book.Names.Item($name).RefersToRange;$values=$table.Value2
    for($row=1;$row -le $values.GetLength(0);$row++){
        if([string]$values[$row,1] -ceq $field){$table.Cells.Item($row,$column).Value2=$Value;return}
    }
    throw ('Missing test setting: '+$Key)
}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Сохраняет настоящую ошибку вместо продолжения с неполным набором данных.
function Test-BatchLoadsAssert([bool]$Condition,[string]$Message) {
    if(-not $Condition){throw $Message}
    $script:checks++
}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Обходит кэш типа COM-setter PowerShell после текстовых записей в Value2.
function Test-BatchLoadsWrite([object]$Range,[object]$Value) {
    [void]$Range.GetType().InvokeMember('Value2',[Reflection.BindingFlags]::SetProperty,$null,$Range,[object[]](,$Value))
}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Проверяет шапку и выбранное Worst по фактически рассчитанным строкам.
function Test-BatchLoadsOutput {
    $anchor=$book.Names.Item('rngBatchSummary').RefersToRange.Cells.Item(1,1)
    $data=$anchor.Resize(42,33).Value2
    Test-BatchLoadsAssert ($data[6,28] -ceq 'Нагрузки в точке приложения') 'New load heading is missing.'
    Test-BatchLoadsAssert ($data[11,29] -ceq 'Исходный изгибающий момент Mx') 'Original Mx comment is wrong.'
    Test-BatchLoadsAssert ($data[11,30] -ceq 'Исходный изгибающий момент My') 'Original My comment is wrong.'
    Test-BatchLoadsAssert (-not ([string]$data[11,31]).Contains([char]0x3b7)) 'Eta must not annotate N.'
    Test-BatchLoadsAssert (([string]$data[11,32]).Contains([char]0x3b7)) 'The moment comment lost eta.'
    Test-BatchLoadsAssert ($anchor.Offset(10,31).MergeArea.Columns.Count -eq 2) 'Only the two moment comments may be merged.'
    foreach($block in @(@(5,27,2,6),@(7,27,2,3),@(7,30,2,3))){
        $range=$anchor.Offset($block[0],$block[1]).Resize($block[2],$block[3])
        Test-BatchLoadsAssert ($range.Cells.Item(1,1).MergeArea.Address() -ceq $range.Address()) 'Wrong load-header merge.'
        Test-BatchLoadsAssert ([bool]$range.WrapText -and $range.HorizontalAlignment -eq -4108) 'Wrong load-header alignment.'
    }
    $worst=[string]$data[1,1]
    if($worst){
        $found=0
        for($row=13;$row -le 42;$row++){if([string]$data[$row,1] -ceq $worst){$found=$row;break}}
        Test-BatchLoadsAssert ($found -gt 0) 'Worst ID does not belong to the displayed combinations.'
        for($column=28;$column -le 33;$column++){
            Test-BatchLoadsAssert ($data[12,$column] -ceq $data[$found,$column]) 'Worst load row differs from its combination.'
        }
    }
    return $worst
}

try {
    $app=New-Object -ComObject Excel.Application
    [uint32]$id=0;[void][BatchLoadsTestOwner]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id)
    $owned=[int]$id;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
    if(-not(Test-OwnedBatchLoadsExcel)){throw 'Own Excel identity is not established.'}
    [ordered]@{PID=$owned;StartTicks=[string]$ticks}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $directory 'Ownership.json') -Encoding UTF8
    $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$true;$app.AutomationSecurity=1
    $book=$app.Workbooks.Open($path,0,$false)
    $plans=@(
        @{Name='Excel_Full';Storage='Excel';Format='Yes';Criterion='StrengthCapacity';Force='tf';Moment='tf*m';Shape='PointedSection';Family='MIXED';Count=3},
        @{Name='File_Minimal';Storage='File';Format='No';Criterion='Stability';Force='kN';Moment='kN*m';Shape='PointedSection';Family='MIXED';Count=3},
        @{Name='Excel_Filtered';Storage='Excel';Format='No';Criterion='StrengthStrain';Force='tf';Moment='tf*m';Shape='PointedSection';Family='MIXED';Count=3;Filtered=$true},
        @{Name='File_Cracks';Storage='File';Format='Yes';Criterion='Cracks';Force='tf';Moment='tf*m';Shape='Circle';Family='CRACK_SP35';Count=2}
    )
    foreach($plan in $plans){
        Test-BatchLoadsSetting 'Units.Force.Output' 'tf';Test-BatchLoadsSetting 'Units.Moment.Output' 'tf*m'
        $app.Run("'$($book.Name)'!modTestPerformance.TestConfigurePerformanceFixture",$plan.Family,[int]$plan.Count,'Newton',$plan.Shape,'Normal',$Mode)
        foreach($item in @(@('General.ExecutionReportEnabled','No'),@('General.NonCriticalMessagesEnabled','No'),@('Plot.Enabled','Yes'),@('Plot.AutoUpdateAfterCalculation','Yes'),@('Plot.LoadCase','Worst'),@('Stability.ElementLength','2000'),@('Stability.Code','SP63'),@('General.WorstCombinationCriterion',$plan.Criterion),@('Results.SnapshotStorage',$plan.Storage),@('Results.SnapshotName',$plan.Name),@('General.ReformatResultsEnabled',$plan.Format),@('Units.Force.Output',$plan.Force),@('Units.Moment.Output',$plan.Moment))){Test-BatchLoadsSetting $item[0] $item[1]}
        if($plan.Filtered){
            Test-BatchLoadsSetting 'Calculation.ZeroMomentPerDepth' '0.5'
            $profiles=$book.Names.Item('rngCalculationProfiles').RefersToRange;$values=$profiles.Value2
            for($row=1;$row -le $values.GetLength(0);$row++){if([string]$values[$row,2] -ceq 'Calculation.Stability.Enabled'){$profiles.Cells.Item($row,3).Value2='No'}}
        }
        $loads=$book.Names.Item('rngLoadCombinations').RefersToRange
        $loadValues=$loads.Value2
        if($plan.Filtered){
            $inputMoment=[double]$loadValues[2,3]
            $loadValues[2,3]=-[Math]::Abs($inputMoment)
            for($row=2;$row -le $plan.Count+1;$row++){$loadValues[$row,6]='Auto'} # Фильтр может обнулить момент, поэтому явный моментный путь здесь неприменим.
        }
        if($plan.Family -eq 'CRACK_SP35'){
            for($row=2;$row -le $plan.Count+1;$row++){for($column=3;$column -le 4;$column++){$loadValues[$row,$column]=[double]$loadValues[$row,$column]*0.25}}
        }
        Test-BatchLoadsWrite $loads $loadValues
        $formation=[string]$app.Run("'$($book.Name)'!modWorkbookCalculation.FormSectionForWorkbook",$book)
        $message=[string]$app.Run("'$($book.Name)'!modWorkbookCalculation.RunSectionCalculationForWorkbook",$book)
        Test-BatchLoadsAssert ($message.Contains('Расчет завершен')) ('Calculation did not finish: '+$message)
        Test-BatchLoadsAssert (-not $message.Contains('ошибками ввода')) ('Invalid calculation fixture: '+$message)
        $worst=[string]$book.Names.Item('rngBatchSummary').RefersToRange.Cells.Item(1,1).Value2
        if(-not $Baseline){$worst=Test-BatchLoadsOutput}
        $save=$plan.Name+'_Ready.xlsm';$book.SaveCopyAs((Join-Path $directory $save))
        $cases+=,[ordered]@{Name=$plan.Name;Storage=$plan.Storage;Format=$plan.Format;Criterion=$plan.Criterion;ForceUnit=$plan.Force;MomentUnit=$plan.Moment;Shape=$plan.Shape;Family=$plan.Family;Count=$plan.Count;Filtered=[bool]$plan.Filtered;WorstID=$worst;Workbook=$save;Message=$message}
        Write-Output ('CALCULATED: '+$Mode+'; '+$plan.Name+'; Worst='+$worst)
    }
    if(-not $Baseline){
        # Посторонний текущий ввод намеренно несовместим с сохранёнными расчётами.
        Test-BatchLoadsSetting 'Solver.MaxIterations' 'not_a_number'
        Test-BatchLoadsSetting 'Units.Force.Output' 'N';Test-BatchLoadsSetting 'Units.Moment.Output' 'N*mm'
        $loads.Cells.Item(2,1).Value2='CURRENT_INPUT_ONLY';Test-BatchLoadsWrite ($loads.Cells.Item(2,2)) 777.125
        foreach($name in @('File_Minimal','File_Cracks','File_Minimal')){
            Test-BatchLoadsSetting 'Results.SnapshotStorage' 'File';Test-BatchLoadsSetting 'Results.SnapshotName' $name
            $index=$restores.Count+1;$before='Restore'+$index+'_Before.xlsm';$after='Restore'+$index+'_After.xlsm'
            $book.SaveCopyAs((Join-Path $directory $before))
            $app.Run("'$($book.Name)'!modSolverWorkStats.ResetSectionEquilibriumSolveCount")
            $app.Run("'$($book.Name)'!modCalculationSnapshot.UpdateSectionResultsForWorkbook",$book)
            $calls=[int]$app.Run("'$($book.Name)'!modSolverWorkStats.SectionEquilibriumSolveCount")
            Test-BatchLoadsAssert ($calls -eq 0) 'Snapshot restore called the solver.'
            $worst=Test-BatchLoadsOutput
            Test-BatchLoadsAssert ($loads.Cells.Item(2,1).Value2 -ceq 'CURRENT_INPUT_ONLY' -and $loads.Cells.Item(2,2).Value2 -eq 777.125) 'Snapshot restore changed current loads.'
            $book.SaveCopyAs((Join-Path $directory $after))
            $restores+=,[ordered]@{Name=$name;Before=$before;After=$after;ExpectedBook=$name+'_Ready.xlsm';WorstID=$worst;SolverCalls=$calls}
        }
        $book.Save();$book.Close($false);$book=$null
        $book=$app.Workbooks.Open($path,0,$false)
        Test-BatchLoadsSetting 'Results.SnapshotName' 'File_Cracks'
        $book.SaveCopyAs((Join-Path $directory 'Restore4_Before.xlsm'))
        $app.Run("'$($book.Name)'!modSolverWorkStats.ResetSectionEquilibriumSolveCount")
        $app.Run("'$($book.Name)'!UpdateSectionPlot")
        $calls=[int]$app.Run("'$($book.Name)'!modSolverWorkStats.SectionEquilibriumSolveCount")
        Test-BatchLoadsAssert ($calls -eq 0) 'Reopened snapshot restore called the solver.'
        $worst=Test-BatchLoadsOutput
        $book.SaveCopyAs((Join-Path $directory 'Restore4_After.xlsm'))
        $restores+=,[ordered]@{Name='File_Cracks';Before='Restore4_Before.xlsm';After='Restore4_After.xlsm';ExpectedBook='File_Cracks_Ready.xlsm';WorstID=$worst;SolverCalls=$calls;AfterReopenButton=$true}
        if($LegacySnapshot){
            $legacy=(Resolve-Path -LiteralPath $LegacySnapshot).Path;$legacyHash=(Get-FileHash -LiteralPath $legacy).Hash
            Copy-Item -LiteralPath $legacy -Destination (Join-Path $directory 'Legacy.ndm')
            Test-BatchLoadsSetting 'Results.SnapshotName' 'Legacy'
            $book.SaveCopyAs((Join-Path $directory 'Legacy_Before.xlsm'))
            $app.Run("'$($book.Name)'!modSolverWorkStats.ResetSectionEquilibriumSolveCount")
            $app.Run("'$($book.Name)'!modCalculationSnapshot.UpdateSectionResultsForWorkbook",$book)
            $anchor=$book.Names.Item('rngBatchSummary').RefersToRange.Cells.Item(1,1)
            Test-BatchLoadsAssert ($app.WorksheetFunction.CountA($anchor.Offset(11,27).Resize(31,6)) -eq 0) 'Old file must not invent missing original loads.'
            Test-BatchLoadsAssert ($anchor.Offset(10,27).Value2 -ceq 'Не сохранено в этом снимке') 'Missing legacy data are not explained.'
            Test-BatchLoadsAssert ((Get-FileHash -LiteralPath $legacy).Hash -ceq $legacyHash) 'Historical snapshot changed.'
            $book.SaveCopyAs((Join-Path $directory 'Legacy_After.xlsm'))
        }
    }
    [ordered]@{Status='NATIVE_PASS_PENDING_INDEPENDENT_REVIEW';SourceSHA256=$sourceHash;Mode=$Mode;Baseline=[bool]$Baseline;Checks=$checks;Cases=$cases;Restores=$restores;ImportedVBA=0;OwnedPID=$owned;StartTicks=[string]$ticks}|ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $directory 'NativeSummaryLoads.json') -Encoding UTF8
} catch {
    [ordered]@{Status='FAIL';Reason=$_.Exception.Message;Stack=$_.ScriptStackTrace}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $directory 'Failure.json') -Encoding UTF8
    throw
} finally {
    if(Test-OwnedBatchLoadsExcel){if($book){$book.Close($false)};if($app){$app.Quit()}}
    if($app){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
if((Get-FileHash -LiteralPath $source).Hash -cne $sourceHash){throw 'The source workbook changed.'}
