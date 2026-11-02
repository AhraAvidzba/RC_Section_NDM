# ==================== ДЛЯ ТЕСТОВ ====================
# Только после actual full PASS читает private user book readonly/macros-disabled/manual-calculation, экспортирует VBA и сохранённые счётчики.
# Книгу не сохраняет и не вызывает ни одну рабочую VBA-процедуру, решение НДС или обновление схемы.
param([Parameter(Mandatory=$true)][ValidateSet('Main','Heavy')][string]$Label)
$ErrorActionPreference='Stop'
Import-Module C:/Windows/System32/WindowsPowerShell/v1.0/Modules/Microsoft.PowerShell.Utility/Microsoft.PowerShell.Utility.psd1 -Global
. (Join-Path $PSScriptRoot 'TestAssertCurrentRun02Pass60.ps1')
$pass=Test-AssertCurrentRun02Pass60 -RequireTaskRemoved -RequireAdditionalIO
$prep=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'PreparedUserBooks62.json') -Raw -Encoding UTF8|ConvertFrom-Json
if($prep.Status -cne 'PRIVATE_CODE_ONLY_BOOKS_PREPARED_NATIVE_VALIDATION_REQUIRED' -or $prep.AcceptedSourceSHA256 -cne $pass.SourceSHA256 -or -not $prep.OfficialFullPass -or $prep.Published){throw 'Guarded private preparation required'}
$matches=@($prep.Books|Where-Object Label -ceq $Label)
if($matches.Count -ne 1){throw 'Exact one prepared book required'}
$item=$matches[0];$source=[IO.Path]::GetFullPath([string]$item.Prepared)
$expected=Join-Path $PSScriptRoot ('PreparedUserBooks62/'+$Label+'/')
if(-not $source.StartsWith([IO.Path]::GetFullPath($expected),[StringComparison]::OrdinalIgnoreCase) -or (Get-FileHash -LiteralPath $source).Hash -cne $item.PreparedSHA256){throw 'Foreign or changed private candidate'}
$directory=Join-Path $PSScriptRoot ('UserBooksNative63/'+$Label)
if(Test-Path -LiteralPath $directory){throw 'Existing native directory; inspect its ownership before any repeat'}
if(Test-Path -LiteralPath (Join-Path $PSScriptRoot 'UserBooksNative63')){
    foreach($file in Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot 'UserBooksNative63') -Filter 'Own.json' -Recurse -File){
        $prior=Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8|ConvertFrom-Json
        Test-AssertRun02OwnTerminal60 ([pscustomobject]@{Name='EXCEL';PID=$prior.PID;StartTicksUTC=$prior.StartTicksUTC})
    }
}
$memory=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
if($null -eq $memory -or $null -eq $memory.AvailableMBytes -or $null -eq $memory.CommitLimit -or $null -eq $memory.CommittedBytes){throw 'Actual memory counters unavailable'}
$ram=[double]$memory.AvailableMBytes/1024;$commit=([double]$memory.CommitLimit-[double]$memory.CommittedBytes)/1GB
if($ram -lt 8 -or $commit -lt 14){throw 'Native reserve below8GiB RAM/14GiB commit'}
[void][IO.Directory]::CreateDirectory($directory)
Add-Type 'using System;using System.Runtime.InteropServices;public static class NDMTestEndBookOwner63{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
$existing=@(Get-Process EXCEL -ErrorAction SilentlyContinue|ForEach-Object Id)
$app=$null;$control=$null;$book=$null;$own=0;$ticks='';$quit=$false;$readPassed=$false
try{
    $app=New-Object -ComObject Excel.Application
    [uint32]$id=0;[void][NDMTestEndBookOwner63]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id)
    if($existing -contains [int]$id){throw 'Existing user Excel refused'}
    $own=[int]$id;$ticks=[string](Get-Process -Id $own).StartTime.ToUniversalTime().Ticks
    [ordered]@{PID=$own;StartTicksUTC=$ticks;Label=$Label;WorkbookSHA256=$item.PreparedSHA256;AtUTC=[DateTimeOffset]::UtcNow.ToString('o')}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $directory 'Own.json') -Encoding UTF8
    $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=3;$app.ScreenUpdating=$false
    $control=$app.Workbooks.Add(-4167) # Пустая собственная книга удерживает ручной режим до открытия файла с пользовательскими данными.
    $app.Calculation=-4135;$app.CalculateBeforeSave=$false
    $book=$app.Workbooks.Open($source,0,$true)
    if(-not $book.ReadOnly -or $app.Calculation -ne -4135 -or $app.EnableEvents -or $app.AutomationSecurity -ne 3){throw 'Readonly/manual/macros-disabled contract required'}
    $export=New-Object Text.StringBuilder
    [void]$export.AppendLine('VBA PROJECT EXPORT');[void]$export.AppendLine('Workbook: '+$book.Name);[void]$export.AppendLine('Date: '+(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
    for($i=1;$i -le $book.VBProject.VBComponents.Count;$i++){
        $component=$book.VBProject.VBComponents.Item($i);$count=$component.CodeModule.CountOfLines
        [void]$export.AppendLine('='*100);[void]$export.AppendLine('COMPONENT: '+$component.Name);[void]$export.AppendLine('TYPE: '+$component.Type);[void]$export.AppendLine('LINES: '+$count);[void]$export.AppendLine('='*100)
        if($count){[void]$export.AppendLine($component.CodeModule.Lines(1,$count))};[void]$export.AppendLine()
    }
    if($book.VBProject.VBComponents.Count -ne 187){throw 'Saved181+6 components required'}
    [IO.File]::WriteAllText((Join-Path $directory 'VBA_All_Code.txt'),$export.ToString(),[Text.UTF8Encoding]::new($true))
    $geometry=$book.Names.Item('rngNDMSectionGeometry').RefersToRange.Cells.Item(1,1)
    $materialColumn=0;for($column=1;$column -le 20;$column++){if([string]$geometry.Cells.Item(1,$column).Value2 -ceq 'MaterialType'){$materialColumn=$column;break}}
    if($materialColumn -eq 0){throw 'Stored geometry MaterialType header missing'}
    $concrete=0;$bars=0
    if([string]$geometry.Cells.Item(2,1).Value2){
        $last=$geometry.End(-4121).Row;$rowCount=$last-$geometry.Row
        if($rowCount -le 0 -or $rowCount -gt 1000000){throw 'Stored geometry extent invalid'}
        $materials=$geometry.Offset(1,$materialColumn-1).Resize($rowCount,1).Value2
        for($row=1;$row -le $rowCount;$row++){
            $value=if($rowCount -eq 1){[string]$materials}else{[string]$materials[$row,1]}
            if($value -ceq 'Concrete'){$concrete++}elseif($value -ceq 'Rebar'){$bars++}else{throw ('Unknown stored material type '+$value)}
        }
    }
    $loads=$book.Names.Item('rngLoadCombinations').RefersToRange.Value2;$ids=@()
    for($row=2;$row -le $loads.GetLength(0);$row++){if([string]$loads[$row,1]){$ids+=,[string]$loads[$row,1]}}
    $summary=$book.Names.Item('rngBatchSummary').RefersToRange.Cells.Item(1,1).Resize(42,4).Value2
    $inputErrors=@();for($row=13;$row -le 42;$row++){if([string]$summary[$row,1] -and [string]$summary[$row,4] -ceq 'InputErr'){$inputErrors+=,[string]$summary[$row,1]}}
    $index=$book.Names.Item('rngNDMElementResultIndex').RefersToRange.Cells.Item(1,1)
    $headers=@('IndexKind','RunID','TableRows','TableColumns','FirstRow','RowCount','LoadCase','ProfileId','StateType')
    for($column=1;$column -le 9;$column++){if([string]$index.Cells.Item(1,$column).Value2 -cne $headers[$column-1]){throw 'Saved compact index header mismatch'}}
    $tableRows=$index.Cells.Item(2,3).Value2;$tableColumns=$index.Cells.Item(2,4).Value2;$states=0
    if([string]$index.Cells.Item(3,1).Value2){$last=$index.End(-4121).Row;$kinds=$index.Offset(2,0).Resize($last-$index.Row-1,1).Value2;$count=$last-$index.Row-1
        for($row=1;$row -le $count;$row++){$kind=if($count -eq 1){[string]$kinds}else{[string]$kinds[$row,1]};if($kind -cne 'Block'){throw 'Unknown saved index row kind'};$states++}}
    if($Label -ceq 'Heavy'){
        $expectedIDs=@(1..30|ForEach-Object{'PERF_'+$_.ToString('00')})
        if($concrete -ne 1232 -or $bars -ne 8 -or ($ids -join '|') -cne ($expectedIDs -join '|') -or ($inputErrors -join '|') -cne 'PERF_14|PERF_21|PERF_27|PERF_28' -or $tableRows -ne 114081 -or $tableColumns -ne 9 -or $states -ne 92){throw 'Actual saved heavy1232/8/30/4InputErr/92states/114080payload contract differs; preserve and review current data'}
    }
    [ordered]@{Status='PASS_READONLY_SAVED_CODE_AND_DATA';Label=$Label;WorkbookSHA256=$item.PreparedSHA256;Components=187;ReadOnly=$true;ManualCalculation=$true;MacrosDisabled=$true;EventsDisabled=$true;NoVBAMacroRun=$true;Saved=$false;NDSRecalculated=$false;Concrete=$concrete;Bars=$bars;LoadIDs=$ids;HistoricalInputErrIDs=$inputErrors;IndexTableRows=$tableRows;IndexTableColumns=$tableColumns;SavedStateBlocks=$states;RAMBeforeGiB=$ram;CommitBeforeGiB=$commit}|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $directory 'Native.json') -Encoding UTF8
    $readPassed=$true
}finally{
    if($book){$book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}
    if($control){$control.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($control)}
    $component=$null;$geometry=$null;$index=$null;[GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($app){if($own -gt 0){$process=Get-Process -Id $own -ErrorAction Stop;if($process.ProcessName -cne 'EXCEL' -or [string]$process.StartTime.ToUniversalTime().Ticks -cne $ticks){throw 'Own native identity changed before Quit'};$app.Quit();$quit=$true};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    [ordered]@{PID=$own;StartTicksUTC=$ticks;QuitReturned=$quit;ReadPassed=$readPassed;AtUTC=[DateTimeOffset]::UtcNow.ToString('o')}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $directory 'Quit.json') -Encoding UTF8
}
if((Get-FileHash -LiteralPath $source).Hash -cne $item.PreparedSHA256){throw 'Readonly inspection changed candidate'}
$process=Get-Process -Id $own -ErrorAction SilentlyContinue
if($process -and $process.ProcessName -ieq 'EXCEL' -and [string]$process.StartTime.ToUniversalTime().Ticks -ceq $ticks){[void]$process.WaitForExit(10000)}
Test-AssertRun02OwnTerminal60 ([pscustomobject]@{Name='EXCEL';PID=$own;StartTicksUTC=$ticks})
[ordered]@{PID=$own;StartTicksUTC=$ticks;ExactIdentityTerminal=$true;OrdinaryQuit=$quit;ForceKilled=$false;AtUTC=[DateTimeOffset]::UtcNow.ToString('o')}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $directory 'OwnClosure.json') -Encoding UTF8
& 'C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe' -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'TestVerifySource18.ps1') -ExportPath (Join-Path $directory 'VBA_All_Code.txt') -ReportPath (Join-Path $directory 'SourceEquality.json')
if($LASTEXITCODE -ne 0){throw 'Native source equality process failed'}
$equality=Get-Content -LiteralPath (Join-Path $directory 'SourceEquality.json') -Raw -Encoding UTF8|ConvertFrom-Json
if(-not $equality.Passed -or $equality.SourceComponents -ne 181 -or $equality.ExportComponents -ne 187 -or @($equality.Differences).Count -or @($equality.Missing).Count){throw 'Exact saved181+6 source bodies required'}
[ordered]@{Status='SAVED_USER_BOOK_CODE_AND_DATA_VERIFIED';Label=$Label;PreparedSHA256=$item.PreparedSHA256;SourceAcceptedSHA256=$pass.SourceSHA256;NativeProofSHA256=(Get-FileHash -LiteralPath (Join-Path $directory 'Native.json')).Hash;SourceEqualitySHA256=(Get-FileHash -LiteralPath (Join-Path $directory 'SourceEquality.json')).Hash;OwnClosureSHA256=(Get-FileHash -LiteralPath (Join-Path $directory 'OwnClosure.json')).Hash;OwnIdentityTerminal=$true;AllOtherPartsByteExact=$true;NDSRecalculated=$false;FullPassAppliesOnlyToAcceptedSource=$true}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $directory 'Accepted.json') -Encoding UTF8
Write-Output ('USER_BOOK_SAVED_READONLY_NATIVE_CODE_DATA_PASS: '+$Label)
