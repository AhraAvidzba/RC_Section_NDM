# ==================== ДЛЯ ТЕСТОВ ====================
# После полного терминала/PASS/all-own выполняет явные девять I/O entrypoint из ValidationPlan в No/Yes.
# Отдельный адресный состав не заменяет полные182+Focused122; в каждом suite собственная копия начинает с Excel.
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'TestAssertCurrentRun02Pass60.ps1')
$pass=Test-AssertCurrentRun02Pass60 -RequireTaskRemoved
$prepared=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'CompletionRequirementMap73.json') -Raw -Encoding UTF8|ConvertFrom-Json
if($prepared.SourceAcceptedSHA256 -cne $pass.SourceSHA256 -or $prepared.SourceCommit -cne $pass.SourceCommit -or $prepared.AdditionalIO.ExpectedMacroExecutions -ne 18 -or -not $prepared.AdditionalIO.ExplicitExcelBeforeEveryMacro){throw 'Exact additional I/O plan required'}
foreach($item in @([pscustomobject]@{Name='TestIOModeRunner73.ps1';Hash=$prepared.PrivateRunnerSHA256},[pscustomobject]@{Name='TestIOWatchdog73.ps1';Hash=$prepared.PrivateWatchdogSHA256})){
    if((Get-FileHash -LiteralPath (Join-Path $PSScriptRoot $item.Name)).Hash -cne $item.Hash){throw 'Prepared private runner changed'}
}
$macros=@($prepared.AdditionalIO.Macros)
if($macros.Count -ne 9){throw 'Nine exact I/O entrypoints required'}
$path=Join-Path $PSScriptRoot 'AdditionalIO74.json'
if(Test-Path -LiteralPath $path){throw 'Additional I/O run exists; inspect rather than duplicate'}
$directory=Join-Path $pass.Root 'docs/regression/FullAcceptanceRecovery_2026-11-01_Run01/AdditionalIO74'
if(Test-Path -LiteralPath $directory){throw 'Fresh additional I/O directory required'}
$state=[ordered]@{Status='RUNNING';StartedUTC=[DateTimeOffset]::UtcNow.ToString('o');ControllerPID=$PID;ControllerStartTicksUTC=[string](Get-Process -Id $PID).StartTime.ToUniversalTime().Ticks;SourceSHA256=$pass.SourceSHA256;SourceCommit=$pass.SourceCommit;ExpectedMacros=$macros;Current='';Completed=@();FullAcceptanceClaim=$false}
try{
    foreach($mode in @('No','Yes')){
        # Резерв всей оставшейся адресной матрицы и финала проверяется заново; это не разрешение нового полного набора.
        $remaining=if($mode -ceq 'No'){30}else{15}
        if([DateTimeOffset]::Now.AddMinutes($remaining+45) -gt [DateTimeOffset]::Parse('2026-11-02T10:00:00+03:00')){throw 'Additional I/O and final review no longer fit the original user deadline'}
        $fresh=Test-AssertCurrentRun02Pass60 -RequireTaskRemoved
        foreach($previous in $state.Completed){Test-AssertRun02OwnTerminal60 $previous.NativeOwner}
        $memory=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
        if($null -eq $memory -or $null -eq $memory.AvailableMBytes -or $null -eq $memory.CommitLimit -or $null -eq $memory.CommittedBytes){throw 'Actual memory counters unavailable'}
        $ram=[double]$memory.AvailableMBytes/1024;$commit=([double]$memory.CommitLimit-[double]$memory.CommittedBytes)/1GB
        if($ram -lt 8 -or $commit -lt 14){throw 'Native reserve below8RAM/14commit'}
        $state.Current=$mode;$state|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $path -Encoding UTF8
        $relative='docs/regression/FullAcceptanceRecovery_2026-11-01_Run01/AdditionalIO74/IO_'+$mode+'.txt'
        $primary=Join-Path $pass.Root $relative
        Set-Location -LiteralPath $pass.Root
        & 'C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe' -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'TestIOWatchdog73.ps1') -SourceWorkbook $pass.Source -ReportPath $relative -Macro ($macros -join ',') -Mode $mode -TimeoutSeconds 900
        $code=$LASTEXITCODE
        $raw=[IO.File]::ReadAllBytes($primary);$text=[Text.Encoding]::UTF8.GetString($raw)
        $started=@([regex]::Matches($text,'(?m)^SUITE_STARTED: ([^;\r\n]+)')|ForEach-Object{$_.Groups[1].Value})
        $finished=@([regex]::Matches($text,'(?m)^SUITE_FINISHED: ([^;\r\n]+)')|ForEach-Object{$_.Groups[1].Value})
        $setup=@([regex]::Matches($text,'(?m)^EXPLICIT_IO_FIXTURE: Results\.SnapshotStorage=Excel; macro=([^;\r\n]+); mode=([^\r\n]+)')|ForEach-Object{[pscustomobject]@{Macro=$_.Groups[1].Value;Mode=$_.Groups[2].Value}})
        $errors=@([regex]::Matches($text,'(?im)^(?:FAIL:|RUNTIME ERROR:|ERROR:|WATCHDOG_FAILURE:|CLEANUP ERROR:)[^\r\n]*')|ForEach-Object{$_.Value})
        $nonzero=@([regex]::Matches($text,'(?i)failed=(\d+)')|Where-Object{[int]$_.Groups[1].Value -gt 0})
        $identity=@([regex]::Matches($text,'(?m)^TEST_EXCEL_PROCESS: id=(\d+); startTicks=(\d+)'))
        if($identity.Count -ne 1){throw 'Actual additional I/O own identity missing'}
        $owner=[pscustomobject]@{Name='EXCEL';PID=[int]$identity[0].Groups[1].Value;StartTicksUTC=$identity[0].Groups[2].Value}
        Test-AssertRun02OwnTerminal60 $owner
        $closure=Get-Content -LiteralPath ($primary+'.ProcessClosure.json') -Raw -Encoding UTF8|ConvertFrom-Json
        $passed=$code -eq 0 -and $closure.Passed -and $closure.ChildHasExited -and $errors.Count -eq 0 -and $nonzero.Count -eq 0 -and ($started -join '|') -ceq ($macros -join '|') -and ($finished -join '|') -ceq ($macros -join '|') -and (@($setup|ForEach-Object Macro) -join '|') -ceq ($macros -join '|') -and @($setup|Where-Object Mode -cne $mode).Count -eq 0 -and $text.Contains('SOURCE_SHA256: '+$pass.SourceSHA256) -and $text.Contains('WATCHDOG_COMPLETED: exit=0; source unchanged=True')
        $entry=[ordered]@{Mode=$mode;Passed=$passed;Primary=$relative;Bytes=$raw.Length;SHA256=(Get-FileHash -LiteralPath $primary).Hash;AllBytesRead=$true;Expected=$macros;Started=$started;Finished=$finished;ExplicitExcelSetups=$setup;Errors=$errors;Totals=@([regex]::Matches($text,'(?m)^TOTAL[^\r\n]*')|ForEach-Object{$_.Value});NativeOwner=$owner;OwnIdentityTerminal=$true;ProcessClosure=$closure;ExitCode=$code;RAMBeforeGiB=$ram;CommitBeforeGiB=$commit}
        $entry|ConvertTo-Json -Depth 10|Set-Content -LiteralPath ($primary+'.PrimaryReview.json') -Encoding UTF8
        $state.Completed+=,$entry;$state|ConvertTo-Json -Depth 11|Set-Content -LiteralPath $path -Encoding UTF8
    }
    $after=Test-AssertCurrentRun02Pass60 -RequireTaskRemoved
    $state.Status=if(@($state.Completed).Count -eq 2 -and @($state.Completed|Where-Object {-not $_.Passed}).Count -eq 0){'ADDITIONAL_IO_NINE_ENTRYPOINTS_BOTH_MODES_PASS'}else{'REVIEW_REQUIRED'}
    $state.Current='';$state['FinishedUTC']=[DateTimeOffset]::UtcNow.ToString('o')
}catch{$state.Status='REVIEW_REQUIRED';$state['Error']=$_.Exception.ToString();throw}finally{$state|ConvertTo-Json -Depth 12|Set-Content -LiteralPath $path -Encoding UTF8}
if($state.Status -cne 'ADDITIONAL_IO_NINE_ENTRYPOINTS_BOTH_MODES_PASS'){throw 'Additional I/O matrix requires fixes; original reports preserved'}
Write-Output 'ADDITIONAL_IO_NINE_ENTRYPOINTS_BOTH_MODES_PASS'
