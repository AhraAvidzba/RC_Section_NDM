# ==================== ДЛЯ ТЕСТОВ ====================
# Разрешает завершающую публикацию только после exact Current_Run02 PASS, всех первичных журналов и фактического выхода own.
# Ничего не запускает, не закрывает и не меняет в книге; незавершённый полный набор отвергает первым условием.
$ErrorActionPreference='Stop'
Import-Module C:/Windows/System32/WindowsPowerShell/v1.0/Modules/Microsoft.PowerShell.Utility/Microsoft.PowerShell.Utility.psd1 -Global

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Не считает Unknown отсутствием; повторное использование PID допускает только при другом времени или независимо подтверждённом executable.
function Test-AssertRun02OwnTerminal60([object]$own){
    $process=$null
    try{$process=Get-Process -Id ([int]$own.PID) -ErrorAction Stop}catch{
        if($_.FullyQualifiedErrorId -match '^NoProcessFoundForGivenId(,|$)'){return}
        throw
    }
    if([string]::IsNullOrWhiteSpace([string]$process.ProcessName)){throw 'Own identity is unknown'}
    if($process.ProcessName -ine [string]$own.Name){
        $native=Get-CimInstance Win32_Process -Filter ('ProcessId='+[int]$own.PID) -ErrorAction Stop
        if($null -eq $native -or $null -eq $native.CreationDate -or [IO.Path]::GetFileNameWithoutExtension([string]$native.Name) -ine $process.ProcessName){throw 'Independent PID reuse proof required'}
        return
    }
    if($null -eq $process.StartTime){throw 'Own UTC start time unavailable'}
    if([string]$process.StartTime.ToUniversalTime().Ticks -ceq [string]$own.StartTicksUTC){throw ('Own identity remains live: '+$own.PID+'/'+$own.StartTicksUTC)}
}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Проверяет весь состав 182+122/26 и source fence. Task cleanup обязателен для публикации, но не для отдельного просмотра PASS.
function Test-AssertCurrentRun02Pass60([switch]$RequireTaskRemoved,[switch]$RequireAdditionalIO){
    $root='C:/Users/avidzba/Desktop/программы/В_РАБОТЕ!/RC_Section_NDM'
    $launch=Join-Path $root 'docs/regression/FullSuiteLaunch_2026-11-01_Current_Run02'
    if(-not (Test-Path -LiteralPath (Join-Path $launch 'Completion.json'))){throw 'CURRENT_RUN02_STILL_RUNNING_END_ONLY_ACTIONS_FORBIDDEN'}
    $read={param($path) Get-Content -LiteralPath $path -Raw -Encoding UTF8|ConvertFrom-Json}
    $request=& $read (Join-Path $launch 'LaunchRequest.json')
    $full=& $read (Join-Path $launch 'FullLaunch.json')
    $completion=& $read (Join-Path $launch 'Completion.json')
    $source='831D725C06439CA1B771645F18B0D546169C1E22A97BA7688D64CB3F8509B74A'
    $commit='5f18de1f88a657ae312bc37b0798bc511f4bb218'
    if($full.PID -ne 98996 -or [string]$full.StartTicks -cne '639291636492982491' -or $full.HostPID -ne 102444 -or [string]$full.HostStartTicks -cne '639291636487808538'){throw 'Foreign driver/host epoch refused'}
    if((Get-FileHash -LiteralPath (Join-Path $launch 'LaunchRequest.json')).Hash -cne $full.RequestSHA256){throw 'Launch request changed'}
    foreach($proof in @($request,$full,$completion)){
        if($proof.SourceSHA256 -cne $source -or $proof.SourceCommit -cne $commit){throw 'Exact accepted source831D/commit5f18 required'}
    }
    if($completion.Status -cne 'PASS_PENDING_PRIMARY_LOG_REVIEW' -or -not $completion.FinishedUTC -or $completion.RunnerExit -ne 0 -or $completion.SummaryExit -ne 0 -or $completion.FinalizeExit -ne 0 -or -not $completion.CompleteAcceptancePassed -or $completion.DriverError -or @($completion.ChangedInputs).Count -or $completion.HeadChanged){throw 'Whole successful terminal completion required'}
    $evidence=Join-Path $root 'docs/regression/FullOnOff_2026-11-01_Current_Run02'
    if([IO.Path]::GetFullPath((Join-Path $root $request.EvidenceDirectory)) -ine [IO.Path]::GetFullPath($evidence)){throw 'Foreign evidence directory'}
    $complete=& $read (Join-Path $evidence 'CompleteCurrentAcceptance.json')
    $combined=& $read (Join-Path $evidence 'Focused/CombinedAcceptance.json')
    foreach($proof in @($complete,$combined)){
        if($proof.Status -cne 'PASS' -or $proof.SourceSHA256 -cne $source -or $proof.CoreSuites -ne 182 -or $proof.FocusedSingleCases -ne 122 -or -not $proof.SameWorkbookBytes -or -not $proof.SingleInvocation -or $proof.IndependentlyReviewedNumFailStates -ne 10){throw 'Official exact-source complete182+122 acceptance required'}
        if($proof.FocusedRunSHA256 -cne (Get-FileHash -LiteralPath (Join-Path $evidence 'Focused/Run.json')).Hash -or $proof.IndependentSHA256 -cne (Get-FileHash -LiteralPath (Join-Path $evidence 'Focused/Independent.json')).Hash){throw 'Official focused/independent proofs changed'}
    }
    $scopePath=Join-Path $launch 'TerminalScopeReview.json'
    $scope=& $read $scopePath
    if($scope.SourceSHA256 -cne $source -or $scope.SourceCommit -cne $commit -or -not $scope.PrimaryLogsReviewed -or $scope.CoreReportsRead -ne 12 -or $scope.CoreMacrosReviewed -ne 182 -or $scope.FocusedRecordCount -ne 26 -or $scope.FocusedSingleCasesReviewed -ne 122 -or @($scope.InventoryErrors).Count -or @($scope.Failures).Count -or @($scope.ChangedInputs).Count -or -not $scope.HEADUnchanged){throw 'Whole raw primary review without failures required'}
    $roles=@{};foreach($item in $scope.Primary){
        if(-not $item.AllBytesRead -or -not $item.Closed -or $item.SourceMatches -eq $false -or (Get-FileHash -LiteralPath (Join-Path $root $item.Path)).Hash -cne $item.SHA256){throw 'Raw primary changed or was not fully read'}
        if(-not $roles.ContainsKey([string]$item.Role)){$roles[[string]$item.Role]=0};$roles[[string]$item.Role]++
    }
    if($roles['Core'] -ne 12 -or $roles['Focused'] -ne 24 -or $roles['FocusedJSON'] -ne 2 -or $roles['NativeCADChild'] -ne 2){throw 'Entire raw core/focused/independent/two-real-CAD composition required'}
    $focus=& $read (Join-Path $evidence 'Focused/Run.json')
    if($focus.SourceSHA256 -cne $source -or @($focus.Records).Count -ne 26 -or -not $focus.ProtectedInputsUnchanged){throw 'Complete focused run with original source required'}
    foreach($record in $focus.Records){if(-not $record.Passed -or ($record.PSObject.Properties['ExitCode'] -and $record.ExitCode -ne 0)){throw 'Unsuccessful focused record'}}
    $independent=& $read (Join-Path $evidence 'Focused/Independent.json')
    if(-not $independent.Passed -or $independent.SourceSHA256 -cne $source -or $independent.ModelCount -ne 5 -or $independent.IndependentIntegrals -ne 560 -or $independent.ReviewedNumFailStates -ne 10 -or @($independent.Unsupported).Count -or @($independent.Failures).Count -ne 10){throw 'Entire independent5models/560integrals/10NumFail contract required'}
    foreach($item in $independent.Failures){if(-not $item.Passed){throw 'An individual expected NumFail check did not pass'}}
    $ownPath=Join-Path $launch 'AllOwnClosuresReviewed.json'
    $own=& $read $ownPath
    if(-not $own.AllOwnIdentitiesTerminal -or -not $own.PrimaryLogsReviewed -or -not $own.InputsUnchanged -or $own.InputsChecked -ne 522 -or $own.SourceSHA256 -cne $source -or $own.SourceCommit -cne $commit -or $own.ScopeReviewSHA256 -cne (Get-FileHash -LiteralPath $scopePath).Hash){throw 'Exact all-own and whole primary proof required'}
    foreach($identity in $own.Identities){Test-AssertRun02OwnTerminal60 $identity}
    $stdio=& $read (Join-Path $PSScriptRoot 'TerminalStdIOReviewed68.json')
    if($stdio.Status -cne 'ENTIRE_TERMINAL_STDIO_READ_NO_RUNTIME_ERRORS' -or $stdio.SourceSHA256 -cne $source -or $stdio.SourceCommit -cne $commit -or -not $stdio.AllOwnIdentitiesTerminal -or @($stdio.Files).Count -ne 2 -or @($stdio.Errors).Count -or $stdio.DecodedFallback){throw 'Entire current terminal stdout/stderr review required'}
    $names=@();foreach($file in $stdio.Files){if($file.Name -cnotin @('FullLaunch.stdout.log','FullLaunch.stderr.log') -or -not $file.AllBytesRead -or @($file.Errors).Count -or (Get-FileHash -LiteralPath (Join-Path $launch $file.Name)).Hash -cne $file.SHA256){throw 'Terminal stdio changed or incomplete'};$names+=,[string]$file.Name}
    if(@($names|Select-Object -Unique).Count -ne 2){throw 'Both distinct current terminal stdio files required'}
    $rows=@($request.ProtectedInputs)+@($request.SourceComponents)+@($request.ParticipatingScripts)
    if($rows.Count -ne 522 -or (& git -C $root rev-parse HEAD) -cne $commit){throw 'Original full input count/HEAD differs'}
    foreach($item in $rows){if((Get-FileHash -LiteralPath (Join-Path $root $item.Path)).Hash -cne $item.SHA256){throw ('Original full input changed: '+$item.Path)}}
    $accepted=Join-Path $root $request.SourceWorkbook
    $frozen=Join-Path $evidence 'Source/RC_Section_NDM.xlsm'
    if((Get-FileHash -LiteralPath $accepted).Hash -cne $source -or (Get-FileHash -LiteralPath $frozen).Hash -cne $source){throw 'Accepted/frozen source bytes differ'}
    if($RequireTaskRemoved){
        $cleanup=& $read (Join-Path $launch 'ScheduledTaskCleanup61.json')
        if($cleanup.TaskName -cne 'RC_NDM_Full_20261101_Current_Run02' -or $cleanup.SourceSHA256 -cne $source -or $cleanup.SourceCommit -cne $commit -or -not $cleanup.AllOwnIdentitiesTerminal){throw 'Exact own task cleanup proof required'}
        if(@(Get-ScheduledTask -TaskPath '\' -ErrorAction Stop|Where-Object TaskName -eq 'RC_NDM_Full_20261101_Current_Run02').Count){throw 'Own task remains registered'}
    }
    if($RequireAdditionalIO){
        $io=& $read (Join-Path $PSScriptRoot 'AdditionalIO74.json')
        $plan=& $read (Join-Path $PSScriptRoot 'CompletionRequirementMap73.json')
        $expected=@($plan.AdditionalIO.Macros)
        if($io.Status -cne 'ADDITIONAL_IO_NINE_ENTRYPOINTS_BOTH_MODES_PASS' -or $io.SourceSHA256 -cne $source -or $io.SourceCommit -cne $commit -or $expected.Count -ne 9 -or @($io.Completed).Count -ne 2 -or (@($io.Completed|ForEach-Object Mode) -join '|') -cne 'No|Yes'){throw 'All explicit additional I/O9x2 checks required before publication'}
        Test-AssertRun02OwnTerminal60 ([pscustomobject]@{Name='powershell';PID=$io.ControllerPID;StartTicksUTC=$io.ControllerStartTicksUTC})
        foreach($entry in $io.Completed){
            if(-not $entry.Passed -or -not $entry.AllBytesRead -or -not $entry.OwnIdentityTerminal -or ($entry.Started -join '|') -cne ($expected -join '|') -or ($entry.Finished -join '|') -cne ($expected -join '|') -or (@($entry.ExplicitExcelSetups|ForEach-Object Macro) -join '|') -cne ($expected -join '|') -or @($entry.ExplicitExcelSetups|Where-Object Mode -cne $entry.Mode).Count -or @($entry.Errors).Count -or (Get-FileHash -LiteralPath (Join-Path $root $entry.Primary)).Hash -cne $entry.SHA256){throw 'Additional I/O primary/setup/source coverage differs'}
            Test-AssertRun02OwnTerminal60 $entry.NativeOwner
        }
    }
    [pscustomobject]@{Root=$root;Launch=$launch;Evidence=$evidence;Request=$request;Source=$accepted;SourceSHA256=$source;SourceCommit=$commit;InputsChecked=522;CoreSuites=182;FocusedSingles=122;FocusedRecords=26;OwnCount=@($own.Identities).Count;AllOwnIdentitiesTerminal=$true;ScopeReviewSHA256=(Get-FileHash -LiteralPath $scopePath).Hash;OwnProofSHA256=(Get-FileHash -LiteralPath $ownPath).Hash;OfficialPASS=$true}
}
