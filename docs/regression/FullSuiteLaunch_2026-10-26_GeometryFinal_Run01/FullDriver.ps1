# ==========================================================================
# ДЛЯ ТЕСТОВ: полный авторизованный запуск текущей книги
# ==========================================================================
# Не меняет участвующие исходники, опубликованную и замороженную книги.
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
Import-Module (Join-Path $PSHOME 'Modules/Microsoft.PowerShell.Utility/Microsoft.PowerShell.Utility.psd1') -Global -ErrorAction Stop
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path
Set-Location -LiteralPath $root
$request=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'LaunchRequest.json') -Raw -Encoding UTF8|ConvertFrom-Json
$driver=Get-Process -Id $PID
[ordered]@{Status='RUNNING';PID=$PID;StartTicks=[string]$driver.StartTime.ToUniversalTime().Ticks;StartedUTC=[DateTimeOffset]::UtcNow.ToString('o');SourceCommit=$request.SourceCommit;SourceSHA256=$request.SourceSHA256}|
    ConvertTo-Json|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'DriverRun.json') -Encoding UTF8
$runnerExit=999;$summaryExit=999;$finalExit=999;$driverError='';$changed=@();$passed=$false
try {
    if((& git rev-parse HEAD) -cne $request.SourceCommit){throw 'Source commit changed before launch'}
    if([int]$request.ExpectedSuites -ne 178 -or [int]$request.CoverageVersion -ne 6 -or -not $request.IncludesFocusedAcceptance){throw 'Full GeometryFinal coverage request differs.'}
    $memory=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
    if([double]$memory.AvailableMBytes/1024 -lt 8 -or ([double]$memory.CommitLimit-[double]$memory.CommittedBytes)/1GB -lt 14){throw 'Insufficient actual8RAM/14commit before full core process; no Excel started.'}
    foreach($item in @($request.ProtectedInputs)+@($request.ParticipatingScripts)+@($request.SourceComponents)){
        $path=Join-Path $root $item.Path
        if(-not(Test-Path -LiteralPath $path) -or (Get-FileHash -LiteralPath $path).Hash -cne $item.SHA256){throw ('Input changed before launch: '+$item.Path)}
    }
    & (Join-Path $PSHOME 'powershell.exe') -NoProfile -ExecutionPolicy Bypass -File tools/build_workbook/Run-FullOnOffAcceptance.ps1 -SourceWorkbook $request.SourceWorkbook -EvidenceDirectory $request.EvidenceDirectory -Deadline $request.Deadline -ForecastHours $request.ForecastHours -Python $request.Python
    $runnerExit=$LASTEXITCODE
    & $request.Python tools/build_workbook/Summarize-FullOnOffAcceptance.py --directory $request.EvidenceDirectory --check
    $summaryExit=$LASTEXITCODE
    & $request.Python tools/build_workbook/Finalize-CompleteAcceptance.py --core-directory $request.EvidenceDirectory --focused-directory ($request.EvidenceDirectory+'/Focused')
    $finalExit=$LASTEXITCODE
} catch {
    $driverError=$_.Exception.Message
    Write-Output ('FULL_DRIVER_ERROR: '+$driverError)
} finally {
    foreach($item in @($request.ProtectedInputs)+@($request.ParticipatingScripts)+@($request.SourceComponents)){
        $path=Join-Path $root $item.Path
        if(-not(Test-Path -LiteralPath $path) -or (Get-FileHash -LiteralPath $path).Hash -cne $item.SHA256){$changed+=$item.Path}
    }
    $headChanged=(& git rev-parse HEAD) -cne $request.SourceCommit
    $completePath=Join-Path $root ($request.EvidenceDirectory+'/CompleteCurrentAcceptance.json')
    $completePass=$false
    if(Test-Path -LiteralPath $completePath){
        $complete=Get-Content -LiteralPath $completePath -Raw -Encoding UTF8|ConvertFrom-Json
        $completePass=$complete.Status -eq 'PASS' -and $complete.SourceSHA256 -ceq $request.SourceSHA256 -and [int]$complete.CoreSuites -eq 178 -and [int]$complete.FocusedSingleCases -eq 122
    }
    $passed=$runnerExit -eq 0 -and $summaryExit -eq 0 -and $finalExit -eq 0 -and $completePass -and $changed.Count -eq 0 -and -not $headChanged -and -not $driverError
    [ordered]@{Status=if($passed){'PASS_PENDING_PRIMARY_LOG_REVIEW'}else{'REVIEW_REQUIRED'};FinishedUTC=[DateTimeOffset]::UtcNow.ToString('o');SourceSHA256=$request.SourceSHA256;SourceCommit=$request.SourceCommit;RunnerExit=$runnerExit;SummaryExit=$summaryExit;FinalizeExit=$finalExit;CompleteAcceptancePassed=$completePass;ExpectedSuites=178;CoverageVersion=6;PrimaryReviewStillRequired=$true;ChangedInputs=$changed;HeadChanged=$headChanged;DriverError=$driverError;EvidenceDirectory=$request.EvidenceDirectory}|
        ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'Completion.json') -Encoding UTF8
}
if(-not $passed){exit 1}
