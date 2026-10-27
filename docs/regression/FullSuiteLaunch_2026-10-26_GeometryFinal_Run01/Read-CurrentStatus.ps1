# ==========================================================================
# ДЛЯ ТЕСТОВ: чтение фактического состояния запущенной полной приёмки
# ==========================================================================
# Не меняет участники прогона; сохраняет только собственную запись наблюдения.
param([switch]$CheckAllHashes)
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSHOME 'Modules/Microsoft.PowerShell.Utility/Microsoft.PowerShell.Utility.psd1') -Global -ErrorAction Stop
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path
Set-Location -LiteralPath $root
$launch=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'FullLaunch.json') -Raw -Encoding UTF8|ConvertFrom-Json
$requestPath=Join-Path $PSScriptRoot 'LaunchRequest.json'
if((Get-FileHash -LiteralPath $requestPath).Hash -cne $launch.RequestSHA256){throw 'Immutable launch request changed'}
$request=Get-Content -LiteralPath $requestPath -Raw -Encoding UTF8|ConvertFrom-Json
$directory=Join-Path $root $request.EvidenceDirectory
$run=Get-Content -LiteralPath (Join-Path $directory 'Run.json') -Raw -Encoding UTF8|ConvertFrom-Json
$now=[DateTimeOffset]::Now
$monitored=@()

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Не принимает повторно использованный PID за прежний принадлежащий тесту процесс.
function Test-ReadOwnedProcess([int]$ProcessId,[string]$ExpectedTicks){
    try {$process=Get-Process -Id $ProcessId -ErrorAction Stop}
    catch {
        if($_.FullyQualifiedErrorId -like 'NoProcessFoundForGivenId*'){return [ordered]@{PID=$ProcessId;Live=$false;IdentityMatches=$false;Observation='ConfirmedNotFound'}}
        throw ('Process identity is unknown for PID '+$ProcessId+': '+$_.Exception.Message)
    }
    $start=$process.StartTime
    if($null -eq $start){throw ('StartTime inaccessible for PID '+$ProcessId+'; absence is not proved; use the same privileges as the launch')}
    $actual=[string]$start.ToUniversalTime().Ticks
    return [ordered]@{PID=$ProcessId;Name=$process.ProcessName;Live=$true;StartTicks=$actual;IdentityMatches=($actual -ceq $ExpectedTicks);CPU=$process.CPU;PrivateBytes=$process.PrivateMemorySize64;Observation=$(if($actual -ceq $ExpectedTicks){'ExactIdentityPresent'}else{'DifferentIdentityPresent'})}
}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Разделяет чтение с писателем; временная эксклюзивная запись не означает остановку.
function Test-ReadProgressText([string]$Path,[Text.Encoding]$Encoding=[Text.Encoding]::Default){
    for($attempt=1;$attempt -le 4;$attempt++){
        $stream=$null;$reader=$null
        try{
            $share=[IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete
            $stream=[IO.FileStream]::new($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,$share)
            $reader=[IO.StreamReader]::new($stream,$Encoding,$true)
            return [ordered]@{Read=$true;Text=$reader.ReadToEnd();Attempts=$attempt}
        }catch [IO.IOException]{
            if($attempt -eq 4){return [ordered]@{Read=$false;Text='';Attempts=$attempt;Reason='Temporarily busy or closed by the running test; retry next observation'}}
        }finally{
            if($reader){$reader.Dispose()}elseif($stream){$stream.Dispose()}
        }
        Start-Sleep -Milliseconds 50
    }
}

$driver=Test-ReadOwnedProcess ([int]$launch.PID) ([string]$launch.StartTicks)
$hostIdentity=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'HostRun.json') -Raw -Encoding UTF8|ConvertFrom-Json
$independentHost=Test-ReadOwnedProcess ([int]$hostIdentity.PID) ([string]$hostIdentity.StartTicks)
$job=[string]$run.CurrentJob
$focusedPath=Join-Path $directory 'Focused/Run.json'
$focused=$null
if(Test-Path -LiteralPath $focusedPath){$focused=Get-Content -LiteralPath $focusedPath -Raw -Encoding UTF8|ConvertFrom-Json}
$logPath=Join-Path $directory ($job+'.txt')
if($focused -and $focused.Status -eq 'RUNNING'){
    $job='Focused'
    $latest=Get-ChildItem -LiteralPath (Join-Path $directory 'Focused') -Filter '*.txt' -File|Sort-Object LastWriteTimeUtc -Descending|Select-Object -First 1
    if($latest){$logPath=$latest.FullName}
}
$failures=@();$progressFailures=@();$completedSuites=@();$tail=@();$excel=$null;$progress=@()
if(Test-Path -LiteralPath $logPath){
    $log=Get-Item -LiteralPath $logPath
    $primaryRead=Test-ReadProgressText $logPath ([Text.Encoding]::UTF8)
    $body=$primaryRead.Text
    $monitored+=,$log
    $failures=@([regex]::Matches($body,'(?m)^(FAIL[: ][^\r\n]*|RUNTIME ERROR[^\r\n]*|SCRIPT ERROR[^\r\n]*|WATCHDOG_FAILURE[^\r\n]*|RESTORE ERROR[^\r\n]*|CLEANUP ERROR[^\r\n]*)')|ForEach-Object Value)
    $completedSuites=@([regex]::Matches($body,'(?m)^SUITE_FINISHED: [^\r\n]*')|ForEach-Object Value)
    $tail=@($body -split '\r?\n'|Where-Object {$_}|Select-Object -Last 6)
    $match=[regex]::Match($body,'TEST_EXCEL_PROCESS: id=(\d+); startTicks=(\d+)')
    if($match.Success){$excel=Test-ReadOwnedProcess ([int]$match.Groups[1].Value) $match.Groups[2].Value}
    $fixture=[regex]::Match($body,'(?m)^FIXTURE: ([^\r\n]+)')
    if($fixture.Success){
        $folder=Split-Path -Parent $fixture.Groups[1].Value
        $allowed=[IO.Path]::GetFullPath([IO.Path]::GetTempPath())
        if([IO.Path]::GetFullPath($folder).StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase) -and (Split-Path -Leaf $folder) -like 'RC_NDM_ModeSuite_*' -and (Test-Path -LiteralPath $folder)){
            foreach($file in Get-ChildItem -LiteralPath $folder -Filter '*Progress*' -File){
                $monitored+=,$file
                $read=Test-ReadProgressText $file.FullName # VBA Print# использует системную ANSI; BOM UTF-8/UTF-16 распознаётся StreamReader автоматически.
                $plainLines=if($read.Read){@($read.Text -split '\r?\n')}else{@()}
                foreach($failure in [regex]::Matches(($plainLines -join "`n"),'(?m)^(FAIL[: ][^\r\n]*|RUNTIME ERROR[^\r\n]*|SCRIPT ERROR[^\r\n]*)')){
                    $progressFailures+=,[ordered]@{File=[string]$file.Name;Message=[string]$failure.Value}
                }
                $progress+=,[ordered]@{Name=[string]$file.Name;Bytes=$file.Length;LastWrite=$file.LastWriteTime.ToString('o');ReadSucceeded=$read.Read;ReadAttempts=$read.Attempts;DefaultCodePage=[Text.Encoding]::Default.CodePage;Tail=@($plainLines|Where-Object {$_}|Select-Object -Last 3)}
            }
        }
    }
}
$changed=@();$checked=0
if($CheckAllHashes){
    foreach($entry in @($request.ProtectedInputs)+@($request.SourceComponents)+@($request.ParticipatingScripts)){
        $checked++
        $path=Join-Path $root $entry.Path
        if(-not(Test-Path -LiteralPath $path) -or (Get-FileHash -LiteralPath $path).Hash -cne $entry.SHA256){$changed+=$entry.Path}
    }
}
$focusedInfo=$null;$nativeProcesses=@();$focusedRecent=@()
if($focused){
    $focusedInfo=[ordered]@{Status=$focused.Status;StartedUTC=$focused.StartedUTC;DeadlineUTC=$focused.DeadlineUTC;CompletedRecords=@($focused.Records).Count;FailedRecords=@($focused.Records|Where-Object {-not $_.Passed}|ForEach-Object {[ordered]@{ID=$_.ID;Kind=$_.Kind;Report=$_.Report}});LastRecord=$(if(@($focused.Records).Count){$focused.Records[-1].ID}else{$null});SourceSHA256=$focused.SourceSHA256}
    $focusRoot=Join-Path $directory 'Focused'
    $recent=Get-ChildItem -LiteralPath $focusRoot -File -Recurse|Where-Object {$_.Extension -in '.txt','.log','.json'}|Sort-Object LastWriteTimeUtc -Descending|Select-Object -First 8
    foreach($file in $recent){
        $monitored+=,$file
        $focusedRecent+=,[ordered]@{Path=$file.FullName.Substring($focusRoot.Length+1);Bytes=$file.Length;LastWrite=$file.LastWriteTime.ToString('o')}
    }
    foreach($identityFile in Get-ChildItem -LiteralPath $focusRoot -Filter '*Ownership.json' -File -Recurse){
        $identityRead=Test-ReadProgressText $identityFile.FullName ([Text.Encoding]::UTF8)
        if($identityRead.Read -and $identityRead.Text.Trim()){
            try{
                $identity=$identityRead.Text|ConvertFrom-Json
                $nativeProcesses+=,[ordered]@{IdentityFile=$identityFile.FullName.Substring($focusRoot.Length+1);Process=(Test-ReadOwnedProcess ([int]$identity.OwnedPID) ([string]$identity.StartTicks))}
            }catch{ # Незавершённый JSON и ошибка чтения PID не доказывают отсутствие процесса.
                $nativeProcesses+=,[ordered]@{IdentityFile=$identityFile.FullName.Substring($focusRoot.Length+1);ReadingDeferred=$true;AbsenceProven=$false;ObservationError=$_.Exception.Message}
            }
        }
    }
}
$terminal=Join-Path $PSScriptRoot 'Completion.json'
$completion=if(Test-Path -LiteralPath $terminal){Get-Content -LiteralPath $terminal -Raw -Encoding UTF8|ConvertFrom-Json}else{$null}
$lastActivity=$monitored|Sort-Object LastWriteTimeUtc -Descending|Select-Object -First 1
$result=[ordered]@{Checked=[DateTimeOffset]::Now.ToString('o');ReadStarted=$now.ToString('o');Driver=$driver;CurrentJob=$job;RunStatus=$run.Status;CompletedJobs=@($run.CompletedJobs).Count;FailedCompletedJobs=@($run.CompletedJobs|Where-Object {$_.ExitCode -ne 0});ExpectedSuites=$request.ExpectedSuites;CurrentLog=$logPath;CompletedSuitesInCurrentLog=$completedSuites;ErrorsInCurrentLog=$failures;ErrorsInProgressFiles=$progressFailures;OwnedExcel=$excel;ProgressFiles=$progress;LastActivity=$(if($lastActivity){$lastActivity.LastWriteTime.ToString('o')}else{$null});LogTail=$tail;InputsChecked=$checked;ChangedInputs=$changed;HEADUnchanged=((& git rev-parse HEAD) -ceq $request.SourceCommit);FrozenSourceSHA256=(Get-FileHash -LiteralPath $run.FrozenWorkbook).Hash;Deadline=$launch.Deadline;Completion=$completion}
$memory=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
$result['Memory']=[ordered]@{FreePhysicalGiB=[math]::Round([double]$memory.AvailableMBytes/1024,3);FreeCommitGiB=[math]::Round(([double]$memory.CommitLimit-[double]$memory.CommittedBytes)/1GB,3);CommitLimitGiB=[math]::Round([double]$memory.CommitLimit/1GB,3);CommittedGiB=[math]::Round([double]$memory.CommittedBytes/1GB,3)}
$result['IndependentHost']=$independentHost;$result['PublishedSourceSHA256']=(Get-FileHash -LiteralPath (Join-Path $root $request.SourceWorkbook)).Hash;$result['AllCodeSHA256']=(Get-FileHash -LiteralPath (Join-Path $root 'workbook/output/VBA_All_Code.txt')).Hash;$result['RequestedAllCodeSHA256']=$request.AllCodeSHA256;$result['CoverageVersion']=$request.CoverageVersion;$result['ObservationIsNotAcceptance']=$true;
$result['Focused']=$focusedInfo;$result['FocusedRecentFiles']=$focusedRecent;$result['OwnedNativeProcesses']=$nativeProcesses
$result|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'MonitorLast.json') -Encoding UTF8
$result|ConvertTo-Json -Depth 8
