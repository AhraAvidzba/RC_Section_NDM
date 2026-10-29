# ==========================================================================
# ДЛЯ ТЕСТОВ: фактический read-only аудит всех записанных PID+UTC ticks после full.
# Не закрывает процессы; Live/Unknown оставляет исходный набор ожидающим завершения.
# ==========================================================================
param([string]$ScopeReview='TerminalScopeReview.json')
$ErrorActionPreference='Stop'
if($PSVersionTable.PSVersion.Major -ne5){throw 'Windows PowerShell5.1 required'}
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path
Set-Location -LiteralPath $root
$reviewPath=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot $ScopeReview)).Path
if(-not $reviewPath.StartsWith($PSScriptRoot+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw 'Own terminal scope review required'}
$review=Get-Content -LiteralPath $reviewPath -Raw -Encoding UTF8|ConvertFrom-Json
$request=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'LaunchRequest.json') -Raw -Encoding UTF8|ConvertFrom-Json
$completion=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'Completion.json') -Raw -Encoding UTF8|ConvertFrom-Json
if($review.SourceSHA256 -cne $request.SourceSHA256 -or $review.SourceCommit -cne $request.SourceCommit -or $completion.SourceSHA256 -cne $request.SourceSHA256){throw 'Historical/foreign scope cannot prove current full exit'}
if(-not $review.PrimaryLogsReviewed -or $review.CoreReportsRead -ne12 -or $review.CoreMacrosReviewed -ne182 -or $review.FocusedRecordCount -ne26 -or $review.FocusedSingleCasesReviewed -ne122){throw 'Entire current182+122/26 inventory required'}
if(-not $completion.FinishedUTC -or @($completion.ChangedInputs).Count -or $completion.HeadChanged){throw 'Original completion and clean input fence required'}
$evidenceRoot=[IO.Path]::GetFullPath((Join-Path $root $request.EvidenceDirectory))
foreach($row in $review.Primary){
    $path=[IO.Path]::GetFullPath((Join-Path $root $row.Path))
    if(-not $path.StartsWith($evidenceRoot+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw 'Primary outside current evidence'}
    if(-not $row.AllBytesRead -or (Get-FileHash -LiteralPath $path).Hash -cne $row.SHA256 -or $row.SourceMatches -eq $false){throw 'Primary evidence changed or source mismatched'}
}
$rows=@($request.ProtectedInputs)+@($request.SourceComponents)+@($request.ParticipatingScripts)
if($rows.Count -ne504 -or (& git rev-parse HEAD) -cne $request.SourceCommit){throw 'Original504/HEAD mismatch'}
foreach($row in $rows){if((Get-FileHash -LiteralPath (Join-Path $root $row.Path)).Hash -cne $row.SHA256){throw ('Original input changed: '+$row.Path)}}
$identities=@();$terminal=$true;$seen=@{}
foreach($row in $review.Identities){
    $key=[string]$row.Name+':'+[string]$row.PID+':'+[string]$row.StartTicksUTC
    if($seen.ContainsKey($key) -or [int]$row.PID -le0 -or [long]$row.StartTicksUTC -le0 -or [string]$row.Name -notin @('EXCEL','acad','powershell')){throw 'Duplicate/invalid own identity'}
    $seen[$key]=$true
    $state=[ordered]@{Name=$row.Name;PID=$row.PID;StartTicksUTC=[string]$row.StartTicksUTC;Provenance=$row.Provenance;State='Unknown';ObservedUTC=[DateTimeOffset]::UtcNow.ToString('o');Error=$null}
    $p=$null
    try{$p=Get-Process -Id ([int]$row.PID) -ErrorAction Stop}catch{if($_.FullyQualifiedErrorId -match'^NoProcessFoundForGivenId(,|$)'){$state.State='Absent'}else{$state.Error=$_.Exception.Message}}
    if($p){
        try{
            $observedName=[string]$p.ProcessName
            if([string]::IsNullOrWhiteSpace($observedName)){throw 'Process name unavailable'}
            $state['ObservedName']=$observedName
            if($observedName -ine[string]$row.Name){
                $native=Get-CimInstance Win32_Process -Filter ('ProcessId='+[int]$row.PID) -ErrorAction Stop
                if($null -eq$native -or $null -eq$native.CreationDate -or [IO.Path]::GetFileNameWithoutExtension([string]$native.Name) -ine$observedName){throw 'Different process name could not be independently confirmed'}
                $state['ObservedStartTicksUTC']=[string]$native.CreationDate.ToUniversalTime().Ticks
                $state['ReuseProof']='Different executable name confirmed by Get-Process and Win32_Process; CIM UTC creation recorded'
                $state.State='ReusedByDifferentProcess'
            }else{
            if($null -eq $p.StartTime){throw 'StartTime unavailable'}
            $state['ObservedName']=$p.ProcessName;$state['ObservedStartTicksUTC']=[string]$p.StartTime.ToUniversalTime().Ticks
            if($p.ProcessName -ieq[string]$row.Name -and [string]$p.StartTime.ToUniversalTime().Ticks -ceq[string]$row.StartTicksUTC){$state.State='Live'}else{$state.State='ReusedByDifferentProcess'}
            }
        }catch{$state.Error=$_.Exception.Message}
    }
    if($state.State -notin @('Absent','ReusedByDifferentProcess')){$terminal=$false}
    $identities+=,$state
}
foreach($file in @('DriverRun.json','HostRun.json')){
    $identity=Get-Content -LiteralPath (Join-Path $PSScriptRoot $file) -Raw -Encoding UTF8|ConvertFrom-Json
    if(@($identities|Where-Object{$_.Name -ieq'powershell' -and [int]$_.PID -eq[int]$identity.PID -and [string]$_.StartTicksUTC -ceq[string]$identity.StartTicks}).Count -ne1){throw 'Exact driver/host missing from inventory'}
}
$result=[ordered]@{Status=$(if($terminal){'ALL_CURRENT_FULL_OWN_IDENTITIES_TERMINAL'}else{'WAITING_LIVE_OR_UNKNOWN_OWN_IDENTITIES'});CheckedUTC=[DateTimeOffset]::UtcNow.ToString('o');SourceSHA256=$request.SourceSHA256;SourceCommit=$request.SourceCommit;AllOwnIdentitiesTerminal=$terminal;PrimaryLogsReviewed=$true;CoreMacrosReviewed=$review.CoreMacrosReviewed;FocusedSingleCasesReviewed=$review.FocusedSingleCasesReviewed;FocusedRecordCount=$review.FocusedRecordCount;IncludesFocusedComposition=$true;InputsChecked=$rows.Count;InputsUnchanged=$true;ScopeReview=$reviewPath;ScopeReviewSHA256=(Get-FileHash -LiteralPath $reviewPath).Hash;Identities=$identities;FullAcceptanceClaim=$false;AnyProcessClosed=$false}
$result|ConvertTo-Json -Depth 10|Set-Content -LiteralPath (Join-Path $PSScriptRoot $(if($terminal){'AllOwnClosuresReviewed.json'}else{'AllOwnClosuresPending.json'})) -Encoding UTF8
[pscustomobject]@{Status=$result.Status;Identities=$identities.Count;AllOwnIdentitiesTerminal=$terminal;FullAcceptanceClaim=$false;AnyProcessClosed=$false}|ConvertTo-Json
if(-not $terminal){exit 1}
