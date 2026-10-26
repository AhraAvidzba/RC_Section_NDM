# -------------------- ДЛЯ ТЕСТОВ --------------------
# Итоговое наблюдение собственных Excel/AutoCAD по прямым PID + UTC startTicks.
# Читает только разрешённые proof JSON, не заходит в Inputs и не завершает процессы.
[CmdletBinding()]
param(
    [string]$EvidenceDirectory=''
)
$ErrorActionPreference='Stop'
if($PSVersionTable.PSVersion.Major -ne 5){throw 'Use Windows PowerShell 5.1.'}
if([string]::IsNullOrWhiteSpace($EvidenceDirectory)){$EvidenceDirectory=$PSScriptRoot}
$evidence=(Resolve-Path -LiteralPath $EvidenceDirectory).Path.TrimEnd([IO.Path]::DirectorySeparatorChar)
if([IO.Path]::GetFileName($evidence) -cne 'ContourCellGeometry_2026-10-26_Run01'){throw 'Only the current owned Run01 evidence directory is allowed.'}
if($evidence -ine $PSScriptRoot.TrimEnd([IO.Path]::DirectorySeparatorChar)){throw 'Evidence must be the same Run01 directory that contains this script.'}
$destination=Join-Path $evidence 'ProcessClosureFinal.json'
$issues=New-Object 'System.Collections.Generic.List[object]'
$filesRead=New-Object 'System.Collections.Generic.List[object]'
$ignored=New-Object 'System.Collections.Generic.List[object]'
$identities=@{}
$protectedProcessIds=@(34980,22916) # Explicitly external Excel/AutoCAD from the current task; never classify as owned.

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Получает только непосредственное свойство: Prior/Existing и вложенные объекты не обходятся.
function Test-DirectProperty($record,[string]$name){
    if($null -eq $record){return $null}
    $property=$record.PSObject.Properties[$name]
    if($null -eq $property){return $null}
    return $property.Value
}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Объединяет повторные прямые доказательства одной identity, сохраняя происхождение каждой записи.
function Test-AddOwnIdentity($record,[string]$relativePath,[string]$identityKind,[string]$pidKey,[string]$ticksKey,[string]$defaultName){
    $ownedFlag=Test-DirectProperty $record 'Owned'
    if($null -ne $ownedFlag -and $ownedFlag -eq $false){return}
    $processValue=Test-DirectProperty $record $pidKey
    $ticksValue=Test-DirectProperty $record $ticksKey
    if($null -eq $processValue -and $null -eq $ticksValue){return}
    $processNumber=0
    $startTicks=[long]0
    if(-not [int]::TryParse([string]$processValue,[ref]$processNumber) -or -not [long]::TryParse([string]$ticksValue,[ref]$startTicks)){
        $issues.Add([ordered]@{File=$relativePath;Kind=$identityKind;Problem='Invalid or incomplete direct PID/startTicks pair';PID=$processValue;StartTicksUTC=$ticksValue})
        return
    }
    if($processNumber -le 0 -or $startTicks -le 0){
        $ignored.Add([ordered]@{File=$relativePath;Kind=$identityKind;Reason='Process was not created';PID=$processNumber;StartTicksUTC=[string]$startTicks})
        return
    }
    if($processNumber -in $protectedProcessIds){
        $ignored.Add([ordered]@{File=$relativePath;Kind=$identityKind;Reason='Explicitly protected external process; never owned';PID=$processNumber;StartTicksUTC=[string]$startTicks})
        return
    }
    $expectedName=$defaultName
    if($identityKind -eq 'Generic'){
        $directName=[string](Test-DirectProperty $record 'Name')
        if(-not [string]::IsNullOrWhiteSpace($directName)){$expectedName=$directName}
    }
    if($expectedName -ine 'EXCEL' -and $expectedName -ine 'acad'){
        $issues.Add([ordered]@{File=$relativePath;Kind=$identityKind;Problem='Direct proof has an unexpected process name';Name=$expectedName;PID=$processNumber;StartTicksUTC=[string]$startTicks})
        return
    }
    $identityKey=([string]$processNumber)+':'+([string]$startTicks)
    if(-not $identities.ContainsKey($identityKey)){
        $identities[$identityKey]=[ordered]@{PID=$processNumber;StartTicksUTC=[string]$startTicks;ExpectedName=$expectedName;Evidence=New-Object 'System.Collections.Generic.List[object]'}
    }elseif($identities[$identityKey].ExpectedName -ine $expectedName){
        $issues.Add([ordered]@{File=$relativePath;Problem='Conflicting process name for one owned identity';PID=$processNumber;StartTicksUTC=[string]$startTicks})
    }
    $identities[$identityKey].Evidence.Add([ordered]@{File=$relativePath;Kind=$identityKind;PIDKey=$pidKey;StartTicksKey=$ticksKey;RecordedClosed=(Test-DirectProperty $record 'ExactIdentityAbsent');RecordedResidualForceKilled=(Test-DirectProperty $record 'ResidualForceKilled')})
}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Отбирает proof-файлы по имени до чтения; Inputs исключается до входа в каталог.
function Test-FindProofFiles([string]$directory){
    foreach($entry in (Get-ChildItem -LiteralPath $directory -Force)){
        if($entry.PSIsContainer){
            if($entry.Name -ine 'Inputs' -and -not ($entry.Attributes -band [IO.FileAttributes]::ReparsePoint)){
                Test-FindProofFiles $entry.FullName
            }
            continue
        }
        $rootFreshBuild=($directory -ieq $evidence -and ($entry.Name -like 'FreshBuild*.json' -or $entry.Name -like 'rootFreshBuild*.json'))
        if($entry.Name -ieq 'Ownership.json' -or $entry.Name -ieq 'ProcessClosure.json' -or $entry.Name -like '*_NativeVerification.json' -or $rootFreshBuild){$entry}
    }
}

foreach($file in @(Test-FindProofFiles $evidence | Sort-Object FullName)){
    $relativePath=$file.FullName.Substring($evidence.Length+1)
    try{
        $raw=[IO.File]::ReadAllText($file.FullName)
        $recordSet=$raw | ConvertFrom-Json
        $sha=(Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
        $filesRead.Add([ordered]@{File=$relativePath;Bytes=$file.Length;SHA256=$sha})
        # ProcessClosure is the only permitted root array; nested arrays are never inspected.
        if($recordSet -is [Array] -and $file.Name -ine 'ProcessClosure.json'){
            $issues.Add([ordered]@{File=$relativePath;Problem='Unexpected root array outside ProcessClosure.json'})
            continue
        }
        foreach($record in @($recordSet)){
            Test-AddOwnIdentity $record $relativePath 'Excel' 'ExcelPID' 'ExcelStartTicksUTC' 'EXCEL'
            Test-AddOwnIdentity $record $relativePath 'CAD' 'CADPID' 'CADStartTicksUTC' 'acad'
            Test-AddOwnIdentity $record $relativePath 'Generic' 'PID' 'StartTicksUTC' 'EXCEL'
        }
    }catch{
        $issues.Add([ordered]@{File=$relativePath;Problem='Proof JSON could not be read';Error=$_.Exception.Message})
    }
}

$observations=New-Object 'System.Collections.Generic.List[object]'
foreach($identity in @($identities.Values | Sort-Object PID,StartTicksUTC)){
    $observedProcess=Get-Process -Id $identity.PID -ErrorAction SilentlyContinue
    $actualTicks=$null
    $actualName=$null
    $exactLive=$false
    $observationError=$null
    $cpu=$null
    $workingSet=$null
    $privateBytes=$null
    if($null -ne $observedProcess){
        try{
            $actualTicks=[string]$observedProcess.StartTime.ToUniversalTime().Ticks
            $actualName=$observedProcess.ProcessName
            $exactLive=($actualTicks -ceq $identity.StartTicksUTC)
            $cpu=$observedProcess.CPU
            $workingSet=$observedProcess.WorkingSet64
            $privateBytes=$observedProcess.PrivateMemorySize64
            if($exactLive -and $actualName -ine $identity.ExpectedName){
                $issues.Add([ordered]@{PID=$identity.PID;StartTicksUTC=$identity.StartTicksUTC;Problem='Exact identity is live under an unexpected process name';ObservedName=$actualName})
            }
        }catch{
            # A process can exit between Get-Process and StartTime; re-observe only this same PID.
            $recheck=Get-Process -Id $identity.PID -ErrorAction SilentlyContinue
            if($null -ne $recheck){
                $observationError=$_.Exception.Message
                $issues.Add([ordered]@{PID=$identity.PID;StartTicksUTC=$identity.StartTicksUTC;Problem='Process start time could not be verified';Error=$observationError})
            }else{
                $exactLive=$false
                $actualTicks=$null
            }
        }
    }
    $state=if($null -ne $observationError){'UNKNOWN'}elseif($exactLive){'OWN_EXACT_IDENTITY_LIVE'}elseif($null -ne $actualTicks){'PID_REUSED_OWN_IDENTITY_ABSENT'}else{'OWN_IDENTITY_ABSENT'}
    $observations.Add([ordered]@{PID=$identity.PID;StartTicksUTC=$identity.StartTicksUTC;ExpectedName=$identity.ExpectedName;State=$state;ExactIdentityLive=$exactLive;ExactIdentityAbsent=(-not $exactLive -and $null -eq $observationError);ObservedName=$actualName;ObservedStartTicksUTC=$actualTicks;CPUSeconds=$cpu;WorkingSetBytes=$workingSet;PrivateBytes=$privateBytes;ObservationError=$observationError;Evidence=@($identity.Evidence.ToArray())})
}

$memory=$null
try{
    $perfMemory=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
    $operatingSystem=Get-CimInstance Win32_OperatingSystem
    $memory=[ordered]@{ObservedUTC=[DateTimeOffset]::UtcNow.ToString('o');PhysicalTotalBytes=[double]$operatingSystem.TotalVisibleMemorySize*1024;AvailablePhysicalBytes=[double]$perfMemory.AvailableBytes;AvailablePhysicalGiB=[double]$perfMemory.AvailableMBytes/1024;CommittedBytes=[double]$perfMemory.CommittedBytes;CommitLimitBytes=[double]$perfMemory.CommitLimit;FreeCommitBytes=([double]$perfMemory.CommitLimit-[double]$perfMemory.CommittedBytes);FreeCommitGiB=([double]$perfMemory.CommitLimit-[double]$perfMemory.CommittedBytes)/1GB;CommitPercent=100*[double]$perfMemory.CommittedBytes/[double]$perfMemory.CommitLimit;Source='Win32_PerfFormattedData_PerfOS_Memory + Win32_OperatingSystem'}
}catch{$issues.Add([ordered]@{Problem='Actual system memory could not be read';Error=$_.Exception.Message})}

$live=@($observations | Where-Object ExactIdentityLive)
$unknown=@($observations | Where-Object {$_.State -eq 'UNKNOWN'})
$clean=($identities.Count -gt 0 -and $filesRead.Count -gt 0 -and $issues.Count -eq 0 -and $live.Count -eq 0 -and $unknown.Count -eq 0)
$result=[ordered]@{Status=$(if($clean){'PASS_ALL_RECORDED_OWN_IDENTITIES_ABSENT'}else{'REVIEW_REQUIRED'});ObservedUTC=[DateTimeOffset]::UtcNow.ToString('o');EvidenceDirectory=$evidence;ScriptSHA256=(Get-FileHash -LiteralPath $PSCommandPath -Algorithm SHA256).Hash;ReadOnlyObservation=$true;ProcessesStarted=0;ProcessesStopped=0;OnlyOutputWritten=$destination;InputsExcluded=$true;NestedPriorExistingArraysInspected=$false;ProtectedExternalPIDs=$protectedProcessIds;ProofFilesRead=@($filesRead.ToArray());ProofFilesReadCount=$filesRead.Count;UniqueOwnIdentityCount=$identities.Count;LiveOwnIdentityCount=$live.Count;UnknownIdentityCount=$unknown.Count;AllRecordedOwnedIdentitiesAbsent=$clean;Observations=@($observations.ToArray());IgnoredDirectRecords=@($ignored.ToArray());Memory=$memory;Issues=@($issues.ToArray());FullRegressionPass=$false;Scope='Only direct own identities in allowlisted evidence files; no claim about unrecorded processes or full regression acceptance.'}
$result | ConvertTo-Json -Depth 15 | Set-Content -LiteralPath $destination -Encoding UTF8
[ordered]@{Status=$result.Status;ProofFilesReadCount=$filesRead.Count;UniqueOwnIdentityCount=$identities.Count;LiveOwnIdentityCount=$live.Count;UnknownIdentityCount=$unknown.Count;IssueCount=$issues.Count;AvailablePhysicalGiB=$memory.AvailablePhysicalGiB;FreeCommitGiB=$memory.FreeCommitGiB;Output=$destination} | ConvertTo-Json -Depth 4
if(-not $clean){exit 1}
