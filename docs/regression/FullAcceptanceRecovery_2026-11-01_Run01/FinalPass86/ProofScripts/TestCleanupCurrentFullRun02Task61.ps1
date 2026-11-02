# -------------------- ДЛЯ ТЕСТОВ --------------------
# Удаляет только однократную задачу Current_Run02 после полного терминала, чтения журналов и свежего аудита всех собственных процессов.
param([Parameter(Mandatory=$true)][string]$LaunchDirectory)
$ErrorActionPreference='Stop'
if($PSVersionTable.PSVersion.Major-ne5){throw 'Windows PowerShell5 required; no task removed'}
Import-Module 'C:/Windows/System32/WindowsPowerShell/v1.0/Modules/Microsoft.PowerShell.Utility/Microsoft.PowerShell.Utility.psd1' -Global -ErrorAction Stop
Import-Module ScheduledTasks -ErrorAction Stop
$root=(Resolve-Path -LiteralPath 'C:/Users/avidzba/Desktop/программы/В_РАБОТЕ!/RC_Section_NDM').Path
$launch=(Resolve-Path -LiteralPath (Join-Path $root $LaunchDirectory)).Path
$expectedLaunch=Join-Path $root 'docs/regression/FullSuiteLaunch_2026-11-01_Current_Run02'
if($launch-ine[IO.Path]::GetFullPath($expectedLaunch)){throw 'Exact Current_Run02 launch directory required'}
$request=Get-Content -LiteralPath (Join-Path $launch 'LaunchRequest.json') -Raw -Encoding UTF8|ConvertFrom-Json
$full=Get-Content -LiteralPath (Join-Path $launch 'FullLaunch.json') -Raw -Encoding UTF8|ConvertFrom-Json
$completion=Get-Content -LiteralPath (Join-Path $launch 'Completion.json') -Raw -Encoding UTF8|ConvertFrom-Json
$name='RC_NDM_Full_20261101_Current_Run02'
$source='831D725C06439CA1B771645F18B0D546169C1E22A97BA7688D64CB3F8509B74A'
$commit='5f18de1f88a657ae312bc37b0798bc511f4bb218'
if($request.ScheduledTaskName-cne$name-or$full.ScheduledTaskName-cne$name-or$full.PID-ne98996-or[string]$full.StartTicks-cne'639291636492982491'){throw 'Exact task/driver identity required'}
if($request.SourceCommit-cne$commit-or$full.SourceCommit-cne$commit-or$request.SourceSHA256-cne$source-or$full.SourceSHA256-cne$source-or$completion.SourceSHA256-cne$source-or-not$completion.FinishedUTC){throw 'Exact terminal Current_Run02 epoch required'}
$proofPath=Join-Path $launch 'ScheduledTaskCleanup61.json'
if(Test-Path -LiteralPath $proofPath){throw 'Cleanup proof already exists; inspect instead of repeating'}
# Повторная штатная проверка читает все PID+UTC ticks и 522 входов; она не закрывает процессы и не объявляет общий PASS.
& (Join-Path $launch 'Test-VerifyAllOwnClosures.ps1')
if($LASTEXITCODE-ne0){throw 'Fresh all-own verification failed; no task removed'}
$own=Get-Content -LiteralPath (Join-Path $launch 'AllOwnClosuresReviewed.json') -Raw -Encoding UTF8|ConvertFrom-Json
if(-not$own.AllOwnIdentitiesTerminal-or-not$own.PrimaryLogsReviewed-or$own.SourceSHA256-cne$source-or$own.SourceCommit-cne$commit-or$own.InputsChecked-ne522-or-not$own.InputsUnchanged-or$own.CoreMacrosReviewed-ne182-or$own.FocusedSingleCasesReviewed-ne122-or$own.FocusedRecordCount-ne26){throw 'Whole current scope and fresh exact own closure required'}
# -------------------- ДЛЯ ТЕСТОВ --------------------
# Читает корневую папку планировщика; ошибка доступа не подменяется отсутствием задачи.
function Test-ReadOwnTask {
    $matches=@(Get-ScheduledTask -TaskPath '\' -ErrorAction Stop|Where-Object {$_.TaskName-ieq$name})
    if($matches.Count-gt1){throw 'Ambiguous root task identity'}
    if($matches.Count-eq0){return $null}
    if($matches[0].TaskName-cne$name-or$matches[0].TaskPath-cne'\'){throw 'Actual task name/path differs'}
    return $matches[0]
}
$task=Test-ReadOwnTask
$backup=Join-Path $PSScriptRoot 'CurrentRun02TaskBeforeCleanup61.xml'
$proof=[ordered]@{ObservedUTC=[DateTimeOffset]::UtcNow.ToString('o');TaskName=$name;SourceSHA256=$source;SourceCommit=$commit;AllOwnIdentitiesTerminal=$true;InputsChecked=522;AnyProcessClosed=$false;DeletedByThisCheck=$false;OwnProofSHA256=(Get-FileHash -LiteralPath (Join-Path $launch 'AllOwnClosuresReviewed.json')).Hash}
if($null-eq$task){
    $proof.Status='ALREADY_ABSENT_AFTER_CURRENT_OWN_CLOSURES'
}else{
    $actions=@($task.Actions)
    $expectedExecutable=Join-Path $PSHOME 'powershell.exe'
    $expectedHost=Join-Path $launch 'FullDriverHost.ps1'
    $expectedArguments='-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "'+$expectedHost+'"'
    if($actions.Count-ne1-or[IO.Path]::GetFullPath([string]$actions[0].Execute)-ine[IO.Path]::GetFullPath($expectedExecutable)-or[string]$actions[0].Arguments-cne$expectedArguments-or[IO.Path]::GetFullPath([string]$actions[0].WorkingDirectory)-ine$root){throw 'Actual task Action differs; no task removed'}
    if(@($task.Triggers|Where-Object {$null-ne$_}).Count-ne0-or[string]$task.State-eq'Running'){throw 'Task has triggers or is running; no task removed'}
    if(Test-Path -LiteralPath $backup){throw 'Task backup already exists; inspect instead of overwriting'}
    $xml=Export-ScheduledTask -TaskName $name -TaskPath '\' -ErrorAction Stop
    [IO.File]::WriteAllText($backup,[string]$xml,[Text.Encoding]::Unicode)
    Unregister-ScheduledTask -TaskName $name -TaskPath '\' -Confirm:$false -ErrorAction Stop
    if(Test-ReadOwnTask){throw 'Own task remains after deletion'}
    $proof.Status='EXACT_CURRENT_FULL_TASK_REMOVED_AFTER_ALL_OWN_EXIT'
    $proof.DeletedByThisCheck=$true
    $proof.Execute=[string]$actions[0].Execute
    $proof.Arguments=[string]$actions[0].Arguments
    $proof.WorkingDirectory=[string]$actions[0].WorkingDirectory
    $proof.Triggers=0
    $proof.TaskXMLBackup=$backup
    $proof.TaskXMLSHA256=(Get-FileHash -LiteralPath $backup).Hash
}
$proof|ConvertTo-Json -Depth 5|Set-Content -LiteralPath $proofPath -Encoding UTF8
[pscustomobject]@{Status=$proof.Status;TaskName=$name;DeletedByThisCheck=$proof.DeletedByThisCheck;AnyProcessClosed=$false}|ConvertTo-Json
