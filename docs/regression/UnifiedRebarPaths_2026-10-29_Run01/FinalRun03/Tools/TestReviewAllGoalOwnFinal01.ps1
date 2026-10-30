# -------------------- ДЛЯ ТЕСТОВ --------------------
# После full PASS и публикации Ready30 заново читает все прежние и текущие own PID+UTC ticks; ничего не закрывает.
param([Parameter(Mandatory=$true)][string]$LaunchDirectory)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
Import-Module 'C:/Windows/System32/WindowsPowerShell/v1.0/Modules/Microsoft.PowerShell.Utility/Microsoft.PowerShell.Utility.psd1' -Global -ErrorAction Stop
. (Join-Path $PSScriptRoot 'TestObserveOwnIdentityFinal01.ps1')
$root='C:/Users/avidzba/Desktop/программы/В_РАБОТЕ!/RC_Section_NDM'
$launch=(Resolve-Path -LiteralPath (Join-Path $root $LaunchDirectory)).Path
if($launch-ine[IO.Path]::GetFullPath((Join-Path $root 'docs/regression/FullSuiteLaunch_2026-10-29_UnifiedRebar_Run03'))){throw 'Exact Run03 launch required'}
$request=Get-Content -LiteralPath (Join-Path $launch 'LaunchRequest.json') -Raw -Encoding UTF8|ConvertFrom-Json
$complete=Get-Content -LiteralPath (Join-Path $root ($request.EvidenceDirectory+'/CompleteCurrentAcceptance.json')) -Raw -Encoding UTF8|ConvertFrom-Json
$full=Get-Content -LiteralPath (Join-Path $launch 'AllOwnClosuresReviewed.json') -Raw -Encoding UTF8|ConvertFrom-Json
$scope=Get-Content -LiteralPath (Join-Path $launch 'TerminalScopeReview.json') -Raw -Encoding UTF8|ConvertFrom-Json
$previous=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'AllPreFullOwnClosures20.json') -Raw -Encoding UTF8|ConvertFrom-Json
$ready=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'Ready30Publication03.json') -Raw -Encoding UTF8|ConvertFrom-Json
$native=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'Ready30SchemaNative02/OwnClosure.json') -Raw -Encoding UTF8|ConvertFrom-Json
if($request.SourceCommit-cne'8910245966806984e349b004cc8e0da0d4dccb86'-or$request.SourceSHA256-cne'D71B27FDA6A7EAF1F8EC708D07111524171C14ACDF1C711B6278491D64E91FBE'){throw 'Exact accepted source epoch required'}
if($complete.Status-cne'PASS'-or$complete.SourceSHA256-cne$request.SourceSHA256-or$complete.CoreSuites-ne182-or$complete.FocusedSingleCases-ne122){throw 'Whole actual current full PASS required'}
if(-not$full.AllOwnIdentitiesTerminal-or-not$full.PrimaryLogsReviewed-or$full.SourceSHA256-cne$request.SourceSHA256-or$full.InputsChecked-ne505-or-not$full.InputsUnchanged-or-not$scope.PrimaryLogsReviewed-or$scope.CoreMacrosReviewed-ne182-or$scope.FocusedSingleCasesReviewed-ne122-or$scope.FocusedRecordCount-ne26-or@($scope.Failures).Count-or@($scope.InventoryErrors).Count){throw 'Whole primary and full own proofs required'}
if($ready.Status-cne'PUBLISHED_CURRENT_CODE_READY30_WITH_ORIGINAL_DATA'-or$ready.LaunchDirectory-cne$LaunchDirectory-or$ready.AcceptedMainSHA256-cne$request.SourceSHA256-or-not$ready.NativeOwnTerminal-or$ready.SolverRecalculated-or-not$ready.UserInputsResultsAndPlotPreserved){throw 'Actual Ready30 publication required'}
if(-not$native.ExactIdentityTerminal-or-not$native.QuitRequested-or$native.ForceKilled-or-not$previous.AllOwnIdentitiesTerminal-or$previous.UniqueOwnCount-ne231-or@($previous.Identities).Count-ne231){throw 'Historical231 and normal Ready native closure required'}
if((Get-FileHash -LiteralPath (Join-Path $root 'workbook/output/RC_Section_NDM.xlsm')).Hash-cne$request.SourceSHA256-or(Get-FileHash -LiteralPath (Join-Path $root 'workbook/output/VBA_All_Code.txt')).Hash-cne$request.AllCodeSHA256-or(Get-FileHash -LiteralPath $ready.Ready).Hash-cne$ready.ReadySHA256){throw 'Published book/export hashes changed'}
$unique=@{}
# -------------------- ДЛЯ ТЕСТОВ --------------------
# Объединяет только полные PID+UTC ticks, сохраняя происхождение; повторное использование PID остаётся отдельной идентичностью.
function Test-AddGoalOwnFinal01([int]$IdentityID,[string]$Ticks,[string]$Name,[string]$Provenance){
    if($IdentityID-le0-or$Ticks-notmatch'^\d{18}$'){throw 'Incomplete goal own identity'}
    $key=[string]$IdentityID+'/'+$Ticks
    if(-not$unique.ContainsKey($key)){$unique[$key]=[ordered]@{PID=$IdentityID;Ticks=$Ticks;Name=$Name;Provenance=@($Provenance)}}else{
        $unique[$key].Provenance+=,$Provenance
        if($Name){if($unique[$key].Name-and$unique[$key].Name-ine$Name){throw 'Conflicting executable names for one exact identity'};$unique[$key].Name=$Name}
    }
}
foreach($row in $previous.Identities){Test-AddGoalOwnFinal01 ([int]$row.PID) ([string]$row.StartTicksUTC) ([string]$row.Name) 'P/AllPreFullOwnClosures20.json'}
foreach($row in $full.Identities){Test-AddGoalOwnFinal01 ([int]$row.PID) ([string]$row.StartTicksUTC) ([string]$row.Name) 'L3/AllOwnClosuresReviewed.json'}
Test-AddGoalOwnFinal01 ([int]$native.PID) ([string]$native.StartTicksUTC) 'EXCEL' 'P/Ready30SchemaNative02/OwnClosure.json'
$rows=@();$allTerminal=$true
foreach($identity in $unique.Values){
    $row=Test-ObserveOwnIdentityFinal01 -IdentityID $identity.PID -ExpectedTicks $identity.Ticks -ExpectedName $identity.Name
    $row|Add-Member -NotePropertyName Provenance -NotePropertyValue $identity.Provenance
    if(-not$row.ExactIdentityTerminal){$allTerminal=$false}
    $rows+=,$row
}
$proof=[ordered]@{Status=$(if($allTerminal){'ALL_GOAL_OWN_IDENTITIES_TERMINAL_AFTER_FULL_AND_READY'}else{'WAITING_LIVE_OR_UNKNOWN_GOAL_OWN'});ObservedUTC=[DateTimeOffset]::UtcNow.ToString('o');SourceSHA256=$request.SourceSHA256;AcceptedCommit=$request.SourceCommit;ReadySHA256=$ready.ReadySHA256;PreviousCount=231;FullOwnCount=@($full.Identities).Count;ReadyNativeCount=1;UniqueOwnCount=$rows.Count;AllOwnIdentitiesTerminal=$allTerminal;UnknownAcceptedAsAbsent=$false;AnyProcessClosed=$false;Identities=$rows;ProofDoesNotReplaceFullAcceptance=$true}
$path=Join-Path $PSScriptRoot $(if($allTerminal){'AllGoalOwnClosuresFinal01.json'}else{'GoalOwnClosuresPendingFinal01.json'})
$proof|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $path -Encoding UTF8
[pscustomobject]@{Status=$proof.Status;UniqueOwnCount=$rows.Count;AllOwnIdentitiesTerminal=$allTerminal;AnyProcessClosed=$false}|ConvertTo-Json
if(-not$allTerminal){exit 1}
