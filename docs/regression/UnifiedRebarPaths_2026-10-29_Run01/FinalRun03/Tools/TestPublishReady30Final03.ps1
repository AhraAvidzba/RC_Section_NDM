# -------------------- ДЛЯ ТЕСТОВ --------------------
# Публикует только точную Ready30, принятую после полного PASS и собственного native Quit.
param([Parameter(Mandatory=$true)][string]$LaunchDirectory)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$root='C:/Users/avidzba/Desktop/программы/В_РАБОТЕ!/RC_Section_NDM'
$private=$PSScriptRoot
$launch=(Resolve-Path -LiteralPath (Join-Path $root $LaunchDirectory)).Path
$request=Get-Content -LiteralPath (Join-Path $launch 'LaunchRequest.json') -Raw|ConvertFrom-Json
$complete=Get-Content -LiteralPath (Join-Path $root ($request.EvidenceDirectory+'/CompleteCurrentAcceptance.json')) -Raw|ConvertFrom-Json
$fullOwn=Get-Content -LiteralPath (Join-Path $launch 'AllOwnClosuresReviewed.json') -Raw|ConvertFrom-Json
$review=Get-Content -LiteralPath (Join-Path $private 'Ready30FinalReview03.json') -Raw|ConvertFrom-Json
$prep=Get-Content -LiteralPath (Join-Path $private 'Ready30SchemaPreparation02.json') -Raw|ConvertFrom-Json
$native=Get-Content -LiteralPath (Join-Path $private 'Ready30SchemaNative02/NativeVerification.json') -Raw|ConvertFrom-Json
$own=Get-Content -LiteralPath (Join-Path $private 'Ready30SchemaNative02/OwnClosure.json') -Raw|ConvertFrom-Json
if($complete.Status-cne'PASS'-or$complete.SourceSHA256-cne$request.SourceSHA256-or$complete.CoreSuites-ne182-or$complete.FocusedSingleCases-ne122-or-not$fullOwn.AllOwnIdentitiesTerminal-or-not$fullOwn.PrimaryLogsReviewed-or$fullOwn.SourceSHA256-cne$request.SourceSHA256-or$fullOwn.CoreMacrosReviewed-ne182-or$fullOwn.FocusedSingleCasesReviewed-ne122-or$fullOwn.FocusedRecordCount-ne26){throw 'Current entire full acceptance and own closures required'}
if($review.Status-cne'PASS_CURRENT_FULL_NATIVE_READY_SCHEMA_DATA_PRESERVATION'-or$review.LaunchDirectory-cne$LaunchDirectory-or$review.AcceptedMainSHA256-cne$request.SourceSHA256-or$review.AcceptedCommit-cne$request.SourceCommit-or$review.Published-or$review.SolverRecalculated-or-not$review.MainReadyExportExactExceptDate-or-not$review.ResultsAndCalculationSheetByteExact-or-not$review.ConfigCellPayloadsExact-or-not$review.OtherGeometryUserCellsExact-or-not$review.MainCommentsMovedExact-or-not$review.OldStylesPrefixExact-or-not$review.OriginalThirtyLoadsAndFourInputErrPreserved){throw 'Actual final native/data review required'}
if(-not$own.ExactIdentityTerminal-or-not$own.QuitRequested-or$own.ForceKilled-or$review.NativeOwnPID-ne$own.PID-or[string]$review.NativeOwnStartTicksUTC-cne[string]$own.StartTicksUTC){throw 'Normal exact native Ready own closure required'}
. (Join-Path $private 'TestObserveOwnIdentityFinal01.ps1')
$ownChecks=@();$allTerminal=$true
foreach($identity in @($fullOwn.Identities)+@([pscustomobject]@{PID=$own.PID;StartTicksUTC=$own.StartTicksUTC;Name='EXCEL'})){
 $state=Test-ObserveOwnIdentityFinal01 -IdentityID ([int]$identity.PID) -ExpectedTicks ([string]$identity.StartTicksUTC) -ExpectedName ([string]$identity.Name)
 $ownChecks+=,$state
 if(-not$state.ExactIdentityTerminal){$allTerminal=$false}
}
$ownReviewPath=Join-Path $private ('Ready30PublicationOwnReview03_'+[DateTimeOffset]::UtcNow.Ticks+'.json')
[ordered]@{ObservedUTC=[DateTimeOffset]::UtcNow.ToString('o');SourceSHA256=$request.SourceSHA256;AllOwnIdentitiesTerminal=$allTerminal;UnknownAcceptedAsAbsent=$false;AnyProcessClosed=$false;Identities=$ownChecks}|ConvertTo-Json -Depth 7|Set-Content -LiteralPath $ownReviewPath -Encoding UTF8
if(-not$allTerminal){throw 'Live or unknown exact own identity; no Ready file replaced'}
$source=[string]$review.Ready
$destination=Join-Path $root 'workbook/output/NDM_Test_1232_30/RC_Section_NDM.xlsm'
if([IO.Path]::GetFullPath([string]$review.Destination)-cne[IO.Path]::GetFullPath($destination)){throw 'Unexpected Ready destination'}
if([IO.Path]::GetFullPath($source)-cne[IO.Path]::GetFullPath([string]$prep.Ready)){throw 'Unexpected private prepared source'}
if($review.ReadySHA256-cne$prep.ReadySHA256-or$native.SHA256-cne$prep.ReadySHA256-or$review.OriginalReadySHA256-cne$prep.SourceSHA256){throw 'Prepared/native/review source hash differs'}
if((Get-FileHash -LiteralPath $source).Hash-cne$review.ReadySHA256-or(Get-FileHash -LiteralPath $destination).Hash-cne$review.OriginalReadySHA256){throw 'Ready changed; preserve and review user edits rather than overwrite'}
$main=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$allCode=Join-Path $root 'workbook/output/VBA_All_Code.txt'
if((Get-FileHash -LiteralPath $main).Hash-cne$request.SourceSHA256-or(Get-FileHash -LiteralPath $allCode).Hash-cne$request.AllCodeSHA256){throw 'Full accepted main/export changed'}
$other=Join-Path (Split-Path -Parent $destination) 'RC_Section_NDM_ok.xlsm'
$otherSHA='31F188692A01CC421A37CF962598778A5AD1FE6D9AB5D96DD41C29389C0919F7'
if((Get-FileHash -LiteralPath $other).Hash-cne$otherSHA){throw 'Unrelated small example preservation fence changed'}
$backup=Join-Path $private 'BeforeReady30FinalPublication03'
if(Test-Path -LiteralPath $backup){throw 'Publication already attempted; inspect proof instead of repeating'}
[void][IO.Directory]::CreateDirectory($backup)
$backupBook=Join-Path $backup 'RC_Section_NDM.xlsm'
[IO.File]::Copy($destination,$backupBook,$false)
if((Get-FileHash -LiteralPath $backupBook).Hash-cne$review.OriginalReadySHA256){throw 'Ready backup hash differs'}
[IO.File]::Copy($source,$destination,$true)
if((Get-FileHash -LiteralPath $destination).Hash-cne$review.ReadySHA256-or(Get-FileHash -LiteralPath $other).Hash-cne$otherSHA-or(Get-FileHash -LiteralPath $main).Hash-cne$request.SourceSHA256){throw 'Publication hash verification failed'}
$proof=[ordered]@{Status='PUBLISHED_CURRENT_CODE_READY30_WITH_ORIGINAL_DATA';ObservedUTC=[DateTimeOffset]::UtcNow.ToString('o');LaunchDirectory=$LaunchDirectory;AcceptedMainSHA256=$request.SourceSHA256;AcceptedCommit=$request.SourceCommit;BeforeReadySHA256=$review.OriginalReadySHA256;ReadySHA256=$review.ReadySHA256;Ready=$destination;Backup=$backupBook;NativeComponents=181;Data='1232 concrete /8 rebar /30 original load IDs /4 historical InputErr';MainReadyCodeExactExceptExportDate=$true;UserInputsResultsAndPlotPreserved=$true;OtherCopyUnchanged=$true;MainCleanBookUnchanged=$true;SolverRecalculated=$false;FullAcceptanceAppliesToMainSourceSHAOnly=$true;NativeOwnTerminal=$true;OptionalNDMCompanionPublished=$false;ReviewSHA256=(Get-FileHash -LiteralPath (Join-Path $private 'Ready30FinalReview03.json')).Hash}
$proof.OwnRecheckProof=$ownReviewPath
$proof.OwnRecheckSHA256=(Get-FileHash -LiteralPath $ownReviewPath).Hash
$proof|ConvertTo-Json -Depth 7|Set-Content -LiteralPath (Join-Path $private 'Ready30Publication03.json') -Encoding utf8
$proof|ConvertTo-Json -Depth 7
