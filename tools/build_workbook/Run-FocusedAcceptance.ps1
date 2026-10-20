# Полный часовой план: 122 случая, batch-15, специальные случаи, CAD и независимый контроль.
param(
    [string]$SourceWorkbook='workbook/output/RC_Section_NDM.xlsm',
    [Parameter(Mandatory=$true)][string]$EvidenceDirectory,
    [ValidateRange(60,3600)][int]$MaxSeconds=3600,
    [string]$DeadlineUTC='',
    [string]$Python=(Join-Path $env:USERPROFILE '.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe')
)
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSHOME 'Modules/Microsoft.PowerShell.Utility/Microsoft.PowerShell.Utility.psd1') -Global -ErrorAction Stop
if ($PSVersionTable.PSEdition -ne 'Desktop') {throw 'Use Windows PowerShell 5.1.'}
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
Set-Location -LiteralPath $root
& $Python (Join-Path $PSScriptRoot 'Test-FocusedAcceptanceVerifier.py')
if($LASTEXITCODE -ne 0){throw 'Independent verifier prerequisites or controls failed before acceptance.'}
$directory=[IO.Path]::GetFullPath((Join-Path $root $EvidenceDirectory))
$allowed=[IO.Path]::GetFullPath((Join-Path $root 'docs/regression'))+[IO.Path]::DirectorySeparatorChar
if (-not $directory.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $directory)) {throw 'A new evidence directory inside docs/regression is required.'}
[void][IO.Directory]::CreateDirectory($directory)
$source=(Resolve-Path -LiteralPath $SourceWorkbook).Path
$protected=@($source,(Join-Path $root 'workbook/output/VBA_All_Code.txt'),(Join-Path $root 'workbook/output/RC_Section_NDM_execution_report.txt')) | Where-Object {Test-Path -LiteralPath $_}
$integrity=@($protected | ForEach-Object {@{Path=$_;SHA256=(Get-FileHash -LiteralPath $_).Hash}})
$started=[DateTimeOffset]::UtcNow
$deadline=$started.AddSeconds($MaxSeconds)
if($DeadlineUTC){$limit=[DateTimeOffset]::Parse($DeadlineUTC);if($limit -lt $deadline){$deadline=$limit}}
$records=New-Object 'System.Collections.Generic.List[object]'
$manifest=[ordered]@{Schema='NDM.FocusedAcceptance/1';Status='RUNNING';SourceWorkbook=$source;SourceSHA256=$integrity[0].SHA256;SourceCommit=(& git rev-parse HEAD);StartedUTC=$started.ToString('o');DeadlineUTC=$deadline.ToString('o');MaxSeconds=$MaxSeconds;Records=@();ProtectedInputs=$integrity}
function Save-Run {$manifest.Records=@($records.ToArray());$manifest | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $directory 'Run.json') -Encoding UTF8}
function Remaining { $seconds=[int][math]::Floor(($deadline-[DateTimeOffset]::UtcNow).TotalSeconds);if($seconds -lt 10){throw 'One-hour deadline reached; acceptance is incomplete.'};return $seconds }
function Invoke-Gate([string]$ID,[string]$Macro,[string]$Mode,[string]$A='',[string]$B='',[int]$Timeout=360,[switch]$Reopen) {
    $relative=$EvidenceDirectory+'/'+$ID+'.txt'
    if(Test-Path -LiteralPath (Join-Path $root $relative)){throw 'Existing evidence cannot be overwritten.'}
    Write-Output ('FOCUSED_STARTED: '+$ID+'; remaining='+(Remaining))
    & (Join-Path $PSScriptRoot 'Run-Audit03Watchdog.ps1') -SourceWorkbook (Join-Path $directory 'Probe/RC_Section_NDM.xlsm') -ReportPath $relative -Macro $Macro -MacroArgument1 $A -MacroArgument2 $B -Mode $Mode -TimeoutSeconds ([math]::Min($Timeout,(Remaining))) -VerifyResultsReopen:$Reopen -VerifyStatusReopen:$Reopen
    $code=$LASTEXITCODE
    $body=Get-Content -LiteralPath (Join-Path $root $relative) -Raw -Encoding UTF8
    $passed=$code -eq 0 -and $body -notmatch '(?m)^(FAIL:|RUNTIME ERROR|SCRIPT ERROR|WATCHDOG_FAILURE|RESTORE ERROR|CLEANUP ERROR)'
    $records.Add(@{ID=$ID;Kind='VBA';Macro=$Macro;Mode=$Mode;Argument1=$A;Argument2=$B;ExitCode=$code;Passed=$passed;Report=$relative;ReportSHA256=(Get-FileHash -LiteralPath (Join-Path $root $relative)).Hash})
    Save-Run
    Write-Output ('FOCUSED_FINISHED: '+$ID+'; passed='+$passed)
}
Save-Run
try {
    $combined=Join-Path $directory 'Combined.bas'
    $helpers=@('CoreRegressionBatchProbe.bas','FocusedSpecialProbe.bas','FocusedForceBounds.bas')
    $body=($helpers | ForEach-Object {[IO.File]::ReadAllText((Join-Path $PSScriptRoot $_),[Text.Encoding]::UTF8)}) -join "`r`n"
    [IO.File]::WriteAllText($combined,$body,[Text.UTF8Encoding]::new($true))
    Copy-Item -LiteralPath $PSCommandPath -Destination (Join-Path $directory 'FocusedRunner.ps1')
    foreach($helper in @('FocusedAcceptanceProbe.bas')+$helpers){Copy-Item -LiteralPath (Join-Path $PSScriptRoot $helper) -Destination (Join-Path $directory $helper)}
    $fixture=Join-Path $PSScriptRoot 'fixtures/FocusedSavedSection.json'
    Copy-Item -LiteralPath $fixture -Destination (Join-Path $directory 'SavedSectionFixture.json')
    & (Join-Path $PSScriptRoot 'Prepare-FocusedAcceptance.ps1') -SourceWorkbook $source -OutputDirectory ($EvidenceDirectory+'/Probe') -ExtraHelperFile ('../../'+$EvidenceDirectory+'/Combined.bas') -SavedFixture $fixture
    & (Join-Path $PSScriptRoot 'Describe-FocusedAcceptance.ps1') -ProbeDirectory ($EvidenceDirectory+'/Probe')
    $input=Get-Content -LiteralPath (Join-Path $directory 'Probe/InputManifest.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach($group in $input.Groups){foreach($mode in @('No','Yes')){
        Invoke-Gate ($group.Shape+'_'+$group.Group+'_'+$mode) 'modTestBatchCalculation.RunFocusedAcceptanceCases' $mode $group.Shape $group.Group -Reopen
    }}
    foreach($mode in @('No','Yes')){
        Invoke-Gate ('Batch15_'+$mode) 'modTestBatchCalculation.RunCoreRegressionBatch15' $mode -Reopen
        foreach($case in @('Stability','AsymmetricSteel')){Invoke-Gate ($case+'_'+$mode) 'modTestBatchCalculation.RunFocusedSpecialCase' $mode $case -Reopen}
    }
    foreach($shape in @('Rectangles','Circle','PointedSection','HollowRectangle','SavedWorkbook')){
        Invoke-Gate ('Model_'+$shape) 'modTestBatchCalculation.DumpFocusedForceBoundsExact' 'No' $shape '' 120
    }
    Invoke-Gate 'OutputPresentation' 'RunConfigPresentationTests' 'No' '' '' 180
    & (Join-Path $PSScriptRoot 'Run-FocusedNativeAcceptance.ps1') -SourceWorkbook $source -ReportDirectory ($EvidenceDirectory+'/NativeCAD') -MaxSeconds ([math]::Min(420,(Remaining)))
    $records.Add(@{ID='NativeCAD';Kind='NativeCAD';Passed=($LASTEXITCODE -eq 0);Report=($EvidenceDirectory+'/NativeCAD/Stage.json')})
    Save-Run
    & $Python (Join-Path $PSScriptRoot 'Verify-FocusedAcceptance.py') --directory $directory --independent
    $records.Add(@{ID='Independent';Kind='Python';Passed=($LASTEXITCODE -eq 0);Report=($EvidenceDirectory+'/Independent.json')})
    $manifest.Status='COMPLETED_PENDING_REVIEW'
} catch {
    $manifest.Status='INCOMPLETE_OR_FAILED';$manifest['Error']=$_.Exception.Message
    Write-Output ('FOCUSED_ERROR: '+$_.Exception.Message)
} finally {
    $manifest['FinishedUTC']=[DateTimeOffset]::UtcNow.ToString('o')
    $manifest['ElapsedSeconds']=([DateTimeOffset]::UtcNow-$started).TotalSeconds
    $manifest['ProtectedInputsUnchanged']=@($integrity | Where-Object {(Get-FileHash -LiteralPath $_.Path).Hash -cne $_.SHA256}).Count -eq 0
    Save-Run
}
& $Python (Join-Path $PSScriptRoot 'Verify-FocusedAcceptance.py') --directory $directory --check
exit $LASTEXITCODE
