# Полный набор приёмки и обязательные дополнения; отказы основных этапов не скрываются.
param(
    [string]$SourceWorkbook='workbook/output/RC_Section_NDM.xlsm',
    [Parameter(Mandatory=$true)][string]$EvidenceDirectory,
    [Parameter(Mandatory=$true)][string]$Deadline,
    [double]$ForecastHours=6
)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$coordinatorBytes=[IO.File]::ReadAllBytes($PSCommandPath)
& (Join-Path $PSHOME 'powershell.exe') -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Run-FullOnOffAcceptance.ps1') -SourceWorkbook $SourceWorkbook -EvidenceDirectory $EvidenceDirectory -Deadline $Deadline -ForecastHours $ForecastHours
$primaryCode=$LASTEXITCODE
$directory=Join-Path $root $EvidenceDirectory
if(-not (Test-Path -LiteralPath (Join-Path $directory 'Run.json'))){throw 'Primary full run did not create its manifest.'}
[IO.File]::WriteAllBytes((Join-Path $directory 'Coordinator.ps1'),$coordinatorBytes)
$supplementCode=1
try{
    & (Join-Path $PSHOME 'powershell.exe') -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Run-CurrentSupplementAcceptance.ps1') -PrimaryDirectory $EvidenceDirectory -Deadline $Deadline
    $supplementCode=$LASTEXITCODE
}catch{
    [IO.File]::WriteAllText((Join-Path $directory 'SupplementStartError.txt'),($_|Out-String),[Text.UTF8Encoding]::new($true))
}
$corePath=Join-Path $directory 'CompleteCurrentAcceptance.json'
$extraPath=Join-Path $directory 'Supplement/Acceptance.json'
$passed=$false;$sourceHash=''
if((Test-Path -LiteralPath $corePath) -and (Test-Path -LiteralPath $extraPath)){
    $core=Get-Content -LiteralPath $corePath -Raw -Encoding UTF8|ConvertFrom-Json
    $extra=Get-Content -LiteralPath $extraPath -Raw -Encoding UTF8|ConvertFrom-Json
    $sourceHash=$core.SourceSHA256
    $passed=$primaryCode -eq 0 -and $supplementCode -eq 0 -and $core.Status -eq 'PASS' -and $extra.Status -eq 'PASS' -and $sourceHash -ceq $extra.SourceSHA256
    $passed=$passed -and $core.CoreSuites -eq 202 -and $core.FocusedSingleCases -eq 122 -and $extra.ExpectedExecutions -eq 14 -and $extra.Results.Count -eq 2
    $passed=$passed -and (($extra.Results|ForEach-Object Mode) -join ',') -ceq 'No,Yes'
}
[ordered]@{Status=if($passed){'PASS'}else{'REVIEW_REQUIRED'};SourceSHA256=$sourceHash;PrimaryExitCode=$primaryCode;SupplementExitCode=$supplementCode;RequiredCoreExecutions=202;RequiredFocusedSingleCases=122;RequiredSupplementExecutions=14;SameWorkbookBytes=$passed;FinishedUTC=[DateTimeOffset]::UtcNow.ToString('o');Deadline=$Deadline;Primary='CompleteCurrentAcceptance.json';Supplement='Supplement/Acceptance.json'}|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $directory 'GoalAcceptance.json') -Encoding UTF8
if(-not $passed){exit 1}
