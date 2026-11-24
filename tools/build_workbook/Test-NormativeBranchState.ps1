# ДЛЯ ТЕСТОВ: проверяет частичный отказ устойчивости и восстановление той же модели.
param(
    [Parameter(Mandatory=$true)][string]$BaselineWorkbook,
    [Parameter(Mandatory=$true)][string]$CandidateWorkbook,
    [Parameter(Mandatory=$true)][string]$EvidenceDirectory
)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$directory=Join-Path $root $EvidenceDirectory
New-Item -ItemType Directory -Path $directory -Force | Out-Null
$results=@()
foreach($label in @('Baseline','Candidate')){
    $source=if($label -eq 'Baseline'){$BaselineWorkbook}else{$CandidateWorkbook}
    foreach($mode in @('No','Yes')){
        $relative=Join-Path $EvidenceDirectory ($label+'_'+$mode+'.txt')
        & (Join-Path $PSScriptRoot 'Run-Audit03Watchdog.ps1') -SourceWorkbook $source -ReportPath $relative -Mode $mode -Macro 'modTestBatchCalculation.RunNormativePartialBranchFailureTest' -TimeoutSeconds 120 -NormativeCoreGuard
        if($LASTEXITCODE -ne 0){throw ('Partial branch test failed: '+$label+' '+$mode)}
        $body=[IO.File]::ReadAllText((Join-Path $root $relative))
        if($body -notmatch '(?m)^TOTAL_NORMATIVE_PARTIAL_BRANCH: passed=11; failed=0\r?$'){throw 'Incomplete partial-branch test'}
        $results+=,[ordered]@{Label=$label;Mode=$mode;Passed=$true;WorkbookSHA256=(Get-FileHash -LiteralPath $source).Hash;Report=$relative;Assertions=11}
        $results|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $directory 'Acceptance.json') -Encoding UTF8
    }
}
