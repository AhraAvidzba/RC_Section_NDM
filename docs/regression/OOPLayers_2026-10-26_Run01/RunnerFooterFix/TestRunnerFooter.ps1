# ДЛЯ ТЕСТОВ: воспроизводит ошибку Measure-Object на Hashtable и проверяет исправленный footer PS5.1.
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSHOME 'Modules/Microsoft.PowerShell.Utility/Microsoft.PowerShell.Utility.psd1') -Global -ErrorAction Stop
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../../..')).Path
$run=Get-Content -LiteralPath (Join-Path $root 'docs/regression/FullOnOff_2026-10-26_OOP_Run01/Run.json') -Raw -Encoding UTF8|ConvertFrom-Json
$jobs=@($run.Jobs|ForEach-Object {@{ExpectedSuites=[int]$_.ExpectedSuites;Group=[string]$_.Group;Mode=[string]$_.Mode}})
$reproduced=$false;$originalError=''
try{$unused=($jobs|Measure-Object ExpectedSuites -Sum).Sum}catch{$reproduced=$true;$originalError=$_.FullyQualifiedErrorId}
if(-not $reproduced -or $originalError -notlike '*GenericMeasurePropertyNotFound*'){throw 'Original PS5.1 footer failure not reproduced.'}
$actual=($jobs|ForEach-Object {[int]$_.ExpectedSuites}|Measure-Object -Sum).Sum
if($actual-ne 168 -or $jobs.Count-ne 12){throw 'Actual declared core coverage changed.'}
$plain=(@(@{ExpectedSuites=8},@{ExpectedSuites=36},@{ExpectedSuites=6})|ForEach-Object {[int]$_.ExpectedSuites}|Measure-Object -Sum).Sum
if($plain-ne 50){throw 'Hashtable sample failed.'}
$objects=($run.Jobs|ForEach-Object {[int]$_.ExpectedSuites}|Measure-Object -Sum).Sum
if($objects-ne 168){throw 'Deserialized manifest coverage differs.'}
$source=Get-Content -LiteralPath (Join-Path $root 'tools/build_workbook/Run-FullOnOffAcceptance.ps1') -Raw -Encoding UTF8
if(-not $source.Contains('($jobs|ForEach-Object {[int]$_.ExpectedSuites}|Measure-Object -Sum).Sum')){throw 'Production runner does not contain the tested expression.'}
$complete=Get-Content -LiteralPath (Join-Path $root 'docs/regression/FullOnOff_2026-10-26_OOP_Run01/CompleteCurrentAcceptance.json') -Raw -Encoding UTF8|ConvertFrom-Json
if($complete.Status-ne 'PASS' -or $complete.CoreSuites-ne $actual){throw 'Independent complete acceptance coverage differs.'}
$footer=[ordered]@{Status='COMPLETED_PENDING_LOG_REVIEW';SourceSHA256=$run.SourceSHA256;ExpectedSuites=$actual;SummarizeExitCode=0;FinalizeExitCode=0;CompleteAcceptance='CompleteCurrentAcceptance.json';ProducedBy='Post-run footer re-evaluation; original driver completion remains unchanged'}
$json=$footer|ConvertTo-Json -Depth 5
if(($json|ConvertFrom-Json).ExpectedSuites-ne 168){throw 'Footer serialization failed.'}
$json|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'CorrectedFooter.json') -Encoding UTF8
$errors=$null;$tokens=$null;[void][Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'tools/build_workbook/Run-FullOnOffAcceptance.ps1'),[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'Patched runner parse failed.'}
[ordered]@{At=(Get-Date -Format o);Status='PASS_FOOTER_ONLY';PSVersion=$PSVersionTable.PSVersion.ToString();OriginalFailureReproduced=$reproduced;OriginalError=$originalError;ActualJobCount=$jobs.Count;FixedExpectedSuites=$actual;JsonManifestExpectedSuites=$objects;HashtableSample50=$plain;CorrectedFooterSerializes=$true;RunnerParses=$true;NoExcelStarted=$true;NoFullRerun=$true;BookSHA256=$complete.SourceSHA256}|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'TestResult.json') -Encoding UTF8
Get-Content -LiteralPath (Join-Path $PSScriptRoot 'TestResult.json') -Encoding UTF8