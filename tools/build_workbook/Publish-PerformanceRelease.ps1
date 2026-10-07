# Publishes only the accepted private release. Original user hashes and source
# hashes are checked again; the user's execution report is never replaced.
param([string]$Directory='docs/regression/Performance/ReleaseV2')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$evidence=Join-Path $root $Directory
$expectedBook='D48A5FFCC17A3FC4F1F2A6FC7080952BD01BA08FCB7151CFBF15EA6E887EDB25'
$expectedExport='E20F213475B2DF0592E4CEECABA4E706C7B1AADBB822D23FD6A3E0D197B5EEFE'
$expectedReport='D65B41D99B4E8969F5ADBB7196C9FB41989DEBB1287827F4F18D5D17B98ECCBC'
function Read-Json([string]$path) {Get-Content -LiteralPath $path -Raw | ConvertFrom-Json}
function Assert-Hash([string]$path,[string]$expected) {
    if((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -cne $expected) {throw "Artifact changed: $path"}
}
$acceptance=Read-Json (Join-Path $evidence 'Acceptance.json')
$equality=Read-Json (Join-Path $evidence 'SourceEquality.json')
if(-not $acceptance.SaveReopenExact -or -not $acceptance.PassiveExportExact -or
    $acceptance.PassiveSolveCount -ne 0 -or $acceptance.RepeatedMacros -ne 10 -or
    -not $acceptance.InputsNamesWidthsPreserved) {throw 'Release acceptance is incomplete.'}
if(-not $equality.AllPassed -or $equality.RepositoryComponents -ne 121 -or
    $equality.ActualComponents -ne 126 -or $equality.ProductionClasses -ne 88 -or
    @($equality.GeneratedDocumentChecks).Count -ne 5) {throw 'VBA agreement is incomplete.'}
foreach($check in $equality.Checks) {
    $files=@(Get-ChildItem -LiteralPath (Join-Path $root 'src'),(Join-Path $root 'tests') -Recurse -File |
        Where-Object {$_.BaseName -eq $check.Component -and $_.Extension -in '.cls','.bas'})
    if($files.Count -ne 1) {throw 'Source identity is ambiguous.'}
    Assert-Hash $files[0].FullName $check.SourceSHA256
}
Assert-Hash (Join-Path $PSScriptRoot 'Build-Workbook.ps1') $equality.BuildScriptSHA256
foreach($path in @((Join-Path $evidence 'Validation.json'),
    (Join-Path $root 'docs/regression/Performance/CompiledFreshBuildV2/Validation.json'))) {
    $checks=Read-Json $path
    $checks=@($checks)
    if($checks.Count -ne 27 -or @($checks | Where-Object {-not $_.Passed}).Count) {throw 'Workbook validation is incomplete.'}
}
$ui=Get-Content -LiteralPath (Join-Path $evidence 'ConfigPresentation.txt') -Raw
if($ui -notmatch 'TOTAL_CONFIG_PRESENTATION: passed=1583; failed=0' -or
    $ui -notmatch 'SOURCE_UNCHANGED: True' -or $ui -match '(?m)^FAIL:') {throw 'Release UI verification failed.'}
foreach($folder in @('DirectedFinal','IsolationFinal')) {
    $checks=Read-Json (Join-Path $root "docs/regression/Performance/$folder/Summary.json")
    $checks=@($checks)
    $count=17;if($folder -eq 'IsolationFinal') {$count=9}
    if($checks.Count -ne $count -or @($checks | Where-Object {-not $_.Passed}).Count) {throw "Directed gates are incomplete: $folder"}
}
$measurements=Read-Json (Join-Path $root 'docs/regression/Performance/AcceptedMeasurements.json')
$measurements=@($measurements)
if($measurements.Count -ne 74 -or @($measurements | Where-Object {
    -not $_.ExactValues -or $_.A.Total.N -lt 5 -or $_.B.Total.N -lt 5}).Count) {throw 'A/B acceptance is incomplete.'}
$outputBook=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$outputExport=Join-Path $root 'workbook/output/VBA_All_Code.txt'
$outputReport=Join-Path $root 'workbook/output/RC_Section_NDM_execution_report.txt'
Assert-Hash $outputBook $expectedBook
Assert-Hash $outputExport $expectedExport
Assert-Hash $outputReport $expectedReport
$releaseBook=Join-Path $evidence 'RC_Section_NDM.xlsm'
$releaseExport=Join-Path $evidence 'VBA_All_Code.txt'
Assert-Hash $releaseBook $acceptance.ReleaseSHA256
$exportHash=(Get-FileHash -LiteralPath $releaseExport -Algorithm SHA256).Hash
$sourceManifest=@($equality.Checks | Sort-Object Component | ForEach-Object {
    [ordered]@{Component=$_.Component;SHA256=$_.SourceSHA256}
})
$sourceText=($sourceManifest | ForEach-Object {$_.Component+'|'+$_.SHA256}) -join "`n"
$sha=[Security.Cryptography.SHA256]::Create()
try {$sourceHash=([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($sourceText)))).Replace('-','')}
finally {$sha.Dispose()}
Copy-Item -LiteralPath $releaseBook -Destination $outputBook
Copy-Item -LiteralPath $releaseExport -Destination $outputExport
Assert-Hash $outputBook $acceptance.ReleaseSHA256
Assert-Hash $outputExport $exportHash
Assert-Hash $outputReport $expectedReport
[ordered]@{
    Published=$true;Utc=(Get-Date).ToUniversalTime().ToString('o');
    BaselineGitSHA='df3d5eb787e95246b2b01dfb19bdc76ed82d98f8';
    IntermediateGitSHA='32ec4755086ac472496f2d5c306503a5436a0261';
    WorkbookSHA256=$acceptance.ReleaseSHA256;ExportSHA256=$exportHash;
    UserReportSHA256=$expectedReport;UserReportPreserved=$true;
    SourceAggregateSHA256=$sourceHash;SourceAggregateRule='SHA256 of UTF8 component|hash lines, sorted by Component, LF, no trailing newline';
    Sources=$sourceManifest;ActualVbaComponents=126;ProductionClasses=88;
    ExactABCases=74;DirectedFamilies=17;IsolationFamilies=9;
    FreshValidation=27;ReleaseValidation=27;ReleaseUiPassed=1583;
    SaveReopenExact=$true;RepeatedMacros=10;PassiveSolveCount=0;
    OldFullOnOffSuitesRun=$false;LiveAutoCADRun=$false
} | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $evidence 'Publication.json') -Encoding UTF8
Write-Output "PUBLISHED: book=$($acceptance.ReleaseSHA256); VBA=$exportHash; user report unchanged."
