# Проверяет геометрические ожидания и подготовку нагрузок до долгой полной приёмки.
param([Parameter(Mandatory=$true)][string]$EvidenceDirectory,
      [string]$SourceWorkbook='workbook/output/RC_Section_NDM.xlsm')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$directory=[IO.Path]::GetFullPath((Join-Path $root $EvidenceDirectory))
if(-not $directory.StartsWith((Join-Path $root 'docs/regression')+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $directory)){throw 'Use a new regression directory.'}
[void][IO.Directory]::CreateDirectory($directory)
$source=(Resolve-Path -LiteralPath (Join-Path $root $SourceWorkbook)).Path
$hash=(Get-FileHash -LiteralPath $source).Hash
$frozen=Join-Path $directory 'Input.xlsm'
Copy-Item -LiteralPath $source -Destination $frozen
$runner=[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'Run-FullOnOffAcceptance.ps1'),[Text.Encoding]::UTF8)
$macros=@('modTestGeometry.RunGeometryTests','modTestMaterialDiagrams.RunMaterialDiagramTests')
foreach($group in @('geometry','sp35','recipes')){
    $section=[regex]::Match($runner,'(?s)\$'+$group+'=@\((.*?)\r?\n\)')
    if(-not $section.Success){throw "Missing full group: $group"}
    $macros+=@([regex]::Matches($section.Groups[1].Value,"'([^']+)'" ) | ForEach-Object {$_.Groups[1].Value})
}
$macros+=@('modTestWorkbookInterface.RunSharedConfigPreflightTests','modTestWorkbookInterface.RunConfigPresentationTests',
    'modTestGeometryReview.RunGeometryReviewTests','modTestRecipeOperations.RunCollapsingOffsetTests','modTestRecipeOperations.RunFullBoundaryOffsetTests')
$macros=@($macros | Select-Object -Unique)
$plan=[ordered]@{SourceSHA256=$hash;SourceWorkbook=$source;Started=(Get-Date).ToString('o');Macros=$macros;Status='RUNNING';Results=@()}
function Save-Plan{$plan | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $directory 'Preflight.json') -Encoding UTF8}
Save-Plan
foreach($job in @(
    @{ID='Expectations';Mode='No';Geometry='Rectangles';Macros=$macros;Seconds=1800},
    @{ID='SavedOffsets';Mode='No';Geometry='PointedSection';Macros=@('modTestRecipeOperations.RunSavedPointedOffsetTests','modTestRecipeOperations.RunSavedOffsetCalculationTests');Seconds=180}
)){
    $report=$EvidenceDirectory+'/'+$job.ID+'.txt'
    & (Join-Path $PSHOME 'powershell.exe') -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Run-Audit03Watchdog.ps1') -SourceWorkbook $frozen -ReportPath $report -Macro ($job.Macros -join ',') -Mode $job.Mode -FixtureGeometryType $job.Geometry -TimeoutSeconds $job.Seconds
    $exit=$LASTEXITCODE
    $text=[IO.File]::ReadAllText((Join-Path $root $report),[Text.Encoding]::UTF8)
    $passed=$exit -eq 0 -and $text.Contains('WATCHDOG_COMPLETED: exit=0; source unchanged=True') -and $text -notmatch '(?m)^(FAIL:|RUNTIME ERROR|SCRIPT ERROR|WATCHDOG_FAILURE|CLEANUP ERROR)'
    $plan.Results+=@{ID=$job.ID;ExitCode=$exit;Passed=$passed;Report=$report;Macros=$job.Macros}
    Save-Plan
    if(-not $passed){$plan.Status='REVIEW_REQUIRED';Save-Plan;exit 1}
}
$plan['SourceUnchanged']=(Get-FileHash -LiteralPath $source).Hash -ceq $hash -and (Get-FileHash -LiteralPath $frozen).Hash -ceq $hash
if(-not $plan.SourceUnchanged){throw 'Source changed during preflight.'}
$plan.Status='PASSED';$plan['Finished']=(Get-Date).ToString('o');Save-Plan
Write-Output 'FULL_EXPECTATIONS_PREFLIGHT_PASSED'
