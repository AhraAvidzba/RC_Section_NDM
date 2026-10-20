# Проверяет геометрические ожидания и подготовку нагрузок до долгой полной приёмки.
param([Parameter(Mandatory=$true)][string]$EvidenceDirectory,
      [string]$SourceWorkbook='workbook/output/RC_Section_NDM.xlsm',
      [string]$SharedConfigReport='')
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
$macros+=@('modTestWorkbookInterface.RunConfigPresentationTests',
    'modTestGeometryReview.RunGeometryReviewTests','modTestRecipeOperations.RunCollapsingOffsetTests','modTestRecipeOperations.RunFullBoundaryOffsetTests')
$macros=@($macros | Select-Object -Unique)
$plan=[ordered]@{SourceSHA256=$hash;SourceWorkbook=$source;Started=(Get-Date).ToString('o');Macros=(@('modTestWorkbookInterface.RunSharedConfigPreflightTests')+$macros);Status='RUNNING';Results=@()}
function Save-Plan{$plan | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $directory 'Preflight.json') -Encoding UTF8}
Save-Plan
$sharedMacro='modTestWorkbookInterface.RunSharedConfigPreflightTests'
$jobs=@()
if($SharedConfigReport){
    # Повтор не нужен, если отдельный завершённый тест проверял те же байты книги.
    $sharedPath=(Resolve-Path -LiteralPath (Join-Path $root $SharedConfigReport)).Path
    if(-not $sharedPath.StartsWith((Join-Path $root 'docs/regression')+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw 'Config evidence must stay inside docs/regression.'}
    $body=[IO.File]::ReadAllText($sharedPath,[Text.Encoding]::UTF8)
    if($body -notmatch ('(?m)^SOURCE_SHA256: '+$hash+'\r?$') -or
        -not $body.Contains('GLOBAL_MODE: General.DiagramExtension=No;') -or
        ([regex]::Matches($body,'(?m)^SUITE_STARTED: ')).Count -ne 1 -or
        $body -notmatch ('(?m)^SUITE_FINISHED: '+[regex]::Escape($sharedMacro)+';') -or
        $body -notmatch '(?m)^TOTAL_SHARED_CONFIG_PREFLIGHT: passed=[1-9][0-9]*; failed=0\r?$' -or
        $body -notmatch '(?m)^WATCHDOG_COMPLETED: exit=0; source unchanged=True\r?$' -or
        $body -match '(?m)^(FAIL:|RUNTIME ERROR|SCRIPT ERROR|WATCHDOG_FAILURE|CLEANUP ERROR)'){throw 'The separate Config check is incomplete, failed, or used another workbook.'}
    Copy-Item -LiteralPath $sharedPath -Destination (Join-Path $directory 'SharedConfig.txt')
    $plan.Results+=@{ID='SharedConfig';ExitCode=0;Passed=$true;Report=($EvidenceDirectory+'/SharedConfig.txt');Macros=@($sharedMacro);ReusedFrom=$SharedConfigReport;ReportSHA256=(Get-FileHash -LiteralPath $sharedPath).Hash}
    Save-Plan
}else{
    $jobs+=@{ID='SharedConfig';Mode='No';Geometry='Rectangles';Macros=@($sharedMacro);Seconds=3600}
}
$jobs+=@(
    @{ID='Expectations';Mode='No';Geometry='Rectangles';Macros=$macros;Seconds=3600},
    @{ID='SavedOffsets';Mode='No';Geometry='PointedSection';Macros=@('modTestRecipeOperations.RunSavedPointedOffsetTests','modTestRecipeOperations.RunSavedOffsetCalculationTests');Seconds=180}
)
foreach($job in $jobs){
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
