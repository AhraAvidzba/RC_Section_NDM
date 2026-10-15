# Полные on/off проверки в независимых копиях с общим сроком завершения.
param(
    [string]$SourceWorkbook='workbook/output/RC_Section_NDM.xlsm',
    [Parameter(Mandatory=$true)][string]$EvidenceDirectory,
    [Parameter(Mandatory=$true)][string]$Deadline,
    [ValidateRange(1,12)][double]$ForecastHours=5
)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$source=(Resolve-Path -LiteralPath (Join-Path $root $SourceWorkbook)).Path
$directory=[IO.Path]::GetFullPath((Join-Path $root $EvidenceDirectory))
$allowed=[IO.Path]::GetFullPath((Join-Path $root 'docs/regression'))+[IO.Path]::DirectorySeparatorChar
if (-not $directory.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)) {throw 'Evidence must stay inside docs/regression.'}
if (Test-Path -LiteralPath $directory) {throw 'A fresh evidence directory is required.'}
$limit=[DateTimeOffset]::Parse($Deadline,[Globalization.CultureInfo]::InvariantCulture)
if ([DateTimeOffset]::UtcNow.AddHours($ForecastHours).AddMinutes(15) -ge $limit) {throw 'Forecast does not fit the requested deadline.'}
[void][IO.Directory]::CreateDirectory((Join-Path $directory 'Source'))
$frozen=Join-Path $directory 'Source/RC_Section_NDM.xlsm'
Copy-Item -LiteralPath $source -Destination $frozen
Copy-Item -LiteralPath $PSCommandPath -Destination (Join-Path $directory 'FullRunner.ps1')
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'Run-Audit02ModeTests.ps1') -Destination (Join-Path $directory 'ModeRunner.ps1')
$sourceHash=(Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
if ((Get-FileHash -LiteralPath $frozen -Algorithm SHA256).Hash -ne $sourceHash) {throw 'Source copy mismatch.'}
$geometry=@(
    'modTestDeclarativeGeometry.RunCoreGeometryTests',
    'modTestPolygonSpec.RunPolygonSpecTests',
    'modTestSectionRecipes.RunRecipeDefinitionTests',
    'modTestSectionRecipes.RunRecipeRuleTests',
    'modTestSectionRecipes.RunRecipeMeshBoundaryTests',
    'modTestSectionRecipes.RunRecipeAnnotationCountTests',
    'modTestGeometryPublication.RunAnnotationIsolationTests',
    'modTestGeometryPublication.RunRecipeContextFreezeTests',
    'modTestGeometryPublication.RunPublishedBuildTests',
    'modTestGeometryPublication.RunDefinitionLifecycleTests',
    'modTestRecipeOperations.RunRecipeOperationTests',
    'modTestRecipeOperations.RunRecipeDiagnosticTests',
    'modTestRecipeOperations.RunResolveAnchorTests',
    'modTestGeometryQuery.RunGeometryQueryTests',
    'modTestRecipeAnnotations.RunArbitraryAnnotationGeometryTests',
    'modTestRecipeAnnotations.RunRecipeAnnotationSnapshotTests',
    'modTestAutoCADContours.RunImportedDimensionTests',
    'modTestAutoCADContours.RunAnnotationSectionScaleTests',
    'modTestContourMesh.RunContourMeshTests',
    'modTestAudit02Api.RunAudit02PolygonApiTests',
    'modTestAudit02Api.RunAudit02OptionsTests',
    'modTestAudit02Api.RunAudit02InputTests',
    'modTestAudit02Api.RunAudit02AnnotationTests',
    'modTestAudit02Api.RunAudit02GuideExamplesTests',
    'modTestAudit02Api.RunAudit02CloseoutContactTests',
    'modTestAudit02Api.RunAudit02CloseoutRectSetInputTests'
)
$sp35=@(
    'modTestSP35CrackWidth.RunSP35FormulaTests',
    'modTestSP35CrackWidth.RunSP35LocalOpeningTests',
    'modTestSP35CrackWidth.RunSP35PreparationTests',
    'modTestSP35CrackWidth.RunSP35NeighborSettingsTests',
    'modTestSP35CrackWidth.RunSP35PipelineTests',
    'modTestSP35CrackWidth.RunSP35EndToEndTests'
)
$jobs=@()
foreach ($mode in @('No','Yes')) {
    $label=if($mode -eq 'No') {'Off'} else {'On'}
    $jobs+=@{Group='Full';Mode=$mode;Label=$label;Macro=@();ExpectedSuites=8;MaximumSeconds=12600}
    $jobs+=@{Group='Geometry';Mode=$mode;Label=$label;Macro=$geometry;ExpectedSuites=$geometry.Count;MaximumSeconds=3600}
    $jobs+=@{Group='SP35';Mode=$mode;Label=$label;Macro=$sp35;ExpectedSuites=$sp35.Count;MaximumSeconds=3600}
}
$manifest=[ordered]@{Status='RUNNING';StartedUTC=[DateTimeOffset]::UtcNow.ToString('o');Deadline=$limit.ToString('o');ForecastHours=$ForecastHours;
    SourceCommit=(& git -C $root rev-parse HEAD);SourceWorkbook=$source;SourceSHA256=$sourceHash;FrozenWorkbook=$frozen;
    RunnerSHA256=(Get-FileHash -LiteralPath $PSCommandPath -Algorithm SHA256).Hash;
    ModeRunnerSHA256=(Get-FileHash -LiteralPath (Join-Path $PSScriptRoot 'Run-Audit02ModeTests.ps1') -Algorithm SHA256).Hash;
    GeometryTypeInTestCopies='RectSet';Jobs=$jobs;CompletedJobs=@();SourceChanged=$false}
$manifestPath=Join-Path $directory 'Run.json'
function Save-Run {$manifest | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $manifestPath -Encoding UTF8}
Save-Run
$failed=$false
foreach ($job in $jobs) {
    $remaining=[int][Math]::Floor(($limit-[DateTimeOffset]::UtcNow).TotalSeconds)-600
    if ($remaining -lt 30) {$manifest.Status='DEADLINE_REACHED';$failed=$true;break}
    $timeout=[Math]::Min([int]$job.MaximumSeconds,$remaining)
    $relative=$EvidenceDirectory+'/'+$job.Group+'_'+$job.Label+'.txt'
    $manifest['CurrentJob']=$job.Group+'_'+$job.Label
    Save-Run
    Write-Output ('START: '+$manifest.CurrentJob+'; '+[DateTimeOffset]::Now.ToString('o'))
    $arguments=@('-NoProfile','-ExecutionPolicy','Bypass','-File',(Join-Path $PSScriptRoot 'Run-Audit03Watchdog.ps1'),
        '-SourceWorkbook',$frozen,'-ReportPath',$relative,'-Mode',$job.Mode,
        '-TimeoutSeconds',[string]$timeout,'-FixtureGeometryType','RectSet')
    if ($job.Macro.Count) {$arguments+=@('-Macro',($job.Macro -join ','))}
    $started=[DateTimeOffset]::UtcNow
    & (Join-Path $PSHOME 'powershell.exe') @arguments
    $code=$LASTEXITCODE
    $manifest.CompletedJobs+=@{Group=$job.Group;Mode=$job.Mode;Label=$job.Label;ExitCode=$code;Report=$relative;
        ExpectedSuites=$job.ExpectedSuites;StartedUTC=$started.ToString('o');FinishedUTC=[DateTimeOffset]::UtcNow.ToString('o')}
    if ($code -ne 0) {$failed=$true}
    Save-Run
    Write-Output ('FINISH: '+$manifest.CurrentJob+'; exit='+$code+'; '+[DateTimeOffset]::Now.ToString('o'))
}
$manifest.SourceChanged=(Get-FileHash -LiteralPath $frozen -Algorithm SHA256).Hash -ne $sourceHash
if ($manifest.SourceChanged) {$failed=$true}
$manifest['RunnerChanged']=((Get-FileHash -LiteralPath $PSCommandPath -Algorithm SHA256).Hash -ne $manifest.RunnerSHA256 -or
    (Get-FileHash -LiteralPath (Join-Path $PSScriptRoot 'Run-Audit02ModeTests.ps1') -Algorithm SHA256).Hash -ne $manifest.ModeRunnerSHA256)
if ($manifest.RunnerChanged) {$failed=$true}
$manifest['FinishedUTC']=[DateTimeOffset]::UtcNow.ToString('o')
if ($manifest.Status -eq 'RUNNING') {$manifest.Status=if($failed){'REVIEW_REQUIRED'}else{'COMPLETED_PENDING_LOG_REVIEW'}}
Save-Run
if ($failed) {exit 1}
