# Accepts only complete On/Off reports for the unchanged published artifacts.
# Writes evidence metadata; never opens Excel/CAD or modifies the workbook/code.
param([string]$Directory='docs/regression/Performance/ContourCADRelease')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$folder=[IO.Path]::GetFullPath((Join-Path $root $Directory))
$allowed=(Join-Path $root 'docs/regression/Performance')+[IO.Path]::DirectorySeparatorChar
if(-not $folder.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)) {throw 'Acceptance metadata must stay in Performance evidence.'}
$publicationPath=Join-Path $folder 'Publication.json'
$publication=Get-Content -LiteralPath $publicationPath -Raw | ConvertFrom-Json
if(-not $publication.Published -or -not $publication.MeasuredCodeMatchesFinal) {throw 'Publication is incomplete.'}
$artifacts=@{
    'workbook/output/RC_Section_NDM.xlsm'=$publication.WorkbookSHA256
    'workbook/output/VBA_All_Code.txt'=$publication.ExportSHA256
    'workbook/output/RC_Section_NDM_execution_report.txt'=$publication.UserReportSHA256
}
foreach($path in $artifacts.Keys) {
    if((Get-FileHash -LiteralPath (Join-Path $root $path)).Hash -cne $artifacts[$path]) {throw "Published artifact changed: $path"}
}
$source=Get-Content -LiteralPath (Join-Path $folder 'SourceEquality.json') -Raw | ConvertFrom-Json
if(-not $source.AllPassed -or $source.RepositoryComponents -ne 121 -or $source.ActualComponents -ne 126 -or $source.ProductionClasses -ne 88) {throw 'Source equality is incomplete.'}
if((Get-FileHash -LiteralPath (Join-Path $PSScriptRoot 'Build-Workbook.ps1')).Hash -cne $source.BuildScriptSHA256) {throw 'Build script changed.'}
foreach($check in $source.Checks) {
    $files=@(Get-ChildItem -LiteralPath (Join-Path $root 'src'),(Join-Path $root 'tests') -Recurse -File |
        Where-Object {$_.BaseName -eq $check.Component -and $_.Extension -in '.bas','.cls'})
    if($files.Count -ne 1 -or (Get-FileHash -LiteralPath $files[0].FullName).Hash -cne $check.SourceSHA256) {throw "Source changed: $($check.Component)"}
}
$expected=@('modTestGeometry.RunGeometryTests','modTestMaterialDiagrams.RunMaterialDiagramTests',
    'modTestSectionSolver.RunSectionSolverTests','modTestCapacitySolver.RunCapacitySolverTests',
    'modTestCrackWidth.RunCrackWidthTests','modTestBatchCalculation.RunBatchCalculationTests',
    'modTestWorkbookInterface.RunWorkbookInterfaceTests','modTestRegressionBaseline.RunRegressionBaselineTests')
$modes=@()
foreach($item in @(@{Report='FullOn';Mode='Yes'},@{Report='FullOff';Mode='No'})) {
    $text=Get-Content -LiteralPath (Join-Path $folder ($item.Report+'.txt')) -Raw -Encoding UTF8
    if($text -match '(?m)^FAIL:|^SCRIPT ERROR:|^RUNTIME ERROR|^WATCHDOG_FAILURE:' -or
        -not $text.Contains('SOURCE_SHA256: '+$publication.WorkbookSHA256) -or
        -not $text.Contains('GLOBAL_MODE: General.DiagramExtension='+$item.Mode+';') -or
        -not $text.Contains('Geometry.Type=RectSet; fixture only, source unchanged') -or
        -not $text.Contains('SOURCE_UNCHANGED: True') -or
        -not $text.Contains('WATCHDOG_COMPLETED: exit=0; source unchanged=True')) {throw "Invalid full report: $($item.Report)"}
    $matches=[regex]::Matches($text,'(?ms)^SUITE_STARTED: (?<suite>[^;\r\n]+); (?<start>[^\r\n]+)\r?\n(?<body>.*?)^SUITE_FINISHED: \k<suite>; (?<finish>[^\r\n]+)\r?$')
    if($matches.Count -ne 8) {throw "Missing suites: $($item.Report)"}
    $suites=@()
    for($i=0;$i -lt 8;$i++) {
        $match=$matches[$i]
        if($match.Groups['suite'].Value -cne $expected[$i]) {throw 'Full suite order changed.'}
        $totals=[regex]::Matches($match.Groups['body'].Value,'(?m)^TOTAL(?:_[A-Z0-9_]+)?: passed=(?<passed>\d+); failed=(?<failed>\d+)')
        if($totals.Count -lt 1 -or @($totals | Where-Object {[int]$_.Groups['failed'].Value -ne 0}).Count) {throw "Failed or missing total: $($expected[$i])"}
        $last=$totals[$totals.Count-1]
        $start=[datetime]::ParseExact($match.Groups['start'].Value,'s',[Globalization.CultureInfo]::InvariantCulture)
        $finish=[datetime]::ParseExact($match.Groups['finish'].Value,'s',[Globalization.CultureInfo]::InvariantCulture)
        $suites += [ordered]@{Suite=$expected[$i];Passed=[int]$last.Groups['passed'].Value;Failed=0;
            StartedLocal=$start.ToString('s');FinishedLocal=$finish.ToString('s');Seconds=($finish-$start).TotalSeconds}
    }
    $modes += [ordered]@{Mode=$item.Mode;Report=$item.Report+'.txt';ReportSHA256=(Get-FileHash -LiteralPath (Join-Path $folder ($item.Report+'.txt'))).Hash;
        Suites=$suites;SourceUnchanged=$true;ExplicitHistoricalOnOverrides=[regex]::Matches($text,'(?m)^SUITE_MODE_OVERRIDE:').Count}
}
$publication.FullOnOffCompleted=$true
$publication | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $publicationPath -Encoding UTF8
[ordered]@{Passed=$true;CompletedUTC=[datetime]::UtcNow.ToString('o');WorkbookSHA256=$publication.WorkbookSHA256;
    ExportSHA256=$publication.ExportSHA256;UserReportPreserved=$true;SourceComponents=121;GeneratedComponents=5;
    ProductionClasses=88;Modes=$modes;FixtureGeometryType='RectSet';HistoricalExplicitOnSetupsUnchanged=$true} |
    ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $folder 'Acceptance.json') -Encoding UTF8
Write-Output 'FINAL ACCEPTANCE: 16 complete suites, no failures, source and published artifacts unchanged.'
