# Verifies normal calculation/save/reopen on the private release, not the user's book.
param([string]$Directory='docs/regression/PlateauSlope_2026-10-11/Release')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
$directoryPath=Join-Path $root $Directory
$target=Join-Path $directoryPath 'RC_Section_NDM.xlsm'
$source=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$preparation=Get-Content -LiteralPath (Join-Path $directoryPath 'Preparation.json') -Raw | ConvertFrom-Json
$equality=Get-Content -LiteralPath (Join-Path $directoryPath 'SourceEquality.json') -Raw | ConvertFrom-Json
if (-not $equality.AllPassed -or -not $preparation.Compiled -or -not $preparation.InputsResultsNamesWidthsPreserved) {throw 'Preparation/source gates are incomplete.'}
foreach ($gate in @(@('MaterialPlateau.txt','TOTAL_PLATEAU_SLOPE: passed=2047; failed=0'),@('ConfigPresentation.txt','TOTAL_CONFIG_PRESENTATION: passed=1592; failed=0'))) {
    $text=Get-Content -LiteralPath (Join-Path $directoryPath $gate[0]) -Raw
    if (-not $text.Contains($gate[1]) -or $text -match '(?m)^FAIL:|RUNTIME ERROR:|SCRIPT ERROR:') {throw 'A directed gate failed.'}
}
foreach ($gate in @(@('MaterialControls.txt','TOTAL_MATERIAL_CONTROL_UNITS'),@('ResultsDiagrams.txt','TOTAL_PLATEAU_RESULTS'))) {
    $text=Get-Content -LiteralPath (Join-Path $directoryPath $gate[0]) -Raw
    if ($text -notmatch ($gate[1]+': passed=\d+; failed=0') -or $text -match '(?m)^FAIL:|RUNTIME ERROR:|SCRIPT ERROR:') {throw 'A diagram output gate failed.'}
}
if ((Get-FileHash -LiteralPath $source).Hash -cne $preparation.SourceSHA256 -or (Get-FileHash -LiteralPath $target).Hash -cne $preparation.ReleaseSHA256) {throw 'A workbook changed before acceptance.'}
Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class PlateauVerifyWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId); }'
function TextHash([string]$value) {
    $sha=[Security.Cryptography.SHA256]::Create()
    try {return [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($value))).Replace('-','')}
    finally {$sha.Dispose()}
}
function InputSnapshot($book) {
    $tables=@()
    foreach ($name in $book.Names) {
        try {$range=$name.RefersToRange} catch {continue}
        if ([string]$range.Worksheet.Name -eq 'Config') {$tables += [ordered]@{Name=[string]$name.Name;Address=$range.Address();Formula=$range.Formula}}
    }
    $widths=@();$results=$book.Worksheets.Item('Results')
    for ($column=1;$column -le 260;$column++) {$widths += [double]$results.Columns.Item($column).ColumnWidth}
    $config=$book.Worksheets.Item('Config')
    $title=$config.Cells.Find('Контрольные точки диаграмм*',$config.Cells.Item(1,1),-4163,1,1,1,$false,$false,$false)
    return TextHash (([ordered]@{Tables=$tables;Widths=$widths;Controls=$title.Offset(2,0).Resize(52,6).Formula} | ConvertTo-Json -Depth 8 -Compress))
}
function Verify-ConfigChartCurves($book) {
    $config=$book.Worksheets.Item('Config')
    $table=$book.Names.Item('rngSystemSettings').RefersToRange
    $cell=$null
    for ($row=2;$row -le $table.Rows.Count;$row++) {
        if ([string]$table.Cells.Item($row,1).Value2 -ceq 'Solver.PlateauSlope.Enabled') {$cell=$table.Cells.Item($row,2);break}
    }
    if ($null -eq $cell) {throw 'Plateau setting is missing.'}
    $saved=$cell.Formula
    $title=$config.Cells.Find('Контрольные точки диаграмм*',$config.Cells.Item(1,1),-4163,1,1,1,$false,$false,$false)
    $records=@()
    try {
        foreach ($mode in @('No','Yes')) {
            $cell.Value2=$mode;$config.Calculate()
            foreach ($item in @(@('chMaterialConcreteDiagram',2,7),@('chMaterialSteelDiagram',26,31))) {
                $chart=$config.ChartObjects([string]$item[0]).Chart
                for ($index=1;$index -le 2;$index++) {
                    $series=$chart.SeriesCollection($index)
                    $x=@($series.XValues | ForEach-Object {[double]$_});$y=@($series.Values | ForEach-Object {[double]$_})
                    $start=$title.Offset([int]$item[$index],4)
                    $count=5;if ($index -eq 2) {$count=7;if ($item[0] -eq 'chMaterialSteelDiagram') {$count=9}}
                    if ($x.Length -ne $count -or $y.Length -ne $count) {throw 'Chart point count differs from its control table.'}
                    for ($point=0;$point -lt $count;$point++) {
                        if ([double]$x[$point] -ne [double]$start.Offset($point,0).Value2 -or [double]$y[$point] -ne [double]$start.Offset($point,1).Value2) {throw 'A chart does not reflect current diagram coordinates.'}
                    }
                    $records += [ordered]@{Mode=$mode;Chart=$item[0];Series=$index;Points=$count;SHA256=(TextHash (($x,$y|ConvertTo-Json -Depth 4 -Compress)))}
                }
                if ($mode -eq 'Yes' -and -not $chart.Export((Join-Path $directoryPath ($item[0]+'.png')),'PNG')) {throw 'Chart preview export failed.'}
            }
        }
        for ($index=0;$index -lt 4;$index++) {
            if ($records[$index].SHA256 -ceq $records[$index+4].SHA256) {throw 'Chart points did not change with the plateau switch.'}
        }
        $records | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directoryPath 'ChartCurves.json') -Encoding UTF8
        Write-Output 'CONFIG_CHARTS: 8 series matched current coordinates; all 4 curves change with the switch.'
    } finally {$cell.Formula=$saved;$config.Calculate()}
}
$printAreas=@(Get-WorkbookPrintAreas $target)
$excel=$null;$book=$null;$ownedPid=0;$ownedTicks=0;$records=@()
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=1
    [uint32]$id=0;[void][PlateauVerifyWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $ownedPid=[int]$id;$ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    Write-Output ('OWNED_EXCEL: id='+$ownedPid+'; utcStartTicks='+$ownedTicks)
    $book=$excel.Workbooks.Open($target,0,$false)
    if ($book.ReadOnly) {throw 'Private candidate is read-only.'}
    Verify-ConfigChartCurves $book
    $inputs=InputSnapshot $book
    $reference=''
    for ($run=1;$run -le 3;$run++) {
        $result=[string]$excel.Run("'$($book.Name)'!modTestPerformance.MeasurePerformanceBatch",'Full')
        $parts=@{};foreach ($field in $result -split ([string][char]30)) {$pair=$field -split '=',2;if ($pair.Count -eq 2) {$parts[$pair[0]]=$pair[1]}}
        if ($run -eq 1) {$reference=$parts.fingerprint} elseif ($parts.fingerprint -cne $reference) {throw 'Repeated calculation changed engineering results.'}
        if ((InputSnapshot $book) -cne $inputs) {throw 'Calculation changed inputs or column widths.'}
        $records += [ordered]@{Run=$run;Seconds=[double]::Parse($parts.seconds,[Globalization.CultureInfo]::InvariantCulture);ExactResults=$true;InputsPreserved=$true}
        Write-Output ('USER_CALCULATION: run='+$run+'; seconds='+$parts.seconds)
    }
    $fingerprint=[string]$excel.Run("'$($book.Name)'!modTestPerformance.PerformanceResultsFingerprint")
    $book.Save();$book.Close($false);$book=$null
    Restore-WorkbookPrintAreas $target $printAreas
    $book=$excel.Workbooks.Open($target,0,$true)
    if (([string]$excel.Run("'$($book.Name)'!modTestPerformance.PerformanceResultsFingerprint")) -cne $fingerprint -or (InputSnapshot $book) -cne $inputs) {throw 'Save/reopen changed data.'}
    $before=[string]$excel.Run("'$($book.Name)'!modAutoCADStressExport.Audit02ReadExportSnapshotForTests",$book)
    [void]$excel.Run("'$($book.Name)'!modSolverWorkStats.ResetSectionEquilibriumSolveCount")
    [void]$excel.Run("'$($book.Name)'!modWorkbookCalculation.UpdateSectionPlotForWorkbook",$book)
    $solves=[long]$excel.Run("'$($book.Name)'!modSolverWorkStats.SectionEquilibriumSolveCount")
    $after=[string]$excel.Run("'$($book.Name)'!modAutoCADStressExport.Audit02ReadExportSnapshotForTests",$book)
    if ($solves -ne 0 -or $before -cne $after) {throw 'Passive plot/export solved again or changed the snapshot.'}
    $book.Close($false);$book=$null
    if ((Get-FileHash -LiteralPath $source).Hash -cne $preparation.SourceSHA256) {throw 'Production changed during acceptance.'}
    [ordered]@{SourceSHA256=$preparation.SourceSHA256;ReleaseSHA256=(Get-FileHash -LiteralPath $target).Hash;InputsPreserved=$true;SaveReopenExact=$true;RepeatExact=$true;Measurements=$records;ResultsFingerprint=(TextHash $fingerprint);PassiveSolveCount=$solves;PassiveExportExact=$true;ProductionChanged=$false} | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $directoryPath 'Acceptance.json') -Encoding UTF8
    Write-Output 'ACCEPTANCE: repeat 3/3 exact; inputs preserved; save/reopen exact; passive solves=0.'
} finally {
    if ($book) {$book.Close($false)}
    if ($excel) {$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if ($ownedPid) {
        $owned=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue
        if ($owned -and $owned.ProcessName -eq 'EXCEL' -and $owned.StartTime.ToUniversalTime().Ticks -eq $ownedTicks) {Stop-Process -Id $ownedPid -Force}
    }
}
