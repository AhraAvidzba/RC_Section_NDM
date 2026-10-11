# Updates a private copy only. Existing inputs, Results and validation sources are preserved.
param([string]$Directory='docs/regression/PlateauSlope_2026-10-11/Release', [switch]$NativeReference)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
$source=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$expected='DA10AE486C920E114E55B28F7C71188B5C51E2E876402AB05022A4EFFCDB2D4F'
if ((Get-FileHash -LiteralPath $source).Hash -cne $expected) {throw 'The user workbook changed; reassess before preparing a release.'}
$directoryPath=[IO.Path]::GetFullPath((Join-Path $root $Directory))
$allowed=[IO.Path]::GetFullPath((Join-Path $root 'docs/regression/PlateauSlope_2026-10-11'))+[IO.Path]::DirectorySeparatorChar
if (-not $directoryPath.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)) {throw 'Release must remain in its evidence directory.'}
New-Item -ItemType Directory -Path $directoryPath -ErrorAction Stop | Out-Null
$target=Join-Path $directoryPath 'RC_Section_NDM.xlsm'
Copy-Item -LiteralPath $source -Destination $target
$printAreas=@(Get-WorkbookPrintAreas $target)
Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class PlateauReleaseWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId); }'
function TextHash([string]$value) {
    $sha=[Security.Cryptography.SHA256]::Create()
    try {return [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($value))).Replace('-','')}
    finally {$sha.Dispose()}
}
function PreservedSnapshot($book) {
    $config=$book.Worksheets.Item('Config'); $used=$config.UsedRange
    $system=$book.Names.Item('rngSystemSettings').RefersToRange
    $values=$system.Formula; $entries=@()
    for ($row=1;$row -le $system.Rows.Count;$row++) {
        if ([string]$values[$row,1] -eq 'Solver.PlateauSlope.Enabled') {continue}
        $entry=@();for ($column=1;$column -le 4;$column++) {$entry += $values[$row,$column]}
        $entries += ,$entry
    }
    $results=$book.Worksheets.Item('Results'); $widths=@()
    for ($column=1;$column -le $results.UsedRange.Columns.Count;$column++) {$widths += [double]$results.Columns.Item($column).ColumnWidth}
    $names=@();foreach ($name in $book.Names) {
        if ([string]$name.Name -notmatch '(^|!)rngSystemSettings$|Print_Area') {$names += [string]$name.Name+'|'+[string]$name.RefersTo}
    }
    $right=$config.Range($config.Cells.Item(1,6),$config.Cells.Item($used.Row+$used.Rows.Count-1,$used.Column+$used.Columns.Count-1))
    $rightData=$right.Formula
    $title=$config.Cells.Find('Контрольные точки диаграмм*',$config.Cells.Item(1,1),-4163,1,1,1,$false,$false,$false)
    if ($null -eq $title) {throw 'Material control table is missing.'}
    # Only the 52 informational coordinate pairs may change, never material inputs.
    for ($row=1;$row -le 52;$row++) {
        for ($column=4;$column -le 5;$column++) {
            $rightData[($title.Row+1+$row),($title.Column+$column-5)]='<material coordinate>'
        }
    }
    return TextHash (([ordered]@{System=$entries;Right=$rightData;Results=$results.UsedRange.Formula;Widths=$widths;Names=@($names|Sort-Object)} | ConvertTo-Json -Depth 8 -Compress))
}
function Export-Code($book,[string]$path) {
    $writer=[IO.StreamWriter]::new($path,$false,[Text.UTF8Encoding]::new($true))
    try {
        $writer.WriteLine('VBA PROJECT EXPORT');$writer.WriteLine('Workbook: '+$book.Name);$writer.WriteLine('Date: '+(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
        foreach ($component in $book.VBProject.VBComponents) {
            $module=$component.CodeModule
            $writer.WriteLine('='*100);$writer.WriteLine('COMPONENT: '+$component.Name);$writer.WriteLine('TYPE: '+$component.Type);$writer.WriteLine('LINES: '+$module.CountOfLines);$writer.WriteLine('='*100)
            if ($module.CountOfLines) {$writer.WriteLine($module.Lines(1,$module.CountOfLines))}
            $writer.WriteLine()
        }
    } finally {$writer.Dispose()}
}
$excel=$null;$book=$null;$ownedPid=0;$ownedTicks=0
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3
    [uint32]$id=0;[void][PlateauReleaseWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $ownedPid=[int]$id;$ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    Write-Output ('OWNED_EXCEL: id='+$ownedPid+'; utcStartTicks='+$ownedTicks)
    $book=$excel.Workbooks.Open($target,0,$false)
    if ($book.ReadOnly) {throw 'Private release is read-only.'}
    $before=PreservedSnapshot $book
    $sourceFiles=@(Get-ChildItem -LiteralPath (Join-Path $root 'src'),(Join-Path $root 'tests') -Recurse -File | Where-Object Extension -in '.bas','.cls')
    $changed=@('CMaterialModelProvider','modTestMaterialDiagrams','modTestWorkbookInterface','modTestBatchCalculation')
    if ($NativeReference) {$changed=@($sourceFiles.BaseName)}
    foreach ($name in $changed) {
        $files=@($sourceFiles|Where-Object BaseName -eq $name)
        if ($files.Count -ne 1) {throw 'Ambiguous source module.'}
        $body=[regex]::Match([IO.File]::ReadAllText($files[0].FullName,[Text.Encoding]::UTF8),'(?ms)^Option Explicit.*').Value
        if (-not $body) {throw 'Missing source body.'}
        $module=$book.VBProject.VBComponents.Item($name).CodeModule
        $module.DeleteLines(1,$module.CountOfLines);$module.AddFromString(($body -split '\r?\n') -join "`r`n")
    }
    foreach ($component in $book.VBProject.VBComponents) {
        if ([string]$component.Name -like '*Plateau*Probe*') {throw 'Experimental instrumentation must not enter the release.'}
    }
    $setting=Update-SolverPlateauSlopeSetting $book
    $config=$book.Worksheets.Item('Config')
    $title=$config.Cells.Find('Контрольные точки диаграмм*',$config.Cells.Item(1,1),-4163,1,1,1,$false,$false,$false)
    Add-MaterialDiagramControlTables $config ($title.Row+1) $title.Column -CoordinatesOnly
    $config.Calculate()
    if ((PreservedSnapshot $book) -cne $before) {throw 'Preparation changed previous inputs, Results, named ranges or widths.'}
    $excel.VBE.ActiveVBProject=$book.VBProject
    $compile=$excel.VBE.CommandBars.FindControl(1,578)
    if ($null -eq $compile) {throw 'Compile command is unavailable.'}
    if ($compile.Enabled) {$compile.Execute()}
    if ($compile.Enabled) {throw 'VBA compilation failed.'}
    $book.Save();$book.Close($false);$book=$null
    Restore-WorkbookPrintAreas $target $printAreas
    $book=$excel.Workbooks.Open($target,0,$true)
    if ((PreservedSnapshot $book) -cne $before) {throw 'Save/reopen changed preserved data.'}
    Export-Code $book (Join-Path $directoryPath 'VBA_All_Code.txt')
    $book.Close($false);$book=$null
    if ((Get-FileHash -LiteralPath $source).Hash -cne $expected) {throw 'Production changed during preparation.'}
    [ordered]@{SourceSHA256=$expected;ReleaseSHA256=(Get-FileHash -LiteralPath $target).Hash;Compiled=$true;PreservedSnapshot=$before;InputsResultsNamesWidthsPreserved=$true;SaveReopenExact=$true;ChangedComponents=$changed;NativeReference=[bool]$NativeReference;Setting=$setting;ProductionChanged=$false} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directoryPath 'Preparation.json') -Encoding UTF8
    Write-Output 'PRIVATE_RELEASE_READY: compiled; user data preserved; production unchanged.'
} finally {
    if ($book) {$book.Close($false)}
    if ($excel) {$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if ($ownedPid) {
        $owned=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue
        if ($owned -and $owned.ProcessName -eq 'EXCEL' -and $owned.StartTime.ToUniversalTime().Ticks -eq $ownedTicks) {Stop-Process -Id $ownedPid -Force}
    }
}
