# Prepares a native release from the tested candidate without rebuilding Config.
param([string]$Directory='docs/regression/OpenIssues_2026-10-11/Release')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
$source=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$expected='3117BC11929039CD64130A79AFA9E76F4FD703FCA73331A58DA269C985683742'
if ((Get-FileHash -LiteralPath $source).Hash -cne $expected) {throw 'The user workbook changed; reassess before preparing a release.'}
$directoryPath=[IO.Path]::GetFullPath((Join-Path $root $Directory))
$allowed=[IO.Path]::GetFullPath((Join-Path $root 'docs/regression/OpenIssues_2026-10-11'))+[IO.Path]::DirectorySeparatorChar
if (-not $directoryPath.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)) {throw 'Unexpected release path.'}
New-Item -ItemType Directory -Path $directoryPath -ErrorAction Stop | Out-Null
$target=Join-Path $directoryPath 'RC_Section_NDM.xlsm'
Copy-Item -LiteralPath (Join-Path $allowed 'Candidate/RC_Section_NDM.xlsm') -Destination $target
$codeUpdate=Get-Content -LiteralPath (Join-Path $allowed 'Candidate/CodeUpdate.json') -Raw | ConvertFrom-Json
if (-not $codeUpdate.Compiled -or $codeUpdate.TraceGeometry -or
    $codeUpdate.WorkbookSHA256 -cne (Get-FileHash -LiteralPath $target).Hash) {throw 'Candidate compilation provenance changed.'}
$printAreas=@(Get-WorkbookPrintAreas $target)
Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class OpenIssuesReleaseWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId); }'
function TextHash([string]$value) {
    $sha=[Security.Cryptography.SHA256]::Create()
    try {return [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($value))).Replace('-','')}
    finally {$sha.Dispose()}
}
function PreservedSnapshot($book) {
    $names=@();foreach ($name in $book.Names) {
        if ([string]$name.Name -notmatch 'Print_Area') {$names += [string]$name.Name+'|'+[string]$name.RefersTo}
    }
    $results=$book.Worksheets.Item('Results');$widths=@()
    for ($column=1;$column -le 260;$column++) {$widths += [double]$results.Columns.Item($column).ColumnWidth}
    return TextHash (([ordered]@{Config=$book.Worksheets.Item('Config').UsedRange.Formula;Results=$results.UsedRange.Formula;Widths=$widths;Names=@($names|Sort-Object)} | ConvertTo-Json -Depth 8 -Compress))
}
function Export-Code($book,[string]$path) {
    $writer=[IO.StreamWriter]::new($path,$false,[Text.UTF8Encoding]::new($true))
    try {
        $writer.WriteLine('VBA PROJECT EXPORT');$writer.WriteLine('Workbook: '+$book.Name);$writer.WriteLine('Date: '+(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
        foreach ($component in $book.VBProject.VBComponents) {
            if ([string]$component.Name -like '*Probe*') {throw 'Experimental component in release.'}
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
    [uint32]$id=0;[void][OpenIssuesReleaseWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $ownedPid=[int]$id;$ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    Write-Output ('OWNED_EXCEL: id='+$ownedPid+'; utcStartTicks='+$ownedTicks)
    $book=$excel.Workbooks.Open($source,0,$true)
    $before=PreservedSnapshot $book
    $book.Close($false);$book=$null
    $book=$excel.Workbooks.Open($target,0,$false)
    if ($book.ReadOnly -or (PreservedSnapshot $book) -cne $before) {throw 'Candidate does not preserve the user workbook.'}
    $guide=$book.Worksheets.Item([string][char]0x421+[char]0x43F+[char]0x440+[char]0x430+[char]0x432+[char]0x43A+[char]0x430)
    $key=$guide.Cells.Find('SLS.Crack.InitiationSolutionStrategy',$guide.Cells.Item(1,1),-4163,1,1,1,$false,$false,$false)
    if ($null -eq $key) {throw 'Strategy help is missing.'}
    $cell=$guide.Cells.Item($key.Row+5,2)
    if ([string]$cell.Value2 -notlike 'Auto*UltimateStrain*LoadMultiplier*') {throw 'Strategy paragraph layout changed.'}
    $lines=@(Get-SettingInstructionLines 'SLS.Crack.InitiationSolutionStrategy' '')
    $cell.Value2=[string]$lines[-1];$cell.WrapText=$true;$cell.EntireRow.RowHeight=66
    $values=$guide.UsedRange.Value2;$matches=@()
    for ($r=1;$r -le $values.GetLength(0);$r++) {
        if ([string]$values[$r,1] -like 'ResultComment *') {$matches += $guide.UsedRange.Row+$r-1}
    }
    if ($matches.Count -ne 1) {throw 'Results guide is missing or ambiguous.'}
    $row=$matches[0]+1
    if ($guide.Cells.Item($row,1).Value2) {throw 'No free row for counter help.'}
    $paragraph=$guide.Range($guide.Cells.Item($row,1),$guide.Cells.Item($row,6))
    $paragraph.Merge();$paragraph.NumberFormat='@';$paragraph.Value2=[string](Get-BatchSummaryInstruction)
    $paragraph.WrapText=$true;$paragraph.Font.Bold=$false;$paragraph.Font.Size=11;$paragraph.VerticalAlignment=-4160;$paragraph.EntireRow.RowHeight=50
    if ((PreservedSnapshot $book) -cne $before) {throw 'Release preparation changed data or named ranges.'}
    # The code is the already compiled candidate; only help cells were changed.
    $book.Save();$book.Close($false);$book=$null
    Restore-WorkbookPrintAreas $target $printAreas
    $book=$excel.Workbooks.Open($target,0,$true)
    if ((PreservedSnapshot $book) -cne $before) {throw 'Save/reopen changed preserved data.'}
    Export-Code $book (Join-Path $directoryPath 'VBA_All_Code.txt')
    $book.Close($false);$book=$null
    if ((Get-FileHash -LiteralPath $source).Hash -cne $expected) {throw 'Production changed during preparation.'}
    [ordered]@{SourceSHA256=$expected;ReleaseSHA256=(Get-FileHash -LiteralPath $target).Hash;Compiled=$true;PreservedSnapshot=$before;InputsResultsNamesWidthsPreserved=$true;SaveReopenExact=$true;ChangedComponents=@();ProductionChanged=$false} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directoryPath 'Preparation.json') -Encoding UTF8
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
