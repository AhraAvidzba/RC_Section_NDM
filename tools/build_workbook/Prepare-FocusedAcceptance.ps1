# Creates an isolated probe; does not execute the acceptance calculations.
param(
    [string]$SourceWorkbook='workbook/output/RC_Section_NDM.xlsm',
    [string]$OutputDirectory='docs/regression/FocusedAcceptance_2026-10-10',
    [string]$ExtraHelperFile='',
    [string]$SavedFixture=''
)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$sourcePath=$SourceWorkbook
if (-not [IO.Path]::IsPathRooted($sourcePath)) {$sourcePath=Join-Path $root $sourcePath}
$source=(Resolve-Path -LiteralPath $sourcePath).Path
$sourceHash=(Get-FileHash -LiteralPath $source).Hash
$directory=Join-Path $root $OutputDirectory
$rootFull=[IO.Path]::GetFullPath($root).TrimEnd('\')+'\'
if (-not [IO.Path]::GetFullPath($directory).StartsWith($rootFull,[StringComparison]::OrdinalIgnoreCase)) {
    throw 'Probe output must stay within the workspace.'
}
New-Item -ItemType Directory -Path $directory -Force | Out-Null
$path=Join-Path $directory 'RC_Section_NDM.xlsm'
if (Test-Path -LiteralPath $path) { throw 'Probe already exists; choose a new directory.' }
Copy-Item -LiteralPath $source -Destination $path
$helper=Join-Path $PSScriptRoot 'FocusedAcceptanceProbe.bas'
$tokens=$null; $errors=$null
$comparisonAst=[Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'Verify-PostAudit03VbaEquality.ps1'),[ref]$tokens,[ref]$errors)
if ($errors.Count) { throw 'VBA comparator parse failed.' }
foreach ($command in $comparisonAst.FindAll({param($node) $node -is [Management.Automation.Language.CommandAst] -and $node.GetCommandName() -eq 'Add-Type'},$true)) {
    . ([scriptblock]::Create($command.Extent.Text))
}
foreach ($definition in $comparisonAst.FindAll({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst]},$true)) {
    . ([scriptblock]::Create($definition.Extent.Text))
}
$excel=$null; $book=$null; $ownedPid=0; $ownedTicks=0
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false; $excel.AutomationSecurity=3
    Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class FocusedPrepareWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId); }'
    [uint32]$id=0
    [void][FocusedPrepareWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $ownedPid=[int]$id; $ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    Write-Output ('OWNED_EXCEL: id='+$ownedPid+'; utcStartTicks='+$ownedTicks)
    $book=$excel.Workbooks.Open($path,0,$false)
    if ($book.ReadOnly) { throw 'Probe is read-only.' }
    $fixtureHash=''
    if ($SavedFixture) {
        $fixture=Get-Content -LiteralPath $SavedFixture -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach($table in $fixture.Tables) {
            $anchor=$book.Names.Item($table.Name).RefersToRange.Cells.Item(1,1)
            $anchor.CurrentRegion.ClearContents()
            $values=New-Object 'object[,]' ([int]$table.Rows),([int]$table.Columns)
            for($r=0;$r -lt $table.Rows;$r++){for($c=0;$c -lt $table.Columns;$c++){$values[$r,$c]=$table.Values[$r][$c]}}
            $anchor.Resize([int]$table.Rows,[int]$table.Columns).Value2=$values
        }
        $fixtureHash=(Get-FileHash -LiteralPath $SavedFixture).Hash
    }
    $before=@{}
    foreach ($component in $book.VBProject.VBComponents) {
        $module=$component.CodeModule
        if ($module.CountOfLines -gt 0) { $before[$component.Name]=$module.Lines(1,$module.CountOfLines) }
    }
    $module=$book.VBProject.VBComponents.Item('modTestBatchCalculation').CodeModule
    $body=$module.Lines(1,$module.CountOfLines)
    $declaration='Option Explicit'
    if ([regex]::Matches($body,[regex]::Escape($declaration)).Count -ne 1) { throw 'Test declaration target mismatch.' }
    $numberDeclarations=@'
Private mFocusedSection As CSectionModel
Private mFocusedBatch As CBatchSectionCalculator
Private mFocusedBinaryDump As Boolean
Private Type TFocusedDouble
    Value As Double
End Type
Private Type TFocusedWords
    Low As Long
    High As Long
End Type
'@
    $body=$body.Replace($declaration,$declaration+"`r`n"+$numberDeclarations)
    $module.DeleteLines(1,$module.CountOfLines); $module.AddFromString($body)
    $module.AddFromString([IO.File]::ReadAllText($helper,[Text.Encoding]::UTF8))
    $extraHash=''
    if ($ExtraHelperFile) {
        $extra=Join-Path $PSScriptRoot $ExtraHelperFile
        $module.AddFromString([IO.File]::ReadAllText($extra,[Text.Encoding]::UTF8))
        $extraHash=(Get-FileHash -LiteralPath $extra).Hash
    }
    $module=$book.VBProject.VBComponents.Item('modWorkbookCalculation').CodeModule
    $body=$module.Lines(1,$module.CountOfLines)
    $target='    batch.Execute'
    if ([regex]::Matches($body,[regex]::Escape($target)).Count -ne 1) { throw 'Calculation capture target mismatch.' }
    $body=$body.Replace($target,$target+"`r`n    modTestBatchCalculation.CaptureFocusedCalculation section, batch")
    $module.DeleteLines(1,$module.CountOfLines); $module.AddFromString($body)
    $modified=@()
    foreach ($component in $book.VBProject.VBComponents) {
        $module=$component.CodeModule
        if ($module.CountOfLines -gt 0 -and
            (ConvertTo-VbaCodeCase $before[$component.Name]) -cne (ConvertTo-VbaCodeCase $module.Lines(1,$module.CountOfLines))) {
            $modified+=$component.Name
        }
    }
    if ($modified.Count -ne 2 -or 'modTestBatchCalculation' -notin $modified -or 'modWorkbookCalculation' -notin $modified) {
        throw ('Unexpected component changed in test preparation: '+($modified -join ', '))
    }
    $excel.VBE.ActiveVBProject=$book.VBProject
    $compile=$excel.VBE.CommandBars.FindControl(1,578)
    if ($compile.Enabled) { $compile.Execute() }
    if ($compile.Enabled) { throw 'Focused probe compilation failed.' }
    $book.Save(); $book.Close($false); $book=$null
    if ((Get-FileHash -LiteralPath $source).Hash -cne $sourceHash) { throw 'Source workbook changed.' }
    [ordered]@{Created=(Get-Date).ToString('o'); SourceWorkbook=$source; SourceSHA256=$sourceHash;
        ProbeWorkbookSHA256=(Get-FileHash -LiteralPath $path).Hash; HelperSHA256=(Get-FileHash -LiteralPath $helper).Hash;
        ExtraHelperSHA256=$extraHash; SavedFixtureSHA256=$fixtureHash; Compiled=$true; ModifiedComponents=$modified; CalculationExecuted=$false; ProductionChanged=$false} |
        ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $directory 'Preparation.json') -Encoding UTF8
    Write-Output ('FOCUSED_PROBE_READY: '+$directory+'; calculation not executed')
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($ownedPid) {
        $owned=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue
        if ($owned -and $owned.ProcessName -eq 'EXCEL' -and $owned.StartTime.ToUniversalTime().Ticks -eq $ownedTicks) {
            Stop-Process -Id $ownedPid -Force -ErrorAction SilentlyContinue
        }
    }
}
