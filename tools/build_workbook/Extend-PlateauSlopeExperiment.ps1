# Adds exact-axial and batch controls only to an isolated copy.
param(
    [string]$SourceProbe='docs/regression/PlateauSlope_2026-10-11/ProbeV3/RC_Section_NDM.xlsm',
    [string]$OutputDirectory='docs/regression/PlateauSlope_2026-10-11/ProbeV4a'
)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$production=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$hash=(Get-FileHash -LiteralPath $production).Hash
$directory=Join-Path $root $OutputDirectory
if (-not [IO.Path]::GetFullPath($directory).StartsWith((Join-Path $root 'docs/regression/PlateauSlope_'),[StringComparison]::OrdinalIgnoreCase)) { throw 'Experimental directory required.' }
$path=Join-Path $directory 'RC_Section_NDM.xlsm'
if (Test-Path -LiteralPath $path) { throw 'Destination already exists.' }
New-Item -ItemType Directory -Path $directory -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $root $SourceProbe) -Destination $path
$tokens=$null; $errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'Prepare-PlateauSlopeExperiment.ps1'),[ref]$tokens,[ref]$errors)
if ($errors.Count) { throw 'Recipe parser errors.' }
foreach ($definition in $ast.FindAll({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst]},$true)) {
    . ([scriptblock]::Create($definition.Extent.Text))
}
Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class PlateauExtendWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId); }'
$excel=$null; $book=$null; $ownedPid=0; $ownedTicks=0; $modified=[Collections.Generic.List[string]]::new()
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false; $excel.AutomationSecurity=3
    [uint32]$id=0; [void][PlateauExtendWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $ownedPid=[int]$id; $ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    Write-Output ('OWNED_EXCEL: id='+$ownedPid+'; utcStartTicks='+$ownedTicks)
    $book=$excel.Workbooks.Open($path,0,$false)
    if ($book.ReadOnly) { throw 'Experiment copy is read-only.' }
    $body=Get-Body 'modTestBatchCalculation'
    $body=Transform-Method $body 'RunFocusedAcceptanceCases' {
        param($method)
        Replace-Exact $method '    settings.LoadFromWorkbook ThisWorkbook' ("    If groupName = ""PureAxial"" Then`r`n        SetSystemSetting ""Load.ReferenceOffsetX"", ""0""`r`n        SetSystemSetting ""Load.ReferenceOffsetY"", ""0""`r`n    End If`r`n    settings.LoadFromWorkbook ThisWorkbook")
    }
    $marker='    ElseIf groupName = "Generated" Or groupName = "Saved" Then'
    $branch=@'
    ElseIf groupName = "Cracked" Then
        Dim factor As Variant
        For Each item In Audit03NormalizedLoads(section, "Light")
            If FocusedNameIn(CStr(item(0)), Array("Mx-1", "Mx1", "My-1", "My1", "mixedC1_-1", "mixedT-1_1")) Then
                For Each factor In Array(6#, 10#)
                    result.Add Array("cracked_" & CStr(item(0)) & "_x" & CStr(factor), item(1), _
                        CDbl(factor) * CDbl(item(2)), CDbl(factor) * CDbl(item(3)), "Auto")
                Next factor
            End If
        Next item
    ElseIf groupName = "PureAxial" Then
        For Each item In Audit03NormalizedLoads(section, "Light")
            If FocusedNameIn(CStr(item(0)), Array("axialSmall-1", "axialSmall1")) Then _
                result.Add Array(item(0), item(1), 0#, 0#, "Auto")
        Next item
        result.Add Array("N500", -500# * 9806.65, 0#, 0#, "Auto")
        result.Add Array("N1000", -1000# * 9806.65, 0#, 0#, "Auto")
        result.Add Array("N2098Plateau", -2098.694 * 9806.65, 0#, 0#, "Auto")
'@
    $body=Replace-Exact $body $marker ($branch+"`r`n"+$marker)
    $body=Replace-Exact $body '        one = RunFocusedAcceptanceCases(CStr(parts(2)), CStr(parts(3)), CStr(parts(4)))' ("        If CStr(parts(3)) = ""Batch15"" Then`r`n            one = RunPlateauBatch15()`r`n        Else`r`n            one = RunFocusedAcceptanceCases(CStr(parts(2)), CStr(parts(3)), CStr(parts(4)))`r`n        End If")
    $body += "`r`n"+[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'PlateauSlopeExtendedProbe.bas'),[Text.Encoding]::UTF8)
    Set-Body 'modTestBatchCalculation' $body
    $excel.VBE.ActiveVBProject=$book.VBProject
    $compile=$excel.VBE.CommandBars.FindControl(1,578)
    if ($compile.Enabled) {$compile.Execute()}
    if ($compile.Enabled) {throw 'Extended experiment compilation failed.'}
    $book.Save(); $book.Close($false); $book=$null
    if ((Get-FileHash -LiteralPath $production).Hash -cne $hash) {throw 'Production book changed.'}
    [ordered]@{Created=(Get-Date).ToString('o'); ProductionUnchanged=$true; Compiled=$true;
        ModifiedComponents=$modified.ToArray(); ExperimentSHA256=(Get-FileHash -LiteralPath $path).Hash} |
        ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $directory 'ExtendedPreparation.json') -Encoding UTF8
    Write-Output ('EXTENDED_PLATEAU_PROBE_READY: '+$path)
} finally {
    if ($book) {$book.Close($false)}
    if ($excel) {$excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($ownedPid) {
        $owned=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue
        if ($owned -and $owned.ProcessName -eq 'EXCEL' -and $owned.StartTime.ToUniversalTime().Ticks -eq $ownedTicks) {
            Stop-Process -Id $ownedPid -Force -ErrorAction SilentlyContinue
        }
    }
}
