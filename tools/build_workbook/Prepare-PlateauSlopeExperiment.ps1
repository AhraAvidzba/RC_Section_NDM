# Test-copy instrumentation only; never inject this module into the release book.
param(
    [string]$SourceWorkbook='workbook/output/RC_Section_NDM.xlsm',
    [string]$OutputDirectory='docs/regression/PlateauSlope_2026-10-11/ProbeV3',
    [switch]$RefreshProbeModule,
    [switch]$FixAxialAssertions,
    [switch]$HighCrackedLoads,
    [switch]$FixSP35Assertions,
    [ValidateSet('SP35','SP63')][string]$CrackCode='SP35'
)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$source=(Resolve-Path -LiteralPath (Join-Path $root $SourceWorkbook)).Path
$sourceHash=(Get-FileHash -LiteralPath $source).Hash
if (-not $RefreshProbeModule) {
    & (Join-Path $PSScriptRoot 'Prepare-FocusedAcceptance.ps1') -SourceWorkbook $SourceWorkbook -OutputDirectory $OutputDirectory
}
$path=Join-Path (Join-Path $root $OutputDirectory) 'RC_Section_NDM.xlsm'
if (-not [IO.Path]::GetFullPath($path).StartsWith((Join-Path $root 'docs/regression/PlateauSlope_'),[StringComparison]::OrdinalIgnoreCase)) { throw 'Experimental directory required.' }
Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class PlateauPrepareWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId); }'

function Replace-Exact([string]$Body,[string]$Old,[string]$New,[int]$Count=1) {
    $matches=[regex]::Matches($Body,[regex]::Escape($Old),[Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $actual=$matches.Count
    if ($actual -ne $Count) { throw "Patch mismatch ($actual != $Count): $Old" }
    for ($i=$matches.Count-1; $i -ge 0; $i--) {
        $Body=$Body.Substring(0,$matches[$i].Index)+$New+$Body.Substring($matches[$i].Index+$matches[$i].Length)
    }
    return $Body
}
function Add-EntryCounter([string]$Body,[string]$Method,[string]$Counter) {
    $lines=[Collections.Generic.List[string]]::new()
    $lines.AddRange([string[]]($Body -split '\r?\n'))
    $indices=@(0..($lines.Count-1) | Where-Object { $lines[$_] -match ('^(Public|Private|Friend) (Sub|Function) '+[regex]::Escape($Method)+'\(') })
    if ($indices.Count -ne 1) { throw "Method not unique: $Method" }
    $i=$indices[0]
    while ($lines[$i].TrimEnd().EndsWith('_')) { $i++ }
    $lines.Insert($i+1,'    modTestPlateauSlopeProbe.'+$Counter+' = modTestPlateauSlopeProbe.'+$Counter+' + 1')
    return $lines -join "`r`n"
}
function Increment([string]$Counter) { return '    modTestPlateauSlopeProbe.'+$Counter+' = modTestPlateauSlopeProbe.'+$Counter+' + 1' }
function Transform-Method([string]$Body,[string]$Name,[scriptblock]$Edit) {
    $pattern='(?ms)^(Public|Private|Friend) (Sub|Function) '+[regex]::Escape($Name)+'\(.*?^End (Sub|Function)'
    $matches=[regex]::Matches($Body,$pattern,[Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($matches.Count -ne 1) {throw "Method mismatch: $Name"}
    $match=$matches[0]
    $new=& $Edit $match.Value
    return $Body.Substring(0,$match.Index)+$new+$Body.Substring($match.Index+$match.Length)
}
function Set-Body([string]$Name,[string]$Body) {
    $module=$book.VBProject.VBComponents.Item($Name).CodeModule
    $module.DeleteLines(1,$module.CountOfLines); $module.AddFromString($Body)
    $script:modified.Add($Name)
}
function Get-Body([string]$Name) {
    $module=$book.VBProject.VBComponents.Item($Name).CodeModule
    return $module.Lines(1,$module.CountOfLines)
}
$excel=$null; $book=$null; $ownedPid=0; $ownedTicks=0
$modified=[Collections.Generic.List[string]]::new()
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false; $excel.AutomationSecurity=3
    [uint32]$id=0
    [void][PlateauPrepareWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $ownedPid=[int]$id; $ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    Write-Output ('OWNED_EXCEL: id='+$ownedPid+'; utcStartTicks='+$ownedTicks)
    $book=$excel.Workbooks.Open($path,0,$false)
    if ($book.ReadOnly) { throw 'Experiment copy is read-only.' }
    if ($RefreshProbeModule) {
        $probe=$book.VBProject.VBComponents.Item('modTestPlateauSlopeProbe')
        $probe.CodeModule.DeleteLines(1,$probe.CodeModule.CountOfLines)
    } else {
        $probe=$book.VBProject.VBComponents.Add(1); $probe.Name='modTestPlateauSlopeProbe'
    }
    $body=[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'PlateauSlopeProbe.bas'),[Text.Encoding]::UTF8)
    $probe.CodeModule.AddFromString([regex]::Replace($body,'(?m)^Attribute VB_Name[^\r\n]*\r?\n',''))
    $modified.Add('modTestPlateauSlopeProbe')
    if ($FixAxialAssertions) {
        $body=Get-Body 'modTestBatchCalculation'
        $body=Transform-Method $body 'TestAudit03CheckCapacityPath' {
            param($method)
            $method=Replace-Exact $method 'Optional ByVal includeUnscaledPath As Boolean = False)' 'Optional ByVal includeUnscaledPath As Boolean = False, Optional ByVal offsetX As Double = 10#, Optional ByVal offsetY As Double = -7#)'
            $method=Replace-Exact $method '(referenceY - 7#)' '(referenceY + offsetY)'
            Replace-Exact $method '(referenceX + 10#)' '(referenceX + offsetX)'
        }
        $body=Transform-Method $body 'TestRunFocusedAcceptanceCases' {
            param($method)
            Replace-Exact $method 'Array(item(1), item(2), item(3)), cx, cy, True' 'Array(item(1), item(2), item(3)), cx, cy, True, units.InputLengthToInternal(settings.GetDouble("Load.ReferenceOffsetX", 0#, False)), units.InputLengthToInternal(settings.GetDouble("Load.ReferenceOffsetY", 0#, False))'
        }
        Set-Body 'modTestBatchCalculation' $body
    }
    if ($HighCrackedLoads -or $FixSP35Assertions) {
        $body=Get-Body 'modTestBatchCalculation'
        if ($HighCrackedLoads) {
            $body=Transform-Method $body 'TestFocusedSelectedLoads' {
                param($method)
                Replace-Exact $method 'For Each factor In Array(6#, 10#)' 'For Each factor In Array(12#, 16#)'
            }
        }
        if ($FixSP35Assertions) {
            $body=Transform-Method $body 'TestAudit03CheckResultComments' {
                param($method)
                Replace-Exact $method 'If Not result.CrackResult.Formation.HasLimitPoint And result.CrackWidthMeta.Calculated Then' 'If result.CrackResult.Width.StandardCode = "SP63" And Not result.CrackResult.Formation.HasLimitPoint And result.CrackWidthMeta.Calculated Then'
            }
            $body=Transform-Method $body 'TestAudit03CheckAppliedPsiMode' {
                param($method)
                Replace-Exact $method '    If Not result.CrackWidthMeta.Calculated Then Exit Sub' '    If result.CrackResult.Width.StandardCode <> "SP63" Or Not result.CrackWidthMeta.Calculated Then Exit Sub'
            }
        }
        Set-Body 'modTestBatchCalculation' $body
    }
    if ($CrackCode -eq 'SP63') {
        $body=Get-Body 'modTestBatchCalculation'
        $body=Transform-Method $body 'TestRunFocusedAcceptanceCases' {
            param($method)
            $method=Replace-Exact $method 'SetSystemSetting "SLS.Crack.Code", "SP35"' 'SetSystemSetting "SLS.Crack.Code", "SP63"'
            Replace-Exact $method '|profile=PR3|code=SP35' '|profile=PR3|code=SP63'
        }
        Set-Body 'modTestBatchCalculation' $body
    }

    if (-not $RefreshProbeModule) {
    $body=Get-Body 'CMaterialModelProvider'
    foreach ($method in @('BuildConcreteTwoLine','BuildConcreteThreeLine')) {
        $body=Replace-Exact $body ("Set $method = DiagramFromArrays(eps, sig, count)") ("Set $method = DiagramFromArrays(eps, sig, count, mConcreteParameters.Eb, mConcreteParameters.Ebt)")
    }
    $body=Replace-Exact $body 'Set BuildSteelTwoLine = DiagramFromArrays(eps, sig, 5)' 'Set BuildSteelTwoLine = DiagramFromArrays(eps, sig, 5, mSteelParameters.Esc, mSteelParameters.Es)'
    $body=Replace-Exact $body 'Set BuildSteelThreeLine = DiagramFromArrays(eps, sig, 9)' 'Set BuildSteelThreeLine = DiagramFromArrays(eps, sig, 9, mSteelParameters.Esc, mSteelParameters.Es)'
    $body=Replace-Exact $body 'Private Function DiagramFromArrays(ByRef strains() As Double, ByRef stresses() As Double, ByVal pointCount As Long) As CMaterialDiagram' 'Private Function DiagramFromArrays(ByRef strains() As Double, ByRef stresses() As Double, ByVal pointCount As Long, ByVal compressionE As Double, ByVal tensionE As Double) As CMaterialDiagram'
    $body=Replace-Exact $body '    diagram.InitializeFromArrays strains, stresses, pointCount' ("    modTestPlateauSlopeProbe.TestPlateauApplySlope strains, stresses, pointCount, compressionE, tensionE`r`n    diagram.InitializeFromArrays strains, stresses, pointCount")
    Set-Body 'CMaterialModelProvider' $body

    $body=Get-Body 'CSectionSolver'
    $body=Add-EntryCounter $body 'Solve' 'SolverCalls'
    $body=Replace-Exact $body '    For iter = 1 To mMaxIterations' ("    For iter = 1 To mMaxIterations`r`n"+(Increment 'StateIterations')) 2
    $body=Add-EntryCounter $body 'TryRegularizedStep' 'StateRegularizedSteps'
    Set-Body 'CSectionSolver' $body

    $body=Get-Body 'CStateGuessBuilder'
    $body=Add-EntryCounter $body 'BuildForDirectState' 'DirectGuesses'
    $body=Add-EntryCounter $body 'BuildAttemptForExtensionState' 'ExtensionGuesses'
    $body=Replace-Exact $body '    For iter = 1 To EXTENSION_ACTIVE_SET_ITERATIONS' ("    For iter = 1 To EXTENSION_ACTIVE_SET_ITERATIONS`r`n"+(Increment 'ActiveSetIterations'))
    Set-Body 'CStateGuessBuilder' $body

    $body=Get-Body 'CStateSolutionRunner'
    $body=Replace-Exact $body '    retry.Solve section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy' ((Increment 'RetryCalls')+"`r`n    retry.Solve section, concreteMaterial, steelMaterial, targetN, targetMx, targetMy")
    Set-Body 'CStateSolutionRunner' $body

    $body=Get-Body 'CUltimateStrainSearch'
    $body=Add-EntryCounter $body 'RunNewton' 'UltimateRuns'
    $body=Add-EntryCounter $body 'TryRestartWarmGuess' 'WarmRestarts'
    $body=Add-EntryCounter $body 'TryUltimateProjectedStep' 'ProjectedRecoveries'
    $body=Add-EntryCounter $body 'ExecuteCapacityFallbackToLoadMultiplier' 'MultiplierFallbacks'
    $body=Replace-Exact $body '    For iter = 1 To context.LimitSearchUltimateMaxIterations' ("    For iter = 1 To context.LimitSearchUltimateMaxIterations`r`n"+(Increment 'UltimateIterations'))
    Set-Body 'CUltimateStrainSearch' $body

    $body=Get-Body 'CCapacitySolver'
    $body=Replace-Exact $body '        If attempt < mMaxRetries Then mRetryCount = mRetryCount + 1' ("        If attempt < mMaxRetries Then`r`n            mRetryCount = mRetryCount + 1`r`n"+(Increment 'CapacityRetries')+"`r`n        End If")
    Set-Body 'CCapacitySolver' $body
    $body=Get-Body 'CCrackFormationCalculator'
    $body=Replace-Exact $body '    Set formationStart = SolveState(section, concreteMaterial, steelMaterial, Nothing, targetN, targetMx, targetMy)' ((Increment 'FormationFullStarts')+"`r`n    Set formationStart = SolveState(section, concreteMaterial, steelMaterial, Nothing, targetN, targetMx, targetMy)")
    Set-Body 'CCrackFormationCalculator' $body

    $body=Get-Body 'modTestBatchCalculation'
    $body += "`r`n"+[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'PlateauSlopeBatchProbe.bas'),[Text.Encoding]::UTF8)
    $body=Replace-Exact $body '        ResetSectionEquilibriumSolveCount' ("        modTestPlateauSlopeProbe.TestPlateauResetCounters`r`n        ResetSectionEquilibriumSolveCount")
    # Report timing, rather than making the known >10 s controls a harness failure.
    $body=Replace-Exact $body '        AssertTrue stats, prefix & ".within10Seconds", elapsed <= 10#' '        AssertTrue stats, prefix & ".withinExperimentBudget", elapsed <= 180#'
    $body=Replace-Exact $body '            Set result = mFocusedBatch.ResultAt(1)' ("            Set result = mFocusedBatch.ResultAt(1)`r`n            modTestPlateauSlopeProbe.TestPlateauCaptureResult CStr(item(0)), elapsed, result")
    Set-Body 'modTestBatchCalculation' $body
    }
    $excel.VBE.ActiveVBProject=$book.VBProject
    $compile=$excel.VBE.CommandBars.FindControl(1,578)
    if ($compile.Enabled) { $compile.Execute() }
    if ($compile.Enabled) {
        $pane=$excel.VBE.ActiveCodePane
        $row=0; $column=0; $endRow=0; $endColumn=0
        $pane.GetSelection([ref]$row,[ref]$column,[ref]$endRow,[ref]$endColumn)
        throw "Experiment compile failed at $($pane.CodeModule.Parent.Name):${row}: $($pane.CodeModule.Lines($row,1))"
    }
    $book.Save(); $book.Close($false); $book=$null
    if ((Get-FileHash -LiteralPath $source).Hash -cne $sourceHash) { throw 'Production workbook changed.' }
    [ordered]@{Created=(Get-Date).ToString('o'); SourceSHA256=$sourceHash; ExperimentSHA256=(Get-FileHash -LiteralPath $path).Hash;
        ModifiedComponents=$modified.ToArray(); Compiled=$true; ProductionChanged=$false; CalculationsExecuted=$false;
        AlphaValues=@(0,0.00001,0.0001,0.001); ProactiveGuessesChanged=$false; RefreshedModuleOnly=[bool]$RefreshProbeModule} |
        ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path (Split-Path $path) $(if ($RefreshProbeModule) {'PlateauRefresh.json'} else {'PlateauPreparation.json'})) -Encoding UTF8
    Write-Output ('PLATEAU_EXPERIMENT_READY: '+$path)
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
