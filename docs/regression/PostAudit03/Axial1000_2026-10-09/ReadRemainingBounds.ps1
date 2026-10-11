param([string]$SourceDirectory='CandidateV16', [string]$OutputDirectory='RemainingBounds', [string]$ShapeName='SavedWorkbook',
    [string]$Component='modTestBatchCalculation', [string]$DiagnosticFile='RemainingBounds.bas', [string]$Macro='DumpRemainingFailureBounds',
    [switch]$TryRepresentableBracketRecovery)
$ErrorActionPreference = 'Stop'
$directory = Join-Path $PSScriptRoot $OutputDirectory
New-Item -ItemType Directory -Path $directory -Force | Out-Null
$source = Join-Path $PSScriptRoot ($SourceDirectory+'/RC_Section_NDM.xlsm')
$sourceHash = (Get-FileHash -LiteralPath $source).Hash
$bookPath = Join-Path $directory 'RC_Section_NDM.xlsm'
Copy-Item -LiteralPath $source -Destination $bookPath -Force
$excel = $null; $book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class RemainingBoundsWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId); }'
    [uint32]$ownedPid = 0
    [void][RemainingBoundsWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$ownedPid)
    $identity = Get-Process -Id $ownedPid
    Write-Output ('OWNED_EXCEL: id=' + $ownedPid + '; utcStartTicks=' + $identity.StartTime.ToUniversalTime().Ticks)
    $book = $excel.Workbooks.Open($bookPath,0,$false)
    if ($TryRepresentableBracketRecovery) {
        $search = $book.VBProject.VBComponents.Item('CLoadMultiplierSearch').CodeModule
        $body = $search.Lines(1,$search.CountOfLines)
        $pattern = '(?im)^            If lowLambda > 0# And reducedLambda - lowLambda <= request\.Problem\.LambdaTolerance Then\r?\n                request\.Problem\.HandleBracketNumericalFailure lowLambda, highLambda, probeState\r?\n                Exit Function\r?\n            End If\r?\n'
        if ([regex]::Matches($body,$pattern).Count -ne 1) { throw 'Temporary recovery gate not found exactly once.' }
        $changed = [regex]::Replace($body,$pattern,'')
        $search.DeleteLines(1,$search.CountOfLines)
        $search.AddFromString($changed)
        Write-Output 'TEMPORARY_EXPERIMENT: no absolute-lambda stopping before a physical bracket; isolated copy only.'
    }
    $module = $book.VBProject.VBComponents.Item($Component).CodeModule
    $module.AddFromString([IO.File]::ReadAllText((Join-Path $PSScriptRoot $DiagnosticFile),[Text.Encoding]::UTF8))
    $report = [string]$excel.Run("'" + $book.Name + "'!" + $Component + '.' + $Macro, $ShapeName)
    $report | Set-Content -LiteralPath (Join-Path $directory 'Model.txt') -Encoding UTF8
    if ($report -match 'RUNTIME ERROR') { throw $report }
    if((Get-FileHash -LiteralPath $source).Hash -cne $sourceHash){throw 'Source changed during model export.'}
    [ordered]@{Captured=(Get-Date).ToString('o'); SourceWorkbook=$source; SourceSHA256=$sourceHash;
        DataSHA256=(Get-FileHash -LiteralPath (Join-Path $directory 'Model.txt')).Hash;
        HelperSHA256=(Get-FileHash -LiteralPath (Join-Path $PSScriptRoot $DiagnosticFile)).Hash;
        IsolatedCopyOnly=$true; Component=$Component; Macro=$Macro; Shape=$ShapeName} |
        ConvertTo-Json -Depth 3 | Set-Content -LiteralPath (Join-Path $directory 'ModelManifest.json') -Encoding UTF8
    Write-Output ('BOUNDS_MODEL_READ: lines=' + ($report -split '\r?\n').Count + '; sourceUnchanged=' + ((Get-FileHash -LiteralPath $source).Hash -ceq $sourceHash))
    $witnessPath = Join-Path $directory 'EquilibriumWitnesses.json'
    if (Test-Path -LiteralPath $witnessPath) {
        $verified = @()
        foreach ($witness in (Get-Content -LiteralPath $witnessPath -Raw | ConvertFrom-Json)) {
            $mode=if($witness.Mode){[string]$witness.Mode}else{'Yes'}
            $profile=if($witness.Profile){[string]$witness.Profile}else{'PR1'}
            $result = [string]$excel.Run("'" + $book.Name + "'!modTestBatchCalculation.VerifyRemainingWitness", [double]$witness.CenterStrain, [double]$witness.KappaX, [double]$witness.KappaY, [double]$witness.N, [double]$witness.MxCenter, [double]$witness.MyCenter, $mode, $profile)
            if ($result -match 'RUNTIME ERROR' -or $result -notmatch 'confirmed=True') { throw ('Witness rejected: ' + $result) }
            if($witness.PSObject.Properties.Name -contains 'MinConcreteStrain'){
                if($result -notmatch '\|minConcreteStrain=([^|]+)'){throw 'Missing strain witness.'}
                $actual=[double]::Parse($Matches[1],[Globalization.CultureInfo]::InvariantCulture)
                if([math]::Abs($actual-[double]$witness.MinConcreteStrain) -gt 1e-10){throw 'Independent strain witness differs.'}
            }
            $verified += ($witness.Load + '|' + $result)
            Write-Output $verified[-1]
        }
        $verified | Set-Content -LiteralPath (Join-Path $directory 'VbaWitnesses.txt') -Encoding UTF8
    }
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit() }
}
