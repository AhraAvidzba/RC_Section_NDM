# Lists actual generated/saved geometry and selected loads without solving.
param([string]$ProbeDirectory='docs/regression/FocusedAcceptance_2026-10-10_V2')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$directory=Join-Path $root $ProbeDirectory
$source=Join-Path $directory 'RC_Section_NDM.xlsm'
$preparation=Get-Content -LiteralPath (Join-Path $directory 'Preparation.json') -Raw | ConvertFrom-Json
if (-not $preparation.Compiled -or $preparation.ProbeWorkbookSHA256 -cne (Get-FileHash -LiteralPath $source).Hash -or
    $preparation.HelperSHA256 -cne (Get-FileHash -LiteralPath (Join-Path $PSScriptRoot 'FocusedAcceptanceProbe.bas')).Hash) {
    throw 'Probe source or helper changed after compilation.'
}
$jobs=@(
    @{Shape='RectSet'; Group='Generated'; Count=8},
    @{Shape='Circle'; Group='Generated'; Count=8},
    @{Shape='RoundedRectangle'; Group='Generated'; Count=8},
    @{Shape='HollowRectangle'; Group='Generated'; Count=8},
    @{Shape='SavedWorkbook'; Group='Saved'; Count=23},
    @{Shape='SavedWorkbook'; Group='Paths'; Count=6}
)
$records=@(); $excel=$null; $book=$null; $ownedPid=0; $ownedTicks=0
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false; $excel.AutomationSecurity=1
    Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class FocusedDescribeWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId); }'
    [uint32]$id=0
    [void][FocusedDescribeWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $ownedPid=[int]$id; $ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    Write-Output ('OWNED_EXCEL: id='+$ownedPid+'; utcStartTicks='+$ownedTicks)
    $book=$excel.Workbooks.Open($source,0,$true)
    foreach ($job in $jobs) {
        $body=[string]$excel.Run("'RC_Section_NDM.xlsm'!modTestBatchCalculation.DescribeFocusedAcceptance",$job.Shape,$job.Group)
        if ($body -match 'RUNTIME ERROR|FAIL:' -or $body -notmatch 'FOCUSED_MANIFEST_COMPLETE: no calculation executed') {
            throw ('Invalid geometry/load manifest: '+$job.Shape+'/'+$job.Group+'; '+$body)
        }
        $inputs=@($body -split '\r?\n' | Where-Object { $_ -like 'FOCUSED_INPUT|*' })
        if ($inputs.Count -ne $job.Count) { throw 'Manifest load count differs from plan.' }
        $report=Join-Path $directory ($job.Shape+'_'+$job.Group+'_Manifest.txt')
        $body | Set-Content -LiteralPath $report -Encoding UTF8
        $records+=@{Shape=$job.Shape; Group=$job.Group; CasesPerMode=$job.Count;
            Report=$report; ReportSHA256=(Get-FileHash -LiteralPath $report).Hash}
        Write-Output (($body -split '\r?\n' | Where-Object { $_ -like 'FOCUSED_MANIFEST|*' }) -join '')
    }
    $book.Close($false); $book=$null
    if ((Get-FileHash -LiteralPath $source).Hash -cne $preparation.ProbeWorkbookSHA256 -or
        (Get-FileHash -LiteralPath $preparation.SourceWorkbook).Hash -cne $preparation.SourceSHA256) {
        throw 'Source changed during manifest preparation.'
    }
    [ordered]@{Passed=$true; Created=(Get-Date).ToString('o'); SourceSHA256=$preparation.SourceSHA256;
        ProbeSHA256=$preparation.ProbeWorkbookSHA256; Groups=$records; SingleCasesTwoModes=122;
        SolverExecuted=$false; AcceptanceExecuted=$false; ProductionChanged=$false} |
        ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $directory 'InputManifest.json') -Encoding UTF8
    Write-Output 'FOCUSED_MANIFEST_READY: geometry and loads only; no acceptance calculations executed.'
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
