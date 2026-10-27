# Lists actual generated/saved geometry and selected loads without solving.
param([string]$ProbeDirectory='docs/regression/FocusedAcceptance_2026-10-10_V2', [string]$DeadlineUTC='')
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSHOME 'Modules/Microsoft.PowerShell.Utility/Microsoft.PowerShell.Utility.psd1') -Global -ErrorAction Stop
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$directory=Join-Path $root $ProbeDirectory
$source=Join-Path $directory 'RC_Section_NDM.xlsm'
$preparation=Get-Content -LiteralPath (Join-Path $directory 'Preparation.json') -Raw | ConvertFrom-Json
if (-not $preparation.Compiled -or $preparation.ProbeWorkbookSHA256 -cne (Get-FileHash -LiteralPath $source).Hash -or
    $preparation.HelperSHA256 -cne (Get-FileHash -LiteralPath (Join-Path $PSScriptRoot 'FocusedAcceptanceProbe.bas')).Hash) {
    throw 'Probe source or helper changed after compilation.'
}
$jobs=@(
    @{Shape='Rectangles'; Group='Generated'; Count=8},
    @{Shape='Circle'; Group='Generated'; Count=8},
    @{Shape='PointedSection'; Group='Generated'; Count=8},
    @{Shape='HollowRectangle'; Group='Generated'; Count=8},
    @{Shape='SavedWorkbook'; Group='Saved'; Count=23},
    @{Shape='SavedWorkbook'; Group='Paths'; Count=6}
)
# Проверяет сохранённую пару PID/время старта перед Close/Quit/Stop-Process.
# Отсутствующий или повторно использованный PID не даёт права закрывать приложение.
function Test-OwnedProcessIdentity {
    param([int]$ProcessId, [long]$StartTicks, [string]$ProcessName)
    if ($ProcessId -le 0 -or $StartTicks -le 0) { return $false }
    $current=Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
    return ($null -ne $current -and $current.ProcessName -eq $ProcessName -and
        $current.StartTime.ToUniversalTime().Ticks -eq $StartTicks)
}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Тот же предел ожидания памяти, что и при подготовке Focused; общий deadline не продлевается.
function Wait-TestFocusedPrepareMemory {
    param([string]$EvidenceDirectory, [string]$DeadlineUTC='')
    $started=[DateTimeOffset]::UtcNow
    $limit=$started.AddSeconds(300)
    if ($DeadlineUTC) { $outer=[DateTimeOffset]::Parse($DeadlineUTC,[Globalization.CultureInfo]::InvariantCulture); if ($outer -lt $limit) { $limit=$outer } }
    $proofPath=Join-Path $EvidenceDirectory 'MemoryBeforeExcel.json'
    while ($true) {
        $proof=[ordered]@{CheckedUTC=[DateTimeOffset]::UtcNow.ToString('o');Status='MEMORY_QUERY_FAILED';Ready=$false;RequiredPhysicalGiB=4;RequiredCommitGiB=8;WaitLimitUTC=$limit.ToString('o');NewExcelStarted=$false}
        try {
            $memory=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory -ErrorAction Stop
            foreach ($name in @('AvailableMBytes','CommitLimit','CommittedBytes')) { if ($null -eq $memory.PSObject.Properties[$name] -or $null -eq $memory.$name) { throw ('Missing memory counter: '+$name) } }
            $physical=[double]$memory.AvailableMBytes*1MB
            $commit=[double]$memory.CommitLimit-[double]$memory.CommittedBytes
            if ([double]::IsNaN($physical) -or [double]::IsInfinity($physical) -or [double]::IsNaN($commit) -or [double]::IsInfinity($commit) -or $physical -lt 0 -or [double]$memory.CommitLimit -le 0 -or [double]$memory.CommittedBytes -lt 0) { throw 'Invalid memory counters.' }
            $proof['FreePhysicalGiB']=$physical/1GB; $proof['FreeCommitGiB']=$commit/1GB
            $proof['CommitLimitBytes']=[double]$memory.CommitLimit; $proof['CommittedBytes']=[double]$memory.CommittedBytes
        } catch {
            $proof['Error']=$_.Exception.Message
            $proof | ConvertTo-Json | Set-Content -LiteralPath $proofPath -Encoding UTF8
            throw 'Memory counters unavailable; no preparation Excel started.'
        }
        $remaining=($limit-[DateTimeOffset]::UtcNow).TotalSeconds
        $ready=$physical -ge 4GB -and $commit -ge 8GB -and $remaining -gt 0
        $proof.Ready=$ready; $proof.Status=if ($remaining -le 0) {'MEMORY_WAIT_DEADLINE'} elseif ($ready) {'READY'} else {'WAITING_FOR_MEMORY'}
        $proof['WaitedSeconds']=([DateTimeOffset]::UtcNow-$started).TotalSeconds
        $proof | ConvertTo-Json | Set-Content -LiteralPath $proofPath -Encoding UTF8
        if ($ready) { return }
        if ($remaining -le 0) { throw 'Memory reserve/deadline insufficient (4 GiB physical, 8 GiB commit); no preparation Excel started.' }
        Start-Sleep -Milliseconds ([int][math]::Max(1,[math]::Min(10000,[math]::Ceiling($remaining*1000))))
    }
}

$records=@(); $excel=$null; $book=$null; $ownedPid=0; $ownedTicks=0
try {
    Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class FocusedDescribeWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId); }'
    $memoryDirectory=Join-Path $directory 'DescribeMemory'
    [void][IO.Directory]::CreateDirectory($memoryDirectory)
    Wait-TestFocusedPrepareMemory -EvidenceDirectory $memoryDirectory -DeadlineUTC $DeadlineUTC
    $existingExcelIds=@([Diagnostics.Process]::GetProcessesByName('EXCEL') | ForEach-Object { $_.Id })
    $excel=New-Object -ComObject Excel.Application
    [uint32]$id=0
    [void][FocusedDescribeWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    if ($id -eq 0 -or $existingExcelIds -contains [int]$id) {
        throw 'COM returned an existing or unidentified Excel; no properties changed and no Quit allowed.'
    }
    $newExcelProcess=Get-Process -Id $id -ErrorAction Stop
    $newExcelTicks=$newExcelProcess.StartTime.ToUniversalTime().Ticks
    if (-not (Test-OwnedProcessIdentity ([int]$id) $newExcelTicks 'EXCEL')) { throw 'New Excel identity could not be verified.' }
    $ownedPid=[int]$id; $ownedTicks=$newExcelTicks
    $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false; $excel.AutomationSecurity=1
    Write-Output ('OWNED_EXCEL: id='+$ownedPid+'; utcStartTicks='+$ownedTicks)
    $book=$excel.Workbooks.Open($source,0,$true)
    foreach ($job in $jobs) {
        $body=[string]$excel.Run("'RC_Section_NDM.xlsm'!modTestBatchCalculation.TestDescribeFocusedAcceptance",$job.Shape,$job.Group)
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
    try {
        if ($book -and (Test-OwnedProcessIdentity $ownedPid $ownedTicks 'EXCEL')) { $book.Close($false) }
    } finally {
        try {
            if ($excel) {
                try { if (Test-OwnedProcessIdentity $ownedPid $ownedTicks 'EXCEL') { $excel.Quit() } }
                finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
            }
        } finally {
            [GC]::Collect(); [GC]::WaitForPendingFinalizers()
            if (Test-OwnedProcessIdentity $ownedPid $ownedTicks 'EXCEL') {
                Stop-Process -Id $ownedPid -Force -ErrorAction SilentlyContinue
            }
        }
    }
}
