param(
    [Parameter(Mandatory=$true)][string]$ReferenceWorkbook,
    [string]$WorkbookPath = "workbook/output/RC_Section_NDM.xlsm"
)
$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "SettingsCatalog.ps1")
$printAreas=@(Get-WorkbookPrintAreas $ReferenceWorkbook)
$keys = @(
    "Capacity.InitialLambda", "Capacity.MaxLambda", "Capacity.ToleranceLambda",
    "Capacity.MaxRetries", "Capacity.BaseLoadSteps", "Capacity.SolverMaxIterations",
    "SLS.Crack.Allowable", "SLS.Crack.TensionZoneMode", "SLS.Crack.CoverDistanceMode",
    "SLS.Crack.Phi1", "SLS.Crack.Phi2", "SLS.Crack.Phi3Mode", "SLS.Crack.Phi3",
    "SLS.Crack.PsiMode", "SLS.Crack.PsiS"
)

# Restores only cells overwritten by the historical seven-column refresh.
# The reference is the independent fixture saved before that refresh ran.
function Find-SettingRow($book, [string]$key) {
    $range = $book.Names.Item("rngSystemSettings").RefersToRange
    $data = $range.Value2
    for($r=2; $r -le $range.Rows.Count; $r++) {
        if([string]$data[$r,1] -eq $key) { return ,($range.Cells.Item($r,2).Resize(1,6)) }
    }
    throw "Setting not found: $key"
}

$excel=$null
$source=$null
$target=$null
$referenceCopy=Join-Path ([IO.Path]::GetTempPath()) ("RC_NDM_ConfigBeforeRefresh_" + [Guid]::NewGuid().ToString("N") + ".xlsm")
Copy-Item -LiteralPath $ReferenceWorkbook -Destination $referenceCopy
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false
    $excel.DisplayAlerts=$false
    $excel.AutomationSecurity=3
    $source=$excel.Workbooks.Open($referenceCopy, $null, $true)
    $target=$excel.Workbooks.Open((Resolve-Path -LiteralPath $WorkbookPath).Path)
    if(-not $source -or -not $target) { throw "Could not open both workbooks." }
    foreach($key in $keys) {
        $from=Find-SettingRow $source $key
        $to=Find-SettingRow $target $key
        $from.Copy($to)
        # Keep workbook-local formulas; Copy alone can create external links.
        $to.Formula=$from.Formula
        Write-Output "RESTORED: $key; $($to.Address())"
    }
    $target.Save()
} catch {
    Write-Output $_.ScriptStackTrace
    throw
} finally {
    if($target) { $target.Close($false) }
    if($source) { $source.Close($false) }
    if($excel) { $excel.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
Restore-WorkbookPrintAreas (Resolve-Path -LiteralPath $WorkbookPath).Path $printAreas
