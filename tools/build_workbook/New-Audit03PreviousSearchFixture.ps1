# Restores one unchanged pre-fix search module in an isolated test workbook.
# All other current dependencies and tests remain intact for a narrow reproducer.
param(
    [string]$SourceWorkbook = 'docs/regression/Audit03/RC_Section_NDM_f06_fixed.xlsm',
    [string]$OutputWorkbook = 'docs/regression/Audit03/RC_Section_NDM_ultimate_negative.xlsm',
    [string]$Revision = '1a65796'
)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
if ($Revision -notmatch '^[0-9a-fA-F]{7,40}$') { throw 'Expected a fixed Git revision hash.' }
$source = (Resolve-Path -LiteralPath (Join-Path $root $SourceWorkbook)).Path
$output = [IO.Path]::GetFullPath((Join-Path $root $OutputWorkbook))
$allowed = [IO.Path]::GetFullPath((Join-Path $root 'docs/regression/Audit03')) + [IO.Path]::DirectorySeparatorChar
if (-not $output.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase)) { throw 'Fixture must remain in Audit03 regression directory.' }
if ($source -eq $output) { throw 'Fixture must not overwrite the source workbook.' }
$hash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$lines = @(& git -C $root show "${Revision}:src/Solver/CUltimateStrainSearch.cls")
if ($LASTEXITCODE -ne 0) { throw 'Could not read the fixed pre-fix source revision.' }
$body = [System.Collections.Generic.List[string]]::new()
$inHeader = $false
foreach ($line in $lines) {
    if ($line -match '^VERSION\s+') { continue }
    if ($line -match '^BEGIN\s*$') { $inHeader = $true; continue }
    if ($inHeader) {
        if ($line -match '^END\s*$') { $inHeader = $false }
        continue
    }
    if ($line -match '^Attribute\s+VB_') { continue }
    $body.Add($line)
}
if ($body.Count -lt 100) { throw 'Unexpected pre-fix module body.' }
Copy-Item -LiteralPath $source -Destination $output -Force
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($output)
    $module = $book.VBProject.VBComponents.Item('CUltimateStrainSearch').CodeModule
    if ($module.CountOfLines -gt 0) { $module.DeleteLines(1, $module.CountOfLines) }
    $module.AddFromString(($body -join "`r`n"))
    $book.Save()
    $book.Close($false)
    $book = $null
    Write-Output "NEGATIVE_FIXTURE: revision=$Revision; replaced=CUltimateStrainSearch; tests=current; output=$OutputWorkbook"
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) {
        $excel.Quit()
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null
    }
    if ((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -ne $hash) { throw 'Source workbook changed.' }
}
