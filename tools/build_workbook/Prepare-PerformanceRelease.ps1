# Calculates only a private candidate with unchanged user inputs. Save/reopen,
# passive plot/export and repeated macros are verified before publication.
param([string]$SourceWorkbook='docs/regression/Performance/FinalCandidate/RC_Section_NDM.xlsm',
    [string]$BaselineWorkbook='docs/regression/Performance/Baseline/RC_Section_NDM.xlsm',
    [string]$Directory='docs/regression/Performance/Release')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
$allowed=[IO.Path]::GetFullPath((Join-Path $root 'docs/regression/Performance'))+[IO.Path]::DirectorySeparatorChar
$output=[IO.Path]::GetFullPath((Join-Path $root $Directory))
if(-not $output.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)) {throw 'Release evidence outside Performance.'}
$source=(Resolve-Path -LiteralPath (Join-Path $root $SourceWorkbook)).Path
$baseline=(Resolve-Path -LiteralPath (Join-Path $root $BaselineWorkbook)).Path
$sourceHash=(Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
$baselineHash=(Get-FileHash -LiteralPath $baseline -Algorithm SHA256).Hash
New-Item -ItemType Directory -Path $output -Force | Out-Null
$target=Join-Path $output 'RC_Section_NDM.xlsm'
if(Test-Path -LiteralPath $target) {throw 'Use a fresh release directory.'}
Copy-Item -LiteralPath $source -Destination $target
$printAreas=@(Get-WorkbookPrintAreas $target)
Add-Type -TypeDefinition @'
using System; using System.Runtime.InteropServices;
public static class PerformanceReleaseIdentity {
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr w, out uint p);
}
'@
function Hash-Text([string]$value) {
    $sha=[Security.Cryptography.SHA256]::Create()
    try {return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($value)))).Replace('-','')}
    finally {$sha.Dispose()}
}
function Input-Fingerprint($book) {
    $sheet=$book.Worksheets.Item('Config');$range=$sheet.UsedRange
    $values=$range.Formula
    $names=@();foreach($name in $book.Names) {$names += ([string]$name.Name+'|'+[string]$name.RefersTo)}
    $widths=@();$results=$book.Worksheets.Item('Results')
    for($column=1;$column -le 260;$column++) {$widths += [double]$results.Columns.Item($column).ColumnWidth}
    $data=[ordered]@{Address=$range.Address();Rows=$range.Rows.Count;Columns=$range.Columns.Count;Formulas=$values;Names=@($names | Sort-Object);ResultWidths=$widths}
    return Hash-Text ($data | ConvertTo-Json -Depth 6 -Compress)
}
function Result-Fingerprint($excel,$book) {return [string]$excel.Run("'$($book.Name)'!modTestPerformance.PerformanceResultsFingerprint")}
$excel=$null;$book=$null;$records=@()
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3
    [uint32]$ownedPID=0;[void][PerformanceReleaseIdentity]::GetWindowThreadProcessId([IntPtr][long]$excel.Hwnd,[ref]$ownedPID)
    $process=Get-Process -Id $ownedPID
    [ordered]@{ExcelPID=$ownedPID;StartTicks=$process.StartTime.ToUniversalTime().Ticks} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $output 'Environment.json') -Encoding UTF8
    $book=$excel.Workbooks.Open($baseline,0,$true)
    $inputHash=Input-Fingerprint $book
    $book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book);$book=$null
    $excel.AutomationSecurity=1
    $book=$excel.Workbooks.Open($target)
    if((Input-Fingerprint $book) -cne $inputHash) {throw 'Candidate changed Config, names or Results widths.'}
    $reference=''
    for($run=1;$run -le 10;$run++) {
        $before=(Get-Process -Id $ownedPID).PrivateMemorySize64
        $result=[string]$excel.Run("'$($book.Name)'!modTestPerformance.MeasurePerformanceBatch",'Full')
        $parts=@{};foreach($part in $result -split ([string][char]30)) {$pair=$part -split '=',2;if($pair.Count -eq 2){$parts[$pair[0]]=$pair[1]}}
        $fingerprint=$parts['fingerprint']
        if($run -eq 1) {$reference=$fingerprint}
        elseif($fingerprint -cne $reference) {throw 'Repeated user macro changed engineering results.'}
        if((Input-Fingerprint $book) -cne $inputHash) {throw 'Calculation changed inputs, named anchors or widths.'}
        $records += [ordered]@{Run=$run;Seconds=[double]::Parse($parts['seconds'],[Globalization.CultureInfo]::InvariantCulture);PrivateBytesBefore=$before;PrivateBytesAfter=(Get-Process -Id $ownedPID).PrivateMemorySize64;ExactResults=$true;InputsPreserved=$true}
        $records | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $output 'RepeatedMacros.json') -Encoding UTF8
    }
    $savedFingerprint=Result-Fingerprint $excel $book
    [IO.File]::WriteAllText((Join-Path $output 'ResultsFingerprint.txt'),$savedFingerprint,[Text.UTF8Encoding]::new($false))
    $book.Save();$book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book);$book=$null
    Restore-WorkbookPrintAreas $target $printAreas
    $book=$excel.Workbooks.Open($target,0,$true)
    if((Result-Fingerprint $excel $book) -cne $savedFingerprint) {throw 'Save/reopen changed Results.'}
    if((Input-Fingerprint $book) -cne $inputHash) {throw 'Save/reopen changed inputs or widths.'}
    $exportBefore=[string]$excel.Run("'$($book.Name)'!modAutoCADStressExport.Audit02ReadExportSnapshotForTests",$book)
    [void]$excel.Run("'$($book.Name)'!modSolverWorkStats.ResetSectionEquilibriumSolveCount")
    [void]$excel.Run("'$($book.Name)'!modWorkbookCalculation.UpdateSectionPlotForWorkbook",$book)
    $exportAfter=[string]$excel.Run("'$($book.Name)'!modAutoCADStressExport.Audit02ReadExportSnapshotForTests",$book)
    $solves=[long]$excel.Run("'$($book.Name)'!modSolverWorkStats.SectionEquilibriumSolveCount")
    if($solves -ne 0 -or $exportAfter -cne $exportBefore -or (Result-Fingerprint $excel $book) -cne $savedFingerprint) {throw 'Passive plot/export changed Results or solved again.'}
    $book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book);$book=$null
    if((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -cne $sourceHash -or (Get-FileHash -LiteralPath $baseline -Algorithm SHA256).Hash -cne $baselineHash) {throw 'Source changed.'}
    [ordered]@{SourceSHA256=$sourceHash;BaselineSHA256=$baselineHash;ReleaseSHA256=(Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash;InputFingerprint=$inputHash;ResultsFingerprint=(Hash-Text $savedFingerprint);SaveReopenExact=$true;PassiveSolveCount=$solves;PassiveExportExact=$true;RepeatedMacros=10;InputsNamesWidthsPreserved=$true} | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $output 'Acceptance.json') -Encoding UTF8
    Write-Output 'RELEASE: unchanged user inputs; repeat 10/10 exact; save/reopen exact; passive plot/export solves=0.'
}
finally {
    if($book) {try {$book.Close($false)} catch {Write-Warning $_};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}
    if($excel) {try {$excel.Quit()} catch {Write-Warning $_};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
