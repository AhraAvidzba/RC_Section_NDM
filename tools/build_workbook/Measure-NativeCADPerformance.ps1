# Measures real import/export on private books and an owned Autodesk AutoCAD.
# Exact geometry signatures are compared before any speedup is accepted.
param(
    [string]$BaselineWorkbook='docs/regression/Performance/CADBaselineV3/RC_Section_NDM.xlsm',
    [string]$CandidateWorkbook='docs/regression/Performance/CADCandidateV5/RC_Section_NDM.xlsm',
    [string]$Directory='docs/regression/Performance/NativeCADV4',
    [ValidateRange(1,10)][int]$Repetitions=5
)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$output=[IO.Path]::GetFullPath((Join-Path $root $Directory))
$allowed=(Join-Path $root 'docs/regression/Performance')+[IO.Path]::DirectorySeparatorChar
if(-not $output.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)) {throw 'Evidence must stay in Performance.'}
if(Test-Path -LiteralPath $output) {throw 'Use a new evidence directory.'}
New-Item -ItemType Directory -Path $output | Out-Null
$sources=@{A=(Resolve-Path -LiteralPath (Join-Path $root $BaselineWorkbook)).Path;B=(Resolve-Path -LiteralPath (Join-Path $root $CandidateWorkbook)).Path}
$hashes=@{};foreach($key in @('A','B')) {$hashes[$key]=(Get-FileHash -LiteralPath $sources[$key]).Hash}
$class=(Get-ItemProperty -LiteralPath 'Registry::HKEY_CLASSES_ROOT\AutoCAD.Application.24.2\CLSID').'(default)'
$server=(Get-ItemProperty -LiteralPath "Registry::HKEY_CLASSES_ROOT\CLSID\$class\LocalServer32").'(default)'
if($server -notlike 'C:\Program Files\Autodesk\AutoCAD 2023\acad.exe*') {throw "Unexpected CAD server: $server"}
Add-Type @'
using System; using System.Runtime.InteropServices;
public static class NativeCADPerformanceIdentity {
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr w,out uint p);
}
'@
function Required-Property($target,[string]$name) {return $target.GetType().InvokeMember($name,[Reflection.BindingFlags]::GetProperty,$null,$target,$null)}
function Signature-Hash([string]$text) {
    $sha=[Security.Cryptography.SHA256]::Create()
    try {return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($text)))).Replace('-','')}
    finally {$sha.Dispose()}
}
$existing=@(Get-Process acad -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$existingExcel=@(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$excel=$null;$book=$null;$acad=$null;$doc=$null;$owned=$false;$ownedExcel=$false;$records=@();$reference=@{}
try {
    $acad=New-Object -ComObject AutoCAD.Application.24.2
    $fullName=[string](Required-Property $acad 'FullName')
    if($fullName -cne 'C:\Program Files\Autodesk\AutoCAD 2023\acad.exe') {throw "Not Autodesk AutoCAD: $fullName"}
    [uint32]$ownedID=0
    [void][NativeCADPerformanceIdentity]::GetWindowThreadProcessId([IntPtr][long](Required-Property $acad 'HWND'),[ref]$ownedID)
    $owned=($ownedID -gt 0 -and $existing -notcontains [int]$ownedID)
    if(-not $owned) {throw 'CAD instance was not created by this measurement.'}
    $acad.Visible=$false
    [ordered]@{Executable=$fullName;PID=$ownedID;StartTicks=(Get-Process -Id $ownedID).StartTime.ToUniversalTime().Ticks;UserDocumentsUsed=$false;Sources=$sources;SourceHashes=$hashes} |
        ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $output 'Environment.json') -Encoding UTF8
    Write-Output "OWNED_CAD: PID=$ownedID; $fullName"
    $excel=New-Object -ComObject Excel.Application
    [uint32]$excelID=0
    [void][NativeCADPerformanceIdentity]::GetWindowThreadProcessId([IntPtr][long](Required-Property $excel 'HWND'),[ref]$excelID)
    $ownedExcel=($excelID -gt 0 -and $existingExcel -notcontains [int]$excelID)
    if(-not $ownedExcel) {throw 'Excel instance was not created by this measurement.'}
    [ordered]@{PID=$excelID;StartTicks=(Get-Process -Id $excelID).StartTime.ToUniversalTime().Ticks} |
        ConvertTo-Json | Set-Content -LiteralPath (Join-Path $output 'ExcelIdentity.json') -Encoding UTF8
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=1
    foreach($scenario in @('Grid','Rotated','GridWithOtherEntities')) {
        for($run=0;$run -le $Repetitions;$run++) {
            $order=@('A','B');if($run%2 -eq 1) {$order=@('B','A')}
            foreach($version in $order) {
                if(Test-Path -LiteralPath (Join-Path $output 'Stop.txt')) {throw 'Measurement cancelled at a safe operation boundary.'}
                Write-Output "RUN: scenario=$scenario; run=$run; version=$version"
                if((Get-FileHash -LiteralPath $sources[$version]).Hash -cne $hashes[$version]) {throw 'Measurement source changed.'}
                $book=$excel.Workbooks.Open($sources[$version],0,$true)
                [void]$excel.Run("'$($book.Name)'!modTestPerformance.PrepareNativeCADPerformance",($scenario -eq 'Rotated'))
                $doc=$acad.Documents.Add()
                $noise=0;if($scenario -eq 'GridWithOtherEntities') {$noise=2000}
                $result=[string]$excel.Run("'$($book.Name)'!modTestPerformance.MeasureNativeCADPerformance",$doc,$noise)
                $parts=@{};foreach($part in $result -split ([string][char]30)) {$pair=$part -split '=',2;if($pair.Count -eq 2) {$parts[$pair[0]]=$pair[1]}}
                $signature=$parts.signature
                $name="$scenario-$run-$version"
                $signature | Set-Content -LiteralPath (Join-Path $output "$name.geometry.txt") -Encoding UTF8
                $digest=Signature-Hash $signature
                $record=[ordered]@{Scenario=$scenario;Run=$run;Warmup=($run -eq 0);Version=$version;
                    ExportSeconds=[double]::Parse($parts.export,[Globalization.CultureInfo]::InvariantCulture);
                    ImportSeconds=[double]::Parse($parts.import,[Globalization.CultureInfo]::InvariantCulture);
                    Entities=[int]$parts.entities;Concrete=[int]$parts.concrete;Rebar=[int]$parts.rebar;UnrelatedEntities=$noise;SignatureSHA256=$digest}
                $records += $record
                $records | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $output 'Measurements.json') -Encoding UTF8
                Write-Output ($record | ConvertTo-Json -Compress)
                $doc.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc);$doc=$null
                $book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book);$book=$null
                $key="$scenario-$run"
                if($reference.ContainsKey($key) -and $reference[$key] -cne $digest) {throw "Exact A/B geometry mismatch: $key. No speedup accepted."}
                $reference[$key]=$digest
            }
        }
    }
    [ordered]@{Passed=$true;ExactABGeometry=$true;Repetitions=$Repetitions;CompletedUTC=[DateTime]::UtcNow.ToString('o')} |
        ConvertTo-Json | Set-Content -LiteralPath (Join-Path $output 'Acceptance.json') -Encoding UTF8
}
finally {
    if($doc) {try {$doc.Close($false)} catch {}}
    if($book) {try {$book.Close($false)} catch {}}
    if($excel) {if($ownedExcel) {try {$excel.Quit()} catch {}};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    if($acad) {if($owned) {try {$acad.Quit()} catch {}};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($acad)}
    foreach($key in @('A','B')) {if((Get-FileHash -LiteralPath $sources[$key]).Hash -cne $hashes[$key]) {throw "Source changed: $key"}}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
