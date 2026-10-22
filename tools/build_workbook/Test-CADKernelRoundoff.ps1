# Сравнивает конечные расчёты после двух способов настоящего CAD-экспорта/импорта.
# Исходные книги открываются только для чтения; контрольные настройки не сохраняются.
param([Parameter(Mandatory=$true)][string]$BaselineWorkbook,
      [Parameter(Mandatory=$true)][string]$CandidateWorkbook,
      [Parameter(Mandatory=$true)][string]$EvidenceDirectory)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$directory=[IO.Path]::GetFullPath((Join-Path $root $EvidenceDirectory))
$allowed=(Join-Path $root 'docs/regression')+[IO.Path]::DirectorySeparatorChar
if(-not $directory.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)){throw 'Evidence must stay in docs/regression.'}
if(Test-Path -LiteralPath $directory){throw 'Use a new evidence directory.'}
[void][IO.Directory]::CreateDirectory($directory)
$sources=@{A=(Resolve-Path -LiteralPath $BaselineWorkbook).Path;B=(Resolve-Path -LiteralPath $CandidateWorkbook).Path}
$hashes=@{};foreach($key in @('A','B')){$hashes[$key]=(Get-FileHash -LiteralPath $sources[$key]).Hash}
$existing=@(Get-Process acad -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class CADMathWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
function Native-Property($target,[string]$name){
 for($attempt=0;$attempt -lt 80;$attempt++){
  try{return $target.GetType().InvokeMember($name,[Reflection.BindingFlags]::GetProperty,$null,$target,$null)}
  catch{if($attempt -eq 79){throw};Start-Sleep -Milliseconds 250}
 }
}
$excel=$null;$acad=$null;$books=@{};$docs=@{};$excelID=0;$cadID=0;$excelTicks=0;$cadTicks=0
try{
 $excel=New-Object -ComObject Excel.Application
 $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=1
 [uint32]$id=0;[void][CADMathWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id);$excelID=[int]$id;$excelTicks=(Get-Process -Id $excelID).StartTime.ToUniversalTime().Ticks
 foreach($key in @('A','B')){$books[$key]=$excel.Workbooks.Open($sources[$key],0,$true);if(-not $books[$key].ReadOnly){throw 'Read-only source required.'}}
 $acad=New-Object -ComObject AutoCAD.Application.24.2
 if((Native-Property $acad 'FullName') -cne 'C:\Program Files\Autodesk\AutoCAD 2023\acad.exe'){throw 'Unexpected AutoCAD executable.'}
 [uint32]$id=0;[void][CADMathWindow]::GetWindowThreadProcessId([IntPtr][long](Native-Property $acad 'HWND'),[ref]$id)
 if($existing -contains [int]$id){throw 'Refusing to use a user AutoCAD instance.'}
 $cadID=[int]$id;$cadTicks=(Get-Process -Id $cadID).StartTime.ToUniversalTime().Ticks;$acad.Visible=$false
 [ordered]@{Sources=$sources;SourceHashes=$hashes;ExcelPID=$excelID;ExcelStartTicks=$excelTicks;CADPID=$cadID;CADStartTicks=$cadTicks;UserDocumentsUsed=$false} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'Identity.json') -Encoding UTF8
 Write-Output ('OWNED_CAD_MATH: Excel='+$excelID+'; CAD='+$cadID)
 $measurements=@()
 foreach($version in @('A','B')){
  $prefix="'"+$books[$version].Name+"'!"
  Write-Output ('PREPARE_RECTANGLE: '+$version)
  [void]$excel.Run($prefix+'modTestPerformance.TestPrepareNativeCADRoundoffFixture')
  $docs[$version]=$acad.Documents.Add();$docs[$version].Activate()
  $result=[string]$excel.Run($prefix+'modTestPerformance.TestMeasureNativeCADPerformance',$docs[$version],0)
  $parts=@{};foreach($part in $result -split ([string][char]30)){$pair=$part -split '=',2;if($pair.Count -eq 2){$parts[$pair[0]]=$pair[1]}}
  $parts.signature | Set-Content -LiteralPath (Join-Path $directory ($version+'.geometry.txt')) -Encoding UTF8
  $measurements+=@{Version=$version;ExportSeconds=$parts.export;ImportSeconds=$parts.import;Concrete=$parts.concrete;Rebar=$parts.rebar;Entities=$parts.entities}
  $measurements | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $directory 'Measurements.json') -Encoding UTF8
 }
 $prefix="'"+$books.B.Name+"'!"
 foreach($mode in @('Yes','No')){
  foreach($version in @('A','B')){
   Write-Output ('CALCULATE_IMPORTED: '+$version+'; extension='+$mode+'; cases=6')
   $watch=[Diagnostics.Stopwatch]::StartNew()
   $result=[string]$excel.Run($prefix+'modTestPerformance.TestMeasureCADRoundoffCalculation',$docs[$version],$mode)
   $watch.Stop()
   $result | Set-Content -LiteralPath (Join-Path $directory ($version+'.'+$mode+'.calc.txt')) -Encoding UTF8
   Write-Output ('CALCULATION_COMPLETE: '+$version+'; extension='+$mode+'; seconds='+$watch.Elapsed.TotalSeconds)
  }
 }
}finally{
 foreach($doc in $docs.Values){try{$doc.Close($false)}catch{}}
 foreach($book in $books.Values){try{$book.Close($false)}catch{}}
 if($excel){try{$excel.Quit()}catch{};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
 if($acad){if($cadID){try{$acad.Quit()}catch{}};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($acad)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
 foreach($entry in @(@($excelID,$excelTicks,'EXCEL'),@($cadID,$cadTicks,'acad'))){if($entry[0]){$p=Get-Process -Id $entry[0] -ErrorAction SilentlyContinue;if($p -and $p.ProcessName -eq $entry[2] -and $p.StartTime.ToUniversalTime().Ticks -eq $entry[1]){Stop-Process -Id $entry[0] -Force -ErrorAction SilentlyContinue}}}
 foreach($key in @('A','B')){if((Get-FileHash -LiteralPath $sources[$key]).Hash -cne $hashes[$key]){throw 'Source workbook changed.'}}
}
$python=Join-Path $env:USERPROFILE '.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe'
& $python (Join-Path $PSScriptRoot 'Compare-CADKernelEvidence.py') --directory $directory
if($LASTEXITCODE -ne 0){throw 'Kernel roundoff comparison failed; do not accept the export change.'}
Write-Output 'CAD_FINAL_RESULTS_EQUIVALENT: both modes, six combinations per source.'
