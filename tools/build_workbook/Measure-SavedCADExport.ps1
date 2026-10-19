# Проверяет сохранённый расчёт без замены входов: время, обратный импорт и WCS.
# Создаёт собственные Excel/AutoCAD, документы и фоновые объекты; исходники не сохраняет.
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
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class SavedCADWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
function Property($target,[string]$name){
 for($attempt=0;$attempt -lt 80;$attempt++){
  try{return $target.GetType().InvokeMember($name,[Reflection.BindingFlags]::GetProperty,$null,$target,$null)}
  catch{if($attempt -eq 79){throw};Start-Sleep -Milliseconds 250}
 }
}
function Hash-Text([string]$text){$sha=[Security.Cryptography.SHA256]::Create();try{[BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($text))).Replace('-','')}finally{$sha.Dispose()}}
function Region-Fingerprint($document){
 $text=New-Object Text.StringBuilder
 foreach($entity in $document.ModelSpace){
  if($entity.ObjectName -ne 'AcDbRegion'){continue}
  [void]$text.Append($entity.Layer+'|'+$entity.Color)
  # GetBoundingBox и Normal возвращают мировые координаты независимо от ПСК.
  # Центроид и инерции здесь не читаем: AutoCAD требует для них совпадения плоскостей.
  [object]$low=$null;[object]$high=$null;$entity.GetBoundingBox([ref]$low,[ref]$high)
  foreach($value in (@($low)+@($high)+@(Property $entity 'Normal'))){[void]$text.Append('|'+[BitConverter]::ToString([BitConverter]::GetBytes([double]$value)))}
  [void]$text.AppendLine()
 }
 $text.ToString()
}
$excel=$null;$acad=$null;$doc=$null;$background=$null;$books=@{};$excelID=0;$cadID=0;$excelTicks=0;$cadTicks=0;$records=@();$fingerprints=@{}
try{
 $excel=New-Object -ComObject Excel.Application
 $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=1
 [uint32]$id=0;[void][SavedCADWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id);$excelID=[int]$id;$excelTicks=(Get-Process -Id $excelID).StartTime.ToUniversalTime().Ticks
 foreach($key in @('A','B')){$books[$key]=$excel.Workbooks.Open($sources[$key],0,$true);if(-not $books[$key].ReadOnly){throw 'Read-only source required.'}}
 $acad=New-Object -ComObject AutoCAD.Application.24.2
 if((Property $acad 'FullName') -cne 'C:\Program Files\Autodesk\AutoCAD 2023\acad.exe'){throw 'Unexpected AutoCAD executable.'}
 [uint32]$id=0;[void][SavedCADWindow]::GetWindowThreadProcessId([IntPtr][long](Property $acad 'HWND'),[ref]$id)
 if($existing -contains [int]$id){throw 'Refusing to use a user AutoCAD instance.'}
 $cadID=[int]$id;$cadTicks=(Get-Process -Id $cadID).StartTime.ToUniversalTime().Ticks;$acad.Visible=$false
 [ordered]@{Sources=$sources;SourceHashes=$hashes;ExcelPID=$excelID;ExcelStartTicks=$excelTicks;CADPID=$cadID;CADStartTicks=$cadTicks;UserDocumentsUsed=$false} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'Identity.json') -Encoding UTF8
 Write-Output ('OWNED_SAVED_EXPORT: Excel='+$excelID+'; CAD='+$cadID)
 foreach($scenario in @('Roundtrip','Normal1','Normal2','BackgroundOpen','BackgroundClosed','ObliqueUCS')){
  if($scenario -eq 'BackgroundOpen'){
   $background=$acad.Documents.Add()
   [double[]]$point=@(0,100000,0)
   for($i=1;$i -le 3000;$i++){$point[0]=$i*10;[void]$background.ModelSpace.AddCircle($point,2.5)}
  }
  if($scenario -eq 'BackgroundClosed'){$background.Close($false);$background=$null}
  $order=@('A','B');if($scenario -in @('Normal1','BackgroundOpen','ObliqueUCS')){$order=@('B','A')}
  foreach($version in $order){
   $doc=$acad.Documents.Add();$doc.Activate()
   if($scenario -eq 'ObliqueUCS'){
    $ucs=$doc.UserCoordinateSystems.Add([double[]]@(1234,-3456,789),[double[]]@(1235,-3456,789),[double[]]@(1234,-3456,790),'NDM_TEST_OBLIQUE')
    $doc.ActiveUCS=$ucs
    $activeOrigin=@(Property (Property $doc 'ActiveUCS') 'Origin')
    if($activeOrigin.Count -ne 3 -or [double]$activeOrigin[2] -ne 789){throw 'Oblique UCS was not activated.'}
   }
   $prefix="'"+$books[$version].Name+"'!"
   if($scenario -eq 'Roundtrip'){
    $result=[string]$excel.Run($prefix+'modTestPerformance.MeasureNativeCADPerformance',$doc,0)
    $parts=@{};foreach($part in $result -split ([string][char]30)){$pair=$part -split '=',2;if($pair.Count -eq 2){$parts[$pair[0]]=$pair[1]}}
    $parts.signature | Set-Content -LiteralPath (Join-Path $directory ($version+'.geometry.txt')) -Encoding UTF8
    $fingerprints[$version]=Hash-Text $parts.signature
    $seconds=[double]::Parse($parts.export,[Globalization.CultureInfo]::InvariantCulture);$entities=[int]$parts.entities
   }else{
    $watch=[Diagnostics.Stopwatch]::StartNew()
    $entities=[int]$excel.Run($prefix+'modAutoCADStressExport.Audit03ExportAutoCADDocumentForTests',$doc)
    $watch.Stop();$seconds=$watch.Elapsed.TotalSeconds
   }
   if($scenario -eq 'ObliqueUCS'){
    $signature=Region-Fingerprint $doc
    $signature | Set-Content -LiteralPath (Join-Path $directory ($version+'.wcs.txt')) -Encoding UTF8
    $fingerprints[$version+'WCS']=Hash-Text $signature
    $firstRegion=$null;foreach($entity in $doc.ModelSpace){if($entity.ObjectName -eq 'AcDbRegion'){$firstRegion=$entity;break}}
    [object]$minimum=$null;[object]$maximum=$null;$firstRegion.GetBoundingBox([ref]$minimum,[ref]$maximum)
    # C1 сохранённого примера: центр (-487,5; -237,5), квадрат 25 мм, в плоскости WCS Z=0.
    $expected=@(-500,-250,0,-475,-225,0);$actual=@($minimum)+@($maximum)
    for($i=0;$i -lt 6;$i++){if([double]$actual[$i] -ne [double]$expected[$i]){throw ('WCS bounds differ: '+$version+'; coordinate '+$i)}}
   }
   $records+=,[ordered]@{Scenario=$scenario;Version=$version;Seconds=$seconds;Entities=$entities}
   $records | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'Measurements.json') -Encoding UTF8
   Write-Output ($records[-1] | ConvertTo-Json -Compress)
   $doc.Close($false);$doc=$null
  }
  if($scenario -eq 'Roundtrip' -and $fingerprints.A -cne $fingerprints.B){throw 'Saved geometry differs after roundtrip.'}
  if($scenario -eq 'ObliqueUCS' -and $fingerprints.AWCS -cne $fingerprints.BWCS){throw 'Region geometry differs in oblique UCS.'}
 }
 [ordered]@{Passed=$true;ExactSavedRoundtrip=$true;ExactRegionsInObliqueUCS=$true;WCSBoundsVerified=$true;Fingerprints=$fingerprints;SourcesUnchanged=$true} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'Acceptance.json') -Encoding UTF8
}finally{
 if($doc){try{$doc.Close($false)}catch{}}
 if($background){try{$background.Close($false)}catch{}}
 foreach($book in $books.Values){try{$book.Close($false)}catch{}}
 if($excel){try{$excel.Quit()}catch{};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
 if($acad){if($cadID){try{$acad.Quit()}catch{}};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($acad)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
 foreach($entry in @(@($excelID,$excelTicks,'EXCEL'),@($cadID,$cadTicks,'acad'))){if($entry[0]){$p=Get-Process -Id $entry[0] -ErrorAction SilentlyContinue;if($p -and $p.ProcessName -eq $entry[2] -and $p.StartTime.ToUniversalTime().Ticks -eq $entry[1]){Stop-Process -Id $entry[0] -Force -ErrorAction SilentlyContinue}}}
 foreach($key in @('A','B')){if((Get-FileHash -LiteralPath $sources[$key]).Hash -cne $hashes[$key]){throw 'Source workbook changed.'}}
}
Write-Output 'SAVED_EXPORT_PASS: roundtrip, background drawing and WCS geometry verified.'
