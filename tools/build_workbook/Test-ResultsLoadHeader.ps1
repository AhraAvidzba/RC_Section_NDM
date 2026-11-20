# ==========================================================================
# ДЛЯ ТЕСТОВ: две строки шапки после штатного расчёта Excel/File и загрузки снимка.
# ==========================================================================
param([Parameter(Mandatory=$true)][string]$WorkbookPath,[Parameter(Mandatory=$true)][string]$EvidenceDirectory)
$ErrorActionPreference='Stop';$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$source=(Resolve-Path -LiteralPath (Join-Path $root $WorkbookPath)).Path;$hash=(Get-FileHash -LiteralPath $source).Hash
$directory=Join-Path $root $EvidenceDirectory;[void][IO.Directory]::CreateDirectory($directory)
$path=Join-Path $directory 'RuntimeHeader.xlsm';Copy-Item -LiteralPath $source -Destination $path
$app=$null;$book=$null;$owned=0;$ticks=0;$stages=@()

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Записывает только используемую в частной проверке настройку по её ключу.
function Test-HeaderSetting([string]$Key,[string]$Value){
 $range=$book.Names.Item('rngSystemSettings').RefersToRange
 for($row=2;$row -le $range.Rows.Count;$row++){if([string]$range.Cells.Item($row,1).Value2 -ceq $Key){$range.Cells.Item($row,2).Value2=$Value;return}}
 throw ('Setting not found: '+$Key)
}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Проверяет фактический результат писателя после рабочего вызова.
function Test-HeaderValue([string]$Stage){
 $value=[string]$book.Worksheets.Item('Results').Range('AL8').Value2
 if($value -cne "Нагрузки с учётом продольного`nизгиба (если применимо)"){throw ('Wrong header after '+$Stage+': '+$value)}
 foreach($address in @('Q44:V58','Y44:AA58')){if([string]$book.Worksheets.Item('Расчет').Range($address).NumberFormat -cne '0.0;-0.0;0.0;@'){throw ('Calculation reformatted inputs: '+$address)}}
 return [ordered]@{Stage=$Stage;Header=$value;InputNumberFormatPreserved=$true}
}
Add-Type 'using System;using System.Runtime.InteropServices;public static class LoadHeaderTestWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
try{
 $memory=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory;if([double]$memory.CommitLimit-[double]$memory.CommittedBytes -lt 4GB){throw 'Insufficient memory'}
 $app=New-Object -ComObject Excel.Application;$app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=1
 [uint32]$id=0;[void][LoadHeaderTestWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id);$owned=[int]$id;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
 Write-Output ('OWNED_EXCEL: '+$owned+'; ticks='+$ticks)
 $book=$app.Workbooks.Open($path,0,$false);if($book.ReadOnly){throw 'Private copy must be writable'}
 $prefix="'"+$book.Name+"'!"
 [void]$app.Run($prefix+'modTestPerformance.TestConfigurePerformanceFixture','DIRECT',1,'Newton','Circle','Normal','No')
 foreach($item in @(@('General.NonCriticalMessagesEnabled','No'),@('General.ExecutionReportEnabled','No'),@('Plot.Enabled','No'),@('General.ReformatResultsEnabled','No'),@('Results.SnapshotName','HeaderCheck'))){Test-HeaderSetting $item[0] $item[1]}
 foreach($storage in @('Excel','File')){
  Test-HeaderSetting 'Results.SnapshotStorage' $storage
  $formation=[string]$app.Run($prefix+'modWorkbookCalculation.FormSectionForWorkbook',$book) # File сначала сохраняет свою геометрию как GeometryPreview; лист Excel не подменяет отсутствующий файл.
  if($formation -notmatch 'Сечение сформировано'){throw $formation}
  $message=[string]$app.Run($prefix+'modWorkbookCalculation.RunSectionCalculationForWorkbook',$book)
  if($message -notmatch 'Расчет завершен'){throw $message}
  $stages+=Test-HeaderValue ('Calculation '+$storage)
 }
 $book.Worksheets.Item('Results').Range('AL8').Value2='Broken header in private test'
 [void]$app.Run($prefix+'modCalculationSnapshot.UpdateSectionResultsForWorkbook',$book)
 $stages+=Test-HeaderValue 'File snapshot restore'
 [ordered]@{Status='PASS';SourceSHA256=$hash;Stages=$stages;PID=$owned;StartTicksUTC=$ticks} | ConvertTo-Json -Depth 7 | Set-Content -Encoding UTF8 -LiteralPath (Join-Path $directory 'RuntimeHeaderReview.json')
 Write-Output 'RUNTIME_HEADER_PASS'
}finally{
 if($book){$book.Close($false)};if($app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
 if($owned){$p=Get-Process -Id $owned -ErrorAction SilentlyContinue;if($p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.ToUniversalTime().Ticks -eq $ticks){Stop-Process -Id $owned -Force}}
}
if((Get-FileHash -LiteralPath $source).Hash -cne $hash){throw 'Source changed during runtime check'}
