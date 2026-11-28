# ДЛЯ ТЕСТОВ: проверяет структуру, оформление и связи новых полей в готовом файле без его сохранения.
param([Parameter(Mandatory=$true)][string]$WorkbookPath,[Parameter(Mandatory=$true)][string]$EvidenceDirectory)
$ErrorActionPreference='Stop';$root=(Get-Location).Path
. (Join-Path $root 'tools/build_workbook/SettingsCatalog.ps1')
$path=(Resolve-Path -LiteralPath $WorkbookPath).Path;$directory=[IO.Path]::GetFullPath((Join-Path $root $EvidenceDirectory));[void][IO.Directory]::CreateDirectory($directory)
if(Test-Path (Join-Path $directory 'Package.json')){throw 'Primary package proof already exists.'}
$memory=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory;if($memory.AvailableMBytes-lt4096-or($memory.CommitLimit-$memory.CommittedBytes)-lt7GB){throw 'Insufficient memory reserve'}
$memory|Select-Object AvailableMBytes,CommitLimit,CommittedBytes|ConvertTo-Json|Set-Content (Join-Path $directory 'Memory.json') -Encoding UTF8
$hash=(Get-FileHash -LiteralPath $path).Hash;$checks=New-Object 'System.Collections.Generic.List[object]'
function Test-OpeningPackageCheck([string]$Name,[bool]$Passed,[object]$Actual,[object]$Expected){$checks.Add([ordered]@{Name=$Name;Passed=$Passed;Actual=$Actual;Expected=$Expected})}
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class OpeningPackageTestWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
$app=$null;$book=$null;$owned=0;$ticks=0
try{
 $previous=@(Get-Process EXCEL -ErrorAction SilentlyContinue|ForEach-Object Id);$app=New-Object -ComObject Excel.Application
 $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=3
 [uint32]$id=0;[void][OpeningPackageTestWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id);if($previous-contains[int]$id){$app=$null;throw 'Independent Excel required'}
 $owned=[int]$id;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
 @{PID=$owned;StartTicksUTC=[string]$ticks}|ConvertTo-Json|Set-Content (Join-Path $directory 'Ownership.json') -Encoding UTF8
 $book=$app.Workbooks.Open($path,0,$true);if(-not $book.ReadOnly){throw 'Read-only artifact opening required.'}
 $layout=Read-WorkbookLayout
 foreach($def in @(Get-GeometrySettingsCatalog)){
  $range=$book.Names.Item($def.RangeName).RefersToRange;$expected=@($layout.Names|Where-Object Name -eq $def.RangeName)[0]
  Test-OpeningPackageCheck ($def.Type+'.address') ($range.Address()-ceq$expected.Address) ($range.Address()) $expected.Address
  Test-OpeningPackageCheck ($def.Type+'.rowheight') ($range.RowHeight-eq15) $range.RowHeight 15
  Test-OpeningPackageCheck ($def.Type+'.nowrap') ($range.WrapText-eq$false) $range.WrapText $false
  Test-OpeningPackageCheck ($def.Type+'.static') ($range.FormatConditions.Count-eq0) $range.FormatConditions.Count 0
  foreach($edge in 7,8,9,10,11,12){Test-OpeningPackageCheck ($def.Type+'.border.'+$edge) ($range.Borders.Item($edge).LineStyle-eq-4115) $range.Borders.Item($edge).LineStyle -4115}
  $blocks=@(Get-GeometryTableBlocks $range)
  for($i=0;$i-lt$blocks.Count;$i++){
   $block=$blocks[$i];$expectedBlock=$def.Blocks[$i]
   Test-OpeningPackageCheck ($def.Type+'.block.'+$i) ($block.Key-ceq$expectedBlock.Key) $block.Key $expectedBlock.Key
   $first=$block.FirstRow;if($expectedBlock.Units){$first++}
   Test-OpeningPackageCheck ($def.Type+'.rows.'+$block.Key) (($block.LastRow-$first+1)-eq$expectedBlock.Rows.Count) ($block.LastRow-$first+1) $expectedBlock.Rows.Count
  }
 }
 $sheet=$book.Worksheets.Item('Сечения');$formulas=@{C7='=Расчет!G29';C8='=Расчет!G30';C9='=Расчет!G31';C16='=Расчет!G40';D16='=Расчет!H40';C23='=Расчет!G52';D23='=Расчет!H52';G23='=Расчет!I52';H23='=Расчет!J52'}
 foreach($address in $formulas.Keys){$cell=$sheet.Range($address);Test-OpeningPackageCheck ('link.'+$address) ($cell.Formula-ceq$formulas[$address]) $cell.Formula $formulas[$address];Test-OpeningPackageCheck ('derived.fill.'+$address) ($cell.Interior.Color-ne15135487) $cell.Interior.Color 'Not input fill'}
 foreach($address in @('E16','F16','G16','H16')){Test-OpeningPackageCheck ('service.fill.'+$address) ($sheet.Range($address).Interior.Color-ne15135487) $sheet.Range($address).Interior.Color 'Not input fill'}
 $separator=[string]$app.International(5)
 foreach($address in @('E23','F23','I23','J23')){
  $cell=$sheet.Range($address);$old=$cell.Formula;$options=@('EachBar','EverySecondBar');if($address-in@('F23','J23')){$options=@('Stacked','SideBySide')}
  Test-OpeningPackageCheck ('editable.'+$address) ($cell.Interior.Color-eq15135487-and$cell.HorizontalAlignment-eq-4108) $cell.Interior.Color 15135487
  Test-OpeningPackageCheck ('validation.'+$address) ($cell.Validation.Type-eq3-and$cell.Validation.ShowError-and$cell.Validation.Formula1-ceq($options-join$separator)) $cell.Validation.Formula1 ($options-join$separator)
  foreach($option in $options){$cell.Value2=$option;Test-OpeningPackageCheck ('programmatic.accept.'+$address+'.'+$option) ([bool]$cell.Validation.Value) $cell.Value2 $option}
  foreach($invalid in @('__INVALID__',($options-join';'),($options-join','))){$cell.Value2=$invalid;Test-OpeningPackageCheck ('programmatic.reject.'+$address+'.'+$invalid) (-not[bool]$cell.Validation.Value) $cell.Value2 'Invalid'}
  $cell.Formula=$old
 }
 $book.Close($false);$book=$null
 $passed=@($checks|Where-Object{-not $_.Passed}).Count-eq0
 @{Status=$(if($passed){'PASS'}else{'FAIL'});SourceSHA256=$hash;Checks=$checks;NativeDropdownAcceptance='Delegated to user explicitly; programmatic checks are not native UI acceptance'}|ConvertTo-Json -Depth 6|Set-Content (Join-Path $directory 'Package.json') -Encoding UTF8
 Write-Output ('PACKAGE: '+$passed+'; checks='+$checks.Count)
 if(-not $passed){throw 'Package checks failed.'}
}finally{
 if($book){$book.Close($false)}
 if($app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
 if($owned){$process=Get-Process -Id $owned -ErrorAction SilentlyContinue;if($process-and$process.StartTime.ToUniversalTime().Ticks-eq$ticks-and$process.ProcessName-eq'EXCEL'){Stop-Process -Id $owned -Force}}
 if((Get-FileHash -LiteralPath $path).Hash-ne$hash){throw 'Read-only package check changed source file.'}
}
