# ==========================================================================
# ДЛЯ ТЕСТОВ: прочерки при отсутствии Worst, точность нагрузок и цвет растянутой N.
# ==========================================================================
param([Parameter(Mandatory=$true)][string]$WorkbookPath,[Parameter(Mandatory=$true)][string]$EvidenceDirectory,[switch]$FreshDefaults)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$path=(Resolve-Path -LiteralPath (Join-Path $root $WorkbookPath)).Path
$directory=Join-Path $root $EvidenceDirectory;[void][IO.Directory]::CreateDirectory($directory)
$sourceHash=(Get-FileHash -LiteralPath $path).Hash
$app=$null;$book=$null;$owned=0;$ticks=0;$script:testCount=0;$signCases=@()

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Останавливает проверку при фактическом несоответствии, не подменяя его предупреждением.
function Test-DisplayAssert([bool]$Condition,[string]$Message){if(-not $Condition){throw $Message};$script:testCount++}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Находит значение настройки по ключу; адресом ячейки не предполагается её положение.
function Test-DisplaySetting([object]$Workbook,[string]$Key){
 $range=$Workbook.Names.Item('rngSystemSettings').RefersToRange
 for($row=2;$row -le $range.Rows.Count;$row++){if([string]$range.Cells.Item($row,1).Value2 -ceq $Key){return $range.Cells.Item($row,2)}}
 throw ('Setting not found: '+$Key)
}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Получает только настоящие ошибки формул в печатной области.
function Test-DisplayErrors([object]$Sheet){
 try{$range=$Sheet.Range('A1:BD59').SpecialCells(-4123,16)}catch{return @()}
 return @($range | ForEach-Object {$_.Address($false,$false)+'='+$_.Text})
}

Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class DisplayTestWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
try{
 $memory=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
 if([double]$memory.CommitLimit-[double]$memory.CommittedBytes -lt 4GB){throw 'Insufficient memory'}
 $app=New-Object -ComObject Excel.Application;$app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=3
 [uint32]$id=0;[void][DisplayTestWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id);$owned=[int]$id;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
 Write-Output ('OWNED_EXCEL: '+$owned+'; ticks='+$ticks)
 $book=$app.Workbooks.Open($path,0,$true);Test-DisplayAssert $book.ReadOnly 'Source must be read-only'
 $calc=$book.Worksheets.Item('Расчет');$results=$book.Worksheets.Item('Results');$calc.Activate();$app.CalculateFull()
 $beforeErrors=@(Test-DisplayErrors $calc);Test-DisplayAssert ($beforeErrors.Count -eq 0) ('Initial formula errors: '+($beforeErrors -join ';'))
 Test-DisplayAssert ([string]$results.Range('AL8').Value2 -ceq "Нагрузки с учётом продольного`nизгиба (если применимо)") 'Wrong Results header line break'
 foreach($address in @('Q44:V58','Y44:AA58','AS8:AS22','AI7:AK7','AH27:AJ27','AH39:AJ39','AY18:BB18','AH48:AJ49','AH53:AJ54')){
  Test-DisplayAssert ([string]$calc.Range($address).NumberFormat -ceq '0.0') ('Wrong precision: '+$address)
 }
 $ids=$results.Range('G12:R12');$savedIDs=$ids.Formula;$ids.ClearContents();$app.CalculateFull()
 foreach($address in @('AH7','AG27','AG39','AX18','AU35','AU42','BI26','BI27','BI31','BI32','BI47','BI52')){
  Test-DisplayAssert ([string]$calc.Range($address).Value2 -ceq '-') ('Missing ID is not dash: '+$address)
 }
 foreach($address in @('AI7','AJ7','AK7','AD12','AH27','AI27','AJ27','AM27','AH39','AI39','AJ39','AN39','AO48','AO49','AO53','AO54')){
  Test-DisplayAssert ([string]$calc.Range($address).Value2 -ceq '-') ('Missing result is not dash: '+$address)
 }
 $noneErrors=@(Test-DisplayErrors $calc);Test-DisplayAssert ($noneErrors.Count -eq 0) ('No-Worst formula errors: '+($noneErrors -join ';'))
 $ids.Formula=$savedIDs
 $zeroSaved=@($results.Range('G12').Formula,$results.Range('A13').Formula,$results.Range('G49').Formula)
 $results.Range('G12').Value2=0;$results.Range('A13').Value2=0;$results.Range('G49').Value2=0;$app.CalculateFull()
 Test-DisplayAssert ([double]$calc.Range('AH7').Value2 -eq 0 -and [double]$calc.Range('AI7').Value2 -eq 0) 'Valid ID0 and numeric0 must not disappear'
 $results.Range('G12').Formula=$zeroSaved[0];$results.Range('A13').Formula=$zeroSaved[1];$results.Range('G49').Formula=$zeroSaved[2]
 $signRange=$book.Names.Item('rngSignConventionSettings').RefersToRange;$signCell=$null
 for($row=2;$row -le $signRange.Rows.Count;$row++){if([string]$signRange.Cells.Item($row,1).Value2 -ceq '+N'){$signCell=$signRange.Cells.Item($row,2);break}}
 Test-DisplayAssert ($null -ne $signCell) 'INPUT sign cell not found';$savedSign=$signCell.Formula
 foreach($sign in @('Tension','Compression')){
  $signCell.Value2=$sign
  foreach($column in @('Q','T','Y')){
   $cell=$calc.Range($column+'44');$saved=$cell.Formula
   foreach($value in @(1.234567,-1.234567,0)){
    $cell.Value2=[double]$value;$app.CalculateFull()
    $actual=[long]$cell.DisplayFormat.Interior.Color
    $expected=($value -ne 0 -and (($sign -eq 'Tension' -and $value -gt 0) -or ($sign -eq 'Compression' -and $value -lt 0)))
    Test-DisplayAssert (($actual -eq 6737151) -eq $expected) ('Wrong tension color: '+$sign+' '+$column+' '+$value)
    Test-DisplayAssert ([double]$cell.Value2 -eq [double]$value) 'Display format changed a stored load'
    $signCases+=,[ordered]@{Sign=$sign;Cell=$cell.Address();Value=$value;Tension=$expected;Color=$actual;Text=$cell.Text}
   }
   $cell.Formula=$saved
  }
 }
 $signCell.Formula=$savedSign;$app.CalculateFull()
 foreach($address in @('AN20','AN34','AN42','AN58','AO58','BB30','BB37','BB44')){
  $cell=$calc.Range($address);$saved=$cell.Formula
  $cell.Value2='OK';$app.CalculateFull();$green=[long]$cell.DisplayFormat.Interior.Color
  $cell.Value2='-';$app.CalculateFull();Test-DisplayAssert ([long]$cell.DisplayFormat.Interior.Color -eq $green) ('Dash must stay green: '+$address)
  $cell.Value2='FAIL';$app.CalculateFull();Test-DisplayAssert ([long]$cell.DisplayFormat.Interior.Color -ne $green) ('Failure must retain warning color: '+$address)
  $cell.Formula=$saved
 }
 $radius=Test-DisplaySetting $book 'SLS.Crack.SP35.InteractionRadiusMode'
 Test-DisplayAssert ([string]$radius.Formula -ceq '=Расчет!AA40') 'Interaction radius link lost'
 Test-DisplayAssert ($calc.Range('AA40').Validation.Type -eq 3) 'Radius validation list lost'
 if($FreshDefaults){foreach($key in @('Stability.Code','SLS.Crack.Code')){Test-DisplayAssert ([string](Test-DisplaySetting $book $key).Value2 -ceq 'SP35') ('Wrong fresh default: '+$key)}}
 $calc.ExportAsFixedFormat(0,(Join-Path $directory 'Calculation.pdf'))
 [ordered]@{Status='PASS';Assertions=$script:testCount;SourceSHA256=$sourceHash;SignCases=$signCases;InitialFormulaErrors=$beforeErrors;NoWorstFormulaErrors=$noneErrors;PID=$owned;StartTicksUTC=$ticks;FreshDefaults=[bool]$FreshDefaults} | ConvertTo-Json -Depth 9 | Set-Content -LiteralPath (Join-Path $directory 'NativeDisplayReview.json') -Encoding UTF8
 Write-Output ('PASS: assertions='+$script:testCount)
}finally{
 if($book){$book.Close($false)};if($app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
 if($owned){$p=Get-Process -Id $owned -ErrorAction SilentlyContinue;if($p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.ToUniversalTime().Ticks -eq $ticks){Stop-Process -Id $owned -Force}}
}
if((Get-FileHash -LiteralPath $path).Hash -cne $sourceHash){throw 'Read-only verification changed source'}
