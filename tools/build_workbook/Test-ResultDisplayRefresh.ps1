# ДЛЯ ТЕСТОВ: штатный расчёт без F9, перезаписи формул статуса и повторного открытия.
param([Parameter(Mandatory=$true)][string]$WorkbookPath,[Parameter(Mandatory=$true)][string]$EvidenceDirectory)
$ErrorActionPreference='Stop';$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$source=(Resolve-Path -LiteralPath (Join-Path $root $WorkbookPath)).Path
$evidence=Join-Path $root $EvidenceDirectory;[void][IO.Directory]::CreateDirectory($evidence)
$allowed=[IO.Path]::GetFullPath((Join-Path $root 'docs/regression'))+[IO.Path]::DirectorySeparatorChar
if(-not [IO.Path]::GetFullPath($evidence).StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)){throw 'Evidence must be inside docs/regression'}
$private=$evidence;$path=Join-Path $evidence 'RefreshFixture.xlsm'
if(Test-Path -LiteralPath $path){throw 'Test fixture exists; use a fresh evidence directory'}
$sourceHash=(Get-FileHash -LiteralPath $source).Hash;Copy-Item -LiteralPath $source -Destination $path
$hash=(Get-FileHash -LiteralPath $path).Hash
$app=$null;$book=$null;$owned=0;$ticks=0;$script:checks=0;$cases=@()
Add-Type 'using System;using System.Runtime.InteropServices;public static class RefreshTestWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Прерывает направленную проверку при первом нарушении фактического результата.
function Test-RefreshAssert([bool]$Passed,[string]$Reason){if(-not $Passed){throw $Reason};$script:checks++}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Меняет настройку только по структурному ключу таблицы в проверочной копии.
function Test-RefreshSetting([string]$Key,[string]$Value){
 $table=$book.Names.Item('rngSystemSettings').RefersToRange
 for($row=2;$row -le $table.Rows.Count;$row++){if([string]$table.Cells.Item($row,1).Value2 -ceq $Key){$table.Cells.Item($row,2).Value2=$Value;return}}
 throw ('Missing setting: '+$Key)
}
try{
 $memory=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
 if($memory.CommitLimit-$memory.CommittedBytes -lt 7GB){throw 'Insufficient memory'}
 $app=New-Object -ComObject Excel.Application;$app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=1
 [uint32]$id=0;[void][RefreshTestWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id);$owned=[int]$id;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
 Write-Output ('OWNED_EXCEL: '+$owned+'; ticks='+$ticks)
 foreach($storage in @('Excel','File')){foreach($mode in @(-4105,-4135)){
  $app.EnableEvents=$false;$book=$app.Workbooks.Open($path,0,$true);$app.Calculation=-4105
  $prefix="'"+$book.Name+"'!";$calc=$book.Worksheets.Item('Расчет')
  Test-RefreshAssert $book.ReadOnly 'Candidate must remain read-only'
  Test-RefreshAssert ([string]$calc.Range('AX10').Value2 -ceq 'шаг вдоль оси сечения') 'AX10 caption'
  $table=$book.Names.Item('rngTorsionSettings').RefersToRange;$band=$null
  for($row=1;$row -le $table.Rows.Count;$row++){if([string]$table.Cells.Item($row,1).Value2 -ceq 'Torsion.BarBand'){$band=$table.Cells.Item($row,2).Value2}}
  Test-RefreshAssert ($band -eq 190) 'BarBand default'
  foreach($address in @('Q44:V58','Y44:AA58','AS8:AS22','AI7:AK7','AH27:AJ27','AH39:AJ39','AY18:BB18','AH48:AJ49','AH53:AJ54')){
   Test-RefreshAssert ([string]$calc.Range($address).NumberFormat -ceq '0.0') ('NumberFormat: '+$address)
  }
  $before=@{};foreach($cell in $calc.Range('A1:BD59').SpecialCells(-4123)){$before[$cell.Address()]=$cell.Formula}
  if($storage -ceq 'Excel' -and $mode -eq -4105){
   $project=$book.VBProject;$app.VBE.ActiveVBProject=$project;$compile=$app.VBE.CommandBars.FindControl(1,578)
   if($compile.Enabled){$compile.Execute()};Test-RefreshAssert (-not $compile.Enabled) 'VBA compile';Test-RefreshAssert ($project.VBComponents.Count -eq 203) 'VBA component count'
   $text=New-Object Text.StringBuilder;[void]$text.AppendLine('VBA PROJECT EXPORT');[void]$text.AppendLine('Workbook: RC_Section_NDM.xlsm')
   foreach($component in $project.VBComponents){
    $module=$component.CodeModule;[void]$text.AppendLine(('='*100));[void]$text.AppendLine('COMPONENT: '+$component.Name)
    [void]$text.AppendLine('TYPE: '+$component.Type);[void]$text.AppendLine('LINES: '+$module.CountOfLines);[void]$text.AppendLine(('='*100))
    if($module.CountOfLines){[void]$text.AppendLine($module.Lines(1,$module.CountOfLines))};[void]$text.AppendLine()
   }
   [IO.File]::WriteAllText((Join-Path $private 'VBA_All_Code.txt'),$text.ToString(),[Text.UTF8Encoding]::new($false))
  }
  foreach($pair in @(@('General.NonCriticalMessagesEnabled','No'),@('General.ExecutionReportEnabled','No'),@('Results.SnapshotStorage',$storage),@('Results.SnapshotName',('StatusRefreshTest'+$storage+$mode)),@('Plot.Enabled','Yes'))){Test-RefreshSetting $pair[0] $pair[1]}
  $loads=$book.Names.Item('rngLoadCombinations').RefersToRange;$torque=$book.Names.Item('rngTorsionCombinations').RefersToRange
  $loads.Offset(1,0).Resize($loads.Rows.Count-1,$loads.Columns.Count).ClearContents();$torque.Offset(1,0).Resize($torque.Rows.Count-1,2).ClearContents()
  $loads.Cells.Item(2,1).Value2='1';$loads.Cells.Item(2,5).Value2='PR1';$torque.Cells.Item(2,1).Value2='1'
  $formed=[string]$app.Run($prefix+'modWorkbookCalculation.FormSectionForWorkbook',$book)
  Test-RefreshAssert ($formed.Contains('Сечение сформировано')) 'Section formation'
  $app.EnableEvents=$true;$app.Calculation=$mode
  foreach($value in @(1.0,200.0,0.0)){
   $torque.Cells.Item(2,2).Value2=[double]$value
   $message=[string]$app.Run($prefix+'modWorkbookCalculation.RunSectionCalculationForWorkbook',$book)
   $actual=@(foreach($address in @('BB30','BB37','BB44')){
    $cell=$calc.Range($address);$status=[string]$cell.Value2;$color=[long]$cell.DisplayFormat.Interior.Color
    $expected=if($value -eq 0){'-'}elseif($value -eq 1){'OK'}else{'FAIL'}
    Test-RefreshAssert ($status -ceq $expected) ($storage+' '+$mode+' T='+$value+' '+$address+' status: '+$status)
    Test-RefreshAssert ($color -eq $(if($expected -ceq 'FAIL'){255}else{5296274})) ($storage+' '+$mode+' T='+$value+' '+$address+' color: '+$color)
    [ordered]@{Cell=$address;Status=$status;Color=$color}
   })
   Test-RefreshAssert ($app.Calculation -eq $mode -and $app.EnableEvents -and $app.ScreenUpdating) 'Application state not restored'
   $cases+=,[ordered]@{Storage=$storage;Calculation=$mode;T=$value;Statuses=$actual;Message=$message}
   Write-Output ('PASS '+$storage+' mode='+$mode+' T='+$value)
  }
  foreach($address in $before.Keys){Test-RefreshAssert ([string]$calc.Range($address).Formula -ceq [string]$before[$address]) ('Formula changed by calculation: '+$address)}
  $book.Close($false);$book=$null
 }}
 Test-RefreshAssert ((Get-FileHash -LiteralPath $path).Hash -ceq $hash) 'Candidate changed'
 Test-RefreshAssert ((Get-FileHash -LiteralPath $source).Hash -ceq $sourceHash) 'Source workbook changed'
 [ordered]@{Status='PASS';Assertions=$checks;CandidateSHA256=$hash;Cases=$cases;UIVisualAcceptance='User will verify; hidden native Excel checks only';PID=$owned;StartTicksUTC=$ticks} | ConvertTo-Json -Depth 9 | Set-Content -Encoding UTF8 -LiteralPath (Join-Path $evidence 'NativeRefresh.json')
 Write-Output ('PASS: assertions='+$checks)
}catch{
 [ordered]@{Reason=$_.Exception.Message;Stack=$_.ScriptStackTrace;Position=$_.InvocationInfo.PositionMessage;Cases=$cases} | ConvertTo-Json -Depth 8 | Set-Content -Encoding UTF8 -LiteralPath (Join-Path $evidence 'NativeFailure.json');throw
}finally{
 if($book){$book.Close($false)};if($app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
 if($owned){$p=Get-Process -Id $owned -ErrorAction SilentlyContinue;if($p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.ToUniversalTime().Ticks -eq $ticks){Stop-Process -Id $owned -Force -ErrorAction SilentlyContinue}}
}
