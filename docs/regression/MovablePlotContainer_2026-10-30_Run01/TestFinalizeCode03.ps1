# -------------------- ДЛЯ ТЕСТОВ --------------------
# Завершает VBA чистой сборки без изменения листов; проверяет предупреждение при пустом снимке без сохранения тестовых изменений.
$ErrorActionPreference='Stop'
Import-Module 'C:/Windows/System32/WindowsPowerShell/v1.0/Modules/Microsoft.PowerShell.Utility/Microsoft.PowerShell.Utility.psd1' -Global
$p=$PSScriptRoot;$out=Join-Path $p 'Final03';if(Test-Path -LiteralPath $out){throw 'Inspect existing finalization first'}
[void][IO.Directory]::CreateDirectory($out)
$path=Join-Path $out 'RC_Section_NDM.xlsm';Copy-Item -LiteralPath (Join-Path $p 'Build02/RC_Section_NDM.xlsm') -Destination $path
$root='C:/Users/avidzba/Desktop/программы/В_РАБОТЕ!/RC_Section_NDM'
$m=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
if($m.AvailableMBytes-lt8192-or($m.CommitLimit-$m.CommittedBytes)-lt14GB){throw 'Insufficient memory reserve'}
Add-Type 'using System;using System.Runtime.InteropServices;public static class TestFinal03Window{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
$app=$null;$book=$null;$own=$false;$quit=$false;$result=[ordered]@{Passed=$false;FreshBase='Build02/RC_Section_NDM.xlsm';UserDataCopied=$false}
try{
 $existing=@([Diagnostics.Process]::GetProcessesByName('EXCEL')|ForEach-Object {$_.Id})
 $app=New-Object -ComObject Excel.Application;[uint32]$id=0;[void][TestFinal03Window]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id)
 if($existing-contains[int]$id){$app=$null;throw 'New own Excel required'};$own=$true
 [ordered]@{PID=[int]$id;StartTicksUTC=[string](Get-Process -Id $id).StartTime.ToUniversalTime().Ticks}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $out 'Ownership.json') -Encoding UTF8
 $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=3
 $book=$app.Workbooks.Open($path,0,$false)
 foreach($name in 'modWorkbookCalculation','modTestPlotConfig'){
  $relative=if($name-eq'modTestPlotConfig'){'tests/'+$name+'.bas'}else{'src/Excel/'+$name+'.bas'}
  $source=[IO.File]::ReadAllText((Join-Path $root $relative),[Text.Encoding]::UTF8)
  $module=$book.VBProject.VBComponents.Item($name).CodeModule
  $module.DeleteLines(1,$module.CountOfLines);$module.AddFromString([regex]::Match($source,'(?ms)^Option Explicit.*').Value)
 }
 $app.VBE.ActiveVBProject=$book.VBProject;$compile=$app.VBE.CommandBars.FindControl(1,578)
 if($compile.Enabled){$compile.Execute()};if($compile.Enabled){throw 'VBA compilation incomplete'}
 $book.Save();$book.Close($false);$book=$null
 $result.BookSHA256=(Get-FileHash -LiteralPath $path).Hash
 $app.AutomationSecurity=1;$book=$app.Workbooks.Open($path,0,$true)
 $book.Worksheets.Item('Расчет').ChartObjects().Item('chtNDMSectionPlot').Delete()
 $prefix="'"+$book.Name+"'!"
 $text=[string]$app.Run($prefix+'modTestPlotConfig.TestEmptyPlotRecoveryMessage',$book)
 [IO.File]::WriteAllText((Join-Path $out 'EmptyRecoveryMessage.txt'),$text,[Text.UTF8Encoding]::new($false))
 if(-not$text.StartsWith('EXPECTED_ERROR|')-or$text-notmatch'автоматически создана'-or$book.Worksheets.Item('Расчет').ChartObjects().Count-ne1){throw 'Empty-snapshot recovery notice missing'}
 $second=[string]$app.Run($prefix+'modTestPlotConfig.TestEmptyPlotRecoveryMessage',$book)
 if($second-match'автоматически создана'-or$book.Worksheets.Item('Расчет').ChartObjects().Count-ne1){throw 'Existing container generated a false recovery notice'}
 $text=[string]$app.Run($prefix+'modTestPlotConfig.RunPlotContainerTests')
 [IO.File]::WriteAllText((Join-Path $out 'ContainerPrimary.txt'),$text,[Text.UTF8Encoding]::new($false))
 if($text-notmatch'passed=24; failed=0'-or$text-match'(?m)^FAIL:'){throw 'Final container suite failed'}
 $result.Passed=$true;$result.EmptySnapshotNotice=$true;$result.NoFalseSecondNotice=$true;$result.ContainerPassed=24
}catch{$result.Error=$_.Exception.ToString();throw}
finally{
 if($book){$book.Close($false)};$book=$null;$module=$null;$compile=$null
 if($app-and$own){$app.Quit();$quit=$true};$app=$null;[GC]::Collect();[GC]::WaitForPendingFinalizers()
 $result.QuitRequested=$quit;$result.SourceAfterSHA256=(Get-FileHash -LiteralPath $path).Hash
 $result|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $out 'Result.json') -Encoding UTF8
}
