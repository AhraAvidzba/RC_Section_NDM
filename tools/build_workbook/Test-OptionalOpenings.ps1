# ДЛЯ ТЕСТОВ: запускает только проверки отверстий и смежных рецептов на частной копии готовой книги.
param([Parameter(Mandatory=$true)][string]$SourceWorkbook,[Parameter(Mandatory=$true)][string]$EvidenceDirectory,[switch]$HollowControlOnly)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$source=(Resolve-Path -LiteralPath (Join-Path $root $SourceWorkbook)).Path
$directory=[IO.Path]::GetFullPath((Join-Path $root $EvidenceDirectory))
if(-not $directory.StartsWith((Join-Path $root 'docs/regression/CircleOpening_') , [StringComparison]::OrdinalIgnoreCase)){throw 'Частные доказательства должны находиться в каталоге CircleOpening.'}
if(Test-Path -LiteralPath (Join-Path $directory 'Started.json')){throw 'Первичный запуск не перезаписывается.'}
[void][IO.Directory]::CreateDirectory($directory)
$memory=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
if($memory.AvailableMBytes -lt 4096 -or ($memory.CommitLimit-$memory.CommittedBytes) -lt 7GB){throw 'Недостаточно памяти для нового Excel.'}
$memory|Select-Object AvailableMBytes,CommitLimit,CommittedBytes|ConvertTo-Json|Set-Content (Join-Path $directory 'Memory.json') -Encoding UTF8
$hash=(Get-FileHash -LiteralPath $source).Hash
@{Source=$source;SHA256=$hash;StartedUTC=[DateTimeOffset]::UtcNow.ToString('o')}|ConvertTo-Json|Set-Content (Join-Path $directory 'Started.json') -Encoding UTF8
$path=Join-Path $directory 'RC_Section_NDM.xlsm';Copy-Item -LiteralPath $source -Destination $path
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class OptionalOpeningsTestWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
$app=$null;$book=$null;$owned=0;$ticks=0;$results=@()
try{
 $previous=@(Get-Process EXCEL -ErrorAction SilentlyContinue|ForEach-Object Id)
 $app=New-Object -ComObject Excel.Application;$app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=1
 [uint32]$id=0;[void][OptionalOpeningsTestWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id)
 if($previous -contains [int]$id){$app=$null;throw 'Нужен собственный независимый Excel.'}
 $owned=[int]$id;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
 @{PID=$owned;StartTicksUTC=[string]$ticks}|ConvertTo-Json|Set-Content (Join-Path $directory 'Ownership.json') -Encoding UTF8
 $book=$app.Workbooks.Open($path,0,$false);if($book.ReadOnly){throw 'Частная книга открылась только для чтения.'}
 $app.VBE.ActiveVBProject=$book.VBProject;$control=$app.VBE.CommandBars.FindControl(1,578)
 if($control.Enabled){$control.Execute()};if($control.Enabled){throw 'Компиляция VBA не завершилась.'}
 # Канонический нативный экспорт читается из готового проекта, без импорта исходников.
 $text=New-Object Text.StringBuilder
 [void]$text.AppendLine('VBA PROJECT EXPORT');[void]$text.AppendLine('Workbook: '+$book.Name);[void]$text.AppendLine('Date: '+(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
 foreach($component in $book.VBProject.VBComponents){
  $lines=$component.CodeModule.CountOfLines
  [void]$text.AppendLine('='*100);[void]$text.AppendLine('COMPONENT: '+$component.Name);[void]$text.AppendLine('TYPE: '+$component.Type);[void]$text.AppendLine('LINES: '+$lines);[void]$text.AppendLine('='*100)
  if($lines -gt 0){[void]$text.AppendLine($component.CodeModule.Lines(1,$lines))};[void]$text.AppendLine()
 }
 [IO.File]::WriteAllText((Join-Path $directory 'VBA_All_Code.txt'),$text.ToString(),[Text.UTF8Encoding]::new($true))
 $macros=@('modTestOptionalOpenings.RunOptionalOpeningTests','modTestOptionalOpenings.RunOptionalOpeningWorkbookTests','modTestSectionRecipes.RunCircleRecipeInputTests','modTestSectionRecipes.RunHollowRecipeInputTests','modTestSectionRecipes.RunCircleBindingTests','modTestRecipeAnnotations.RunArbitraryAnnotationGeometryTests','modTestSectionRecipes.RunRecipeRuleTests','modTestRecipeOperations.RunRecipeOperationTests')
 if($HollowControlOnly){$macros=@()}
 foreach($macro in $macros){
  $start=[Diagnostics.Stopwatch]::StartNew();$body='';$errorText=''
  try{$body=[string]$app.Run("'"+$book.Name+"'!"+$macro)}catch{$errorText=$_.Exception.ToString()}
  $start.Stop();$body|Set-Content -LiteralPath (Join-Path $directory ($macro+'.txt')) -Encoding UTF8
  $ok=($errorText-eq'' -and $body-notmatch '(?m)^FAIL' -and $body-match 'failed=0')
  $results+=,[ordered]@{Macro=$macro;Passed=$ok;Seconds=$start.Elapsed.TotalSeconds;Error=$errorText}
  $results|ConvertTo-Json -Depth 5|Set-Content (Join-Path $directory 'Progress.json') -Encoding UTF8
  Write-Output ($macro+': '+$ok)
 }
 foreach($case in @(@('HollowMatrix','modTestSectionRecipes.TestDumpHollowRecipeMatrix'),@('HollowEdgeMatrix','modTestSectionRecipes.TestDumpHollowRecipeEdgeMatrix'))){
  $body=[string]$app.Run("'"+$book.Name+"'!"+$case[1]);$body|Set-Content -LiteralPath (Join-Path $directory ($case[0]+'.txt')) -Encoding UTF8
  $results+=,[ordered]@{Macro=$case[1];Passed=($body-notmatch '(?m)^FAIL' -and $body-match 'HollowMatrix');Seconds=0;Error=''}
 }
 if(-not $HollowControlOnly){
 $body=[string]$app.Run("'"+$book.Name+"'!modTestOptionalOpenings.TestOptionalOpeningPreview",$directory)
 $body|Set-Content -LiteralPath (Join-Path $directory 'Preview.txt') -Encoding UTF8
 $results+=,[ordered]@{Macro='Preview';Passed=(Test-Path -LiteralPath (Join-Path $directory 'CircleOpening.png'));Seconds=0;Error=''}
 }
 $book.Save();$book.Close($false);$book=$null
 @{Status=$(if(@($results|Where-Object{-not $_.Passed}).Count){'FAIL'}else{'PASS'});SourceSHA256=$hash;Results=$results}|ConvertTo-Json -Depth 6|Set-Content (Join-Path $directory 'Acceptance.json') -Encoding UTF8
 if(@($results|Where-Object{-not $_.Passed}).Count){throw 'Есть отказы; полный адресный набор завершён.'}
}finally{
 if($book){$book.Close($false)}
 if($app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
 if($owned){$process=Get-Process -Id $owned -ErrorAction SilentlyContinue;if($process -and $process.StartTime.ToUniversalTime().Ticks-eq$ticks -and $process.ProcessName-eq'EXCEL'){Stop-Process -Id $owned -Force}}
 if((Get-FileHash -LiteralPath $source).Hash -ne $hash){throw 'Проверка изменила исходную книгу.'}
}
