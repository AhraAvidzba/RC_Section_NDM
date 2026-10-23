# ==========================================================================
# ДЛЯ ТЕСТОВ: все внутренние ссылки справки и разные состояния её разделов
# ==========================================================================
# Штатная команда Excel открывает ссылку из фактически выбранной ячейки.
# Проверка выполняется на read-only копии, без расчёта и записи исходных данных.
param([Parameter(Mandatory=$true)][string]$WorkbookPath,
      [Parameter(Mandatory=$true)][string]$ReportDirectory,
      [string]$Scenarios='Collapsed,InView,PartialHidden',
      [string]$SourceSheetFilter='',
      [switch]$VisibleExcel)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'HelpOutline.ps1')
$path=(Resolve-Path -LiteralPath $WorkbookPath).Path
$directory=[IO.Path]::GetFullPath($ReportDirectory);[IO.Directory]::CreateDirectory($directory)|Out-Null
$hash=(Get-FileHash -LiteralPath $path).Hash
$previous=@(Get-Process EXCEL -ErrorAction SilentlyContinue|ForEach-Object Id)
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class AllHelpNavigationTestWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
$app=$null;$book=$null;$owned=0;$ticks=0

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Повторно подтверждает PID и время старта перед очисткой собственного Excel.
function Test-OwnedExcel {
 if(-not $owned -or $owned -in $previous){return $false}
 $process=Get-Process -Id $owned -ErrorAction SilentlyContinue
 return $process -and $process.ProcessName -eq 'EXCEL' -and $process.StartTime.ToUniversalTime().Ticks -eq $ticks
}

try{
 $app=New-Object -ComObject Excel.Application
 [uint32]$id=0;[void][AllHelpNavigationTestWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id)
 $owned=[int]$id;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
 if(-not(Test-OwnedExcel)){throw 'Excel ownership failed.'}
 Write-Output ('OWNED_EXCEL: id='+$owned+'; utcStartTicks='+$ticks)
 $app.Visible=[bool]$VisibleExcel;$app.DisplayAlerts=$false;$app.EnableEvents=$true;$app.AutomationSecurity=1
 $book=$app.Workbooks.Open($path,0,$true)
 if(-not $book.ReadOnly){throw 'Read-only book required.'}
 $help=$book.Worksheets.Item('Справка');$main=Get-HelpMainRows $help
 $headers=@($main.Values|Sort-Object);$last=Get-HelpContentLastRow $help
 $links=@(foreach($sheet in $book.Worksheets){
  if($SourceSheetFilter -and $sheet.Name -cne $SourceSheetFilter){continue}
  foreach($link in $sheet.Hyperlinks){
  if([string]$link.SubAddress -match '^''?Справка''?!\$?([A-Z]+)\$?(\d+)$'){
   [ordered]@{Link=$link;Sheet=$sheet;Source=$link.Range.Address();Target=[string]$link.SubAddress;Row=[int]$Matches[2]}
  }
 }})
 $checks=New-Object 'System.Collections.Generic.List[object]'
 foreach($scenario in $Scenarios.Split(',')){
  foreach($record in $links){
   $target=[int]$record.Row;$first=@($headers|Where-Object {$_ -le $target})[-1]
   $next=@($headers|Where-Object {$_ -gt $target});$end=if($next.Count){[int]$next[0]-1}else{$last}
   $body=$help.Rows.Item(($first+1).ToString()+':'+$end)
   try{
    $help.Outline.ShowLevels(8)|Out-Null
    $help.Activate();$app.ActiveWindow.ScrollRow=[Math]::Max(1,$target-3);$app.ActiveWindow.ScrollColumn=1
    switch($scenario){
     'Collapsed'{$help.Outline.ShowLevels(1)|Out-Null}
     'InView'{}
     'PartialHidden'{$help.Rows.Item([Math]::Min($first+2,$end)).Hidden=$true}
     default{throw ('Unknown scenario: '+$scenario)}
    }
    if($record.Sheet.Name -ceq 'Справка'){
     $sourceFirst=@($headers|Where-Object {$_ -le $record.Link.Range.Row})[-1]
     if($record.Link.Range.EntireRow.Hidden){$help.Rows.Item([int]$sourceFirst).ShowDetail=$true} # Пользователь может нажать только видимую исходную ссылку; уже открытую группу повторно не раскрываем.
    }
    $record.Sheet.Activate();$app.Goto($record.Link.Range,$false)
    $open=@($app.CommandBars.Item('Cell').Controls|Where-Object {$_.Caption.Replace('&','') -match '^Открыть гиперссылку|^Open Hyperlink'})
    if($open.Count -ne 1 -or -not $open[0].Enabled){throw 'Native Open Hyperlink unavailable.'}
    $open[0].Execute()
    $actualRow=[int]$app.ActiveWindow.ScrollRow;$actualColumn=[int]$app.ActiveWindow.ScrollColumn
    $sheetName=[string]$app.ActiveSheet.Name;$expanded=($body.Hidden -eq $false)
    $passed=($sheetName -ceq 'Справка' -and $actualRow -eq $target -and $actualColumn -eq 1 -and $expanded)
    $checks.Add([ordered]@{Passed=$passed;Scenario=$scenario;SourceSheet=$record.Sheet.Name;Source=$record.Source;Target=$record.Target;ActualRow=$actualRow;ActualColumn=$actualColumn;WholeSectionExpanded=$expanded;ActiveSheet=$sheetName})
   }catch{
    $checks.Add([ordered]@{Passed=$false;Scenario=$scenario;SourceSheet=$record.Sheet.Name;Source=$record.Source;Target=$record.Target;Error=$_.Exception.Message})
   }
  }
  $fail=@($checks|Where-Object {-not $_.Passed})
  $checks|ConvertTo-Json -Depth 7|Set-Content -LiteralPath (Join-Path $directory ($scenario+'.json')) -Encoding UTF8
  Write-Output ('HELP_SCENARIO: '+$scenario+'; checks='+$checks.Count+'; fail='+$fail.Count)
 }
 $failed=@($checks|Where-Object {-not $_.Passed})
 [ordered]@{Status=if($failed.Count){'FAIL'}else{'PASS'};WorkbookSHA256=$hash;Links=$links.Count;Checks=$checks.Count;Failures=$failed;Details=$checks;OwnedPID=$owned;StartTicks=[string]$ticks}|
  ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $directory 'AllHelpNavigation.json') -Encoding UTF8
 Write-Output ('ALL_HELP_NAVIGATION: links='+$links.Count+'; checks='+$checks.Count+'; fail='+$failed.Count)
 if($failed.Count){throw ('Help navigation failed: '+$failed.Count)}
}catch{Write-Output $_.ScriptStackTrace;throw}
finally{
 if(Test-OwnedExcel){try{if($book){$book.Close($false)}}catch{};try{if($app){$app.Quit()}}catch{}}
 if($app){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
 if(Test-OwnedExcel){Stop-Process -Id $owned -Force -ErrorAction SilentlyContinue}
}
if((Get-FileHash -LiteralPath $path).Hash -cne $hash){throw 'Verification changed the source book.'}
