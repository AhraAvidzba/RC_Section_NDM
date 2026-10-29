# -------------------- ДЛЯ ТЕСТОВ --------------------
# Обновляет три общих геометрических класса в отдельной копии чистой Build03; данные пользователя не переносятся.
$ErrorActionPreference='Stop'
Import-Module 'C:/Windows/System32/WindowsPowerShell/v1.0/Modules/Microsoft.PowerShell.Utility/Microsoft.PowerShell.Utility.psd1' -Global
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path
Set-Location -LiteralPath $root
$out=Join-Path $PSScriptRoot 'Build04'
if(Test-Path -LiteralPath $out){throw 'Fresh candidate directory required'}
[void][IO.Directory]::CreateDirectory($out)
$path=Join-Path $out 'RC_Section_NDM.xlsm'
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'Build03/RC_Section_NDM.xlsm') -Destination $path
$m=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
if($m.AvailableMBytes -lt8192 -or ($m.CommitLimit-$m.CommittedBytes)-lt14GB){throw 'Memory reserve insufficient'}
Add-Type 'using System;using System.Runtime.InteropServices;public static class TestUpdateOwner04{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
$app=$null;$book=$null;$ownNew=$false
try{
 $pre=@([Diagnostics.Process]::GetProcessesByName('EXCEL')|ForEach-Object {$_.Id})
 $app=New-Object -ComObject Excel.Application
 [uint32]$id=0;[void][TestUpdateOwner04]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id)
 if($id-le0 -or $pre-contains$id){$app=$null;throw 'A new own Excel instance required'}
 $ownNew=$true;$owned=[int]$id;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
 [ordered]@{PID=$owned;StartTicksUTC=[string]$ticks;Base='Build03';Output=$path}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $out 'BuildOwnership.json') -Encoding UTF8
 $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=3
 $book=$app.Workbooks.Open($path,0,$false);if($book.ReadOnly){throw 'Writable private candidate required'}
 $sources=@()
 foreach($file in @('src/Geometry/CGeometryPath.cls','src/Geometry/CGeometryVariableOffset.cls','src/Section/CSectionGeometryQuery.cls')){
  $source=[IO.File]::ReadAllText((Join-Path $root $file),[Text.Encoding]::UTF8)
  if($source -notmatch 'Attribute VB_Name = "([^"]+)"'){throw 'Component name missing'}
  $name=$Matches[1];$body=[Collections.Generic.List[string]]::new();$header=$false
  foreach($line in ($source -split'\r?\n')){
   if($line -match '^\uFEFF?VERSION\s'){continue}
   if($line -match '^BEGIN\s*$'){$header=$true;continue}
   if($header){if($line -match '^END\s*$'){$header=$false};continue}
   if($line -match '^\uFEFF?Attribute\s+VB_'){continue}
   $body.Add($line)
  }
  $code=$book.VBProject.VBComponents.Item($name).CodeModule
  if($code.CountOfLines-gt0){$code.DeleteLines(1,$code.CountOfLines)}
  $code.AddFromString(($body-join"`r`n").Trim());$code=$null
  $sources+=@{Path=$file;SHA256=(Get-FileHash -LiteralPath (Join-Path $root $file)).Hash}
 }
 $app.VBE.ActiveVBProject=$book.VBProject;$command=$app.VBE.CommandBars.FindControl(1,578)
 if($null-eq$command){throw 'Compile command578 missing'}
 if($command.Enabled){$command.Execute()}
 if($command.Enabled){throw 'Compile did not finish'}
 $book.Save();$book.Close($false);$book=$null
 [ordered]@{Status='COMPILED_CODE_UPDATE_OF_CLEAN_BUILD03';SourceUserDataCopied=$false;ChangedComponents=$sources;BookSHA256=(Get-FileHash -LiteralPath $path).Hash;FinishedUTC=[DateTimeOffset]::UtcNow.ToString('o')}|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $out 'Build.json') -Encoding UTF8
}finally{
 $code=$null;$command=$null;[GC]::Collect();[GC]::WaitForPendingFinalizers()
 if($book){$book.Close($false)};$book=$null
 if($app-and$ownNew){
  $p=Get-Process -Id $owned -ErrorAction Stop
  if([string]$p.StartTime.ToUniversalTime().Ticks -cne[string]$ticks){throw 'Own identity changed; no Quit performed'}
  $app.Quit()
 }
 $app=$null;[GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect();[GC]::WaitForPendingFinalizers()
 if($ownNew){[ordered]@{PID=$owned;StartTicksUTC=[string]$ticks;QuitRequested=$true;ForceKilled=$false}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $out 'Quit.json') -Encoding UTF8}
}