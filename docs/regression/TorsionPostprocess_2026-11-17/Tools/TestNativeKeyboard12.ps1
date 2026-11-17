param([Parameter(Mandatory=$true)][string]$CandidatePath,[Parameter(Mandatory=$true)][string]$EvidenceDirectory)
$ErrorActionPreference='Stop'
Import-Module C:/Windows/System32/WindowsPowerShell/v1.0/Modules/Microsoft.PowerShell.Utility/Microsoft.PowerShell.Utility.psd1 -Global
$directory=[IO.Path]::GetFullPath($EvidenceDirectory)
$allowed='C:\Users\avidzba\.codex\visualizations\2026\10\12\01a12ece-8192-7141-89b3-f83f1176184c\TorsionPostprocess_2026-11-17\'
if(-not $directory.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $directory)){throw 'Fresh evidence required'}
$m=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory;$o=Get-CimInstance Win32_OperatingSystem
if(($m.CommitLimit-$m.CommittedBytes)/1GB -lt 7 -or $o.FreePhysicalMemory*1KB/1GB -lt 4){throw 'Insufficient serial native reserve'}
New-Item -ItemType Directory -Path $directory|Out-Null
$path=Join-Path $directory 'RC_Section_NDM.xlsm';Copy-Item -LiteralPath $CandidatePath -Destination $path
Add-Type -TypeDefinition @'
using System;using System.Runtime.InteropServices;
public static class TestNativeKeyboardIdentity12{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}
'@
$receipt=[ordered]@{Status='RUNNING';SourceSHA256=(Get-FileHash -LiteralPath $CandidatePath -Algorithm SHA256).Hash;Own=$null;Selected=@();Current=$null;ScreenshotUnavailable=$true}
$app=$null;$book=$null;$pidn=0;$ticks=0;$failure=$null
try{
 $app=New-Object -ComObject Excel.Application;[uint32]$n=0;[void][TestNativeKeyboardIdentity12]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$n)
 $pidn=[int]$n;$ticks=(Get-Process -Id $pidn).StartTime.ToUniversalTime().Ticks
 $receipt.Own=[ordered]@{PID=$pidn;UTCticks=$ticks;QuitRequested=$false;Absent=$false;Forced=$false}
 $receipt|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $directory 'Review.json') -Encoding UTF8
 $app.Caption='NDM Keyboard List Test';$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=3
 $book=$app.Workbooks.Open($path,0,$true)
 $app.Visible=$true;$app.ScreenUpdating=$true;$app.WindowState=-4137;$book.Windows.Item(1).Visible=$true;$book.Windows.Item(1).WindowState=-4137
 $sheet=$book.Worksheets.Item('Config');$sheet.Activate();$table=$book.Names.Item('rngTorsionSettings').RefersToRange
 $counter=0
 foreach($spec in @(@{Key='Torsion.StateType';Options=@('StrengthState','CapacityState','PreCrackState','PostCrackState','CrackedState')},@{Key='Torsion.ConcreteMethod';Options=@('Strut','GorskyRectangle','None')})){
  $cell=$null;for($r=2;$r -le $table.Rows.Count;$r++){if([string]$table.Cells.Item($r,1).Value2 -ceq $spec.Key){$cell=$table.Cells.Item($r,2);break}}
  if($null -eq $cell){throw 'Missing list cell'}
  for($index=0;$index -lt $spec.Options.Count;$index++){
   $cell.Value2='__BEFORE_NATIVE_PICK__';$cell.Select();$app.ActiveWindow.ScrollRow=89;$app.ActiveWindow.ScrollColumn=32
   if([bool]$cell.Validation.Value){throw 'Invalid sentinel unexpectedly accepted'}
   $receipt.Current=[ordered]@{Case=$counter;Key=$spec.Key;Index=$index;Expected=$spec.Options[$index];Seed=[string]$cell.Value2;ActualActiveCell=[string]$app.ActiveCell.Address();ActualSheet=[string]$app.ActiveSheet.Name;ActualWindowState=[int]$app.WindowState;ScreenUpdating=[bool]$app.ScreenUpdating;InCellDropdown=[bool]$cell.Validation.InCellDropdown;Interactive=[bool]$app.Interactive;AwaitingKeys=$true}
   $receipt|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $directory 'Review.json') -Encoding UTF8
   $signal=Join-Path $directory ('Picked_'+$counter);$until=[DateTime]::UtcNow.AddMinutes(3)
   while(-not(Test-Path -LiteralPath $signal)){
    if(Test-Path -LiteralPath (Join-Path $directory 'Cancel')){throw 'UI test cancelled without acceptance'}
    if([DateTime]::UtcNow -gt $until){throw ('Native selection timeout '+$counter)}
    Start-Sleep -Milliseconds 200
   }
   $value=[string]$cell.Value2;$accepted=[bool]$cell.Validation.Value
   $receipt.Selected+=,[ordered]@{Key=$spec.Key;Index=$index;Expected=$spec.Options[$index];Actual=$value;ValidationAccepted=$accepted;SeedWasInvalid=$true;SelectedByKeyboard=$true;AfterActiveCell=[string]$app.ActiveCell.Address()}
   if($value -cne $spec.Options[$index] -or -not $accepted){throw ('Actual native selection differs '+$spec.Key+' index='+$index+' actual='+$value)}
   $counter++
  }
 }
 $receipt.Current=$null;$receipt.Status='ALL_EIGHT_NATIVE_DROPDOWN_KEYBOARD_SELECTIONS_PASS'
}catch{$failure=$_;$receipt.Status='FAILED';$receipt['Error']=[string]$_}
finally{
 if($book){try{$book.Close($false)}catch{$receipt['CloseError']=[string]$_};$book=$null}
 if($app){try{$receipt.Own.QuitRequested=$true;$app.Quit()}catch{$receipt['QuitError']=[string]$_};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app);$app=$null}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
 if($pidn){$p=Get-Process -Id $pidn -ErrorAction SilentlyContinue;if($p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.ToUniversalTime().Ticks -eq $ticks){$receipt.Own.Forced=$true;Stop-Process -Id $pidn};for($w=0;$w -lt 10 -and (Get-Process -Id $pidn -ErrorAction SilentlyContinue);$w++){Start-Sleep -Milliseconds 200};$receipt.Own.Absent=($null -eq (Get-Process -Id $pidn -ErrorAction SilentlyContinue))}
 $receipt|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $directory 'Review.json') -Encoding UTF8
}
if($failure){throw $failure}
$receipt|ConvertTo-Json -Depth 8
