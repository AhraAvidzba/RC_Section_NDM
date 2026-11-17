param([Parameter(Mandatory=$true)][string]$CandidatePath,[Parameter(Mandatory=$true)][string]$EvidenceDirectory)
$ErrorActionPreference='Stop'
Import-Module C:/Windows/System32/WindowsPowerShell/v1.0/Modules/Microsoft.PowerShell.Utility/Microsoft.PowerShell.Utility.psd1 -Global
$repo='C:\Users\avidzba\Desktop\программы\В_РАБОТЕ!\RC_Section_NDM'
$directory=[IO.Path]::GetFullPath($EvidenceDirectory)
$allowed='C:\Users\avidzba\.codex\visualizations\2026\10\12\01a12ece-8192-7141-89b3-f83f1176184c\TorsionPostprocess_2026-11-17\'
if(-not $directory.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $directory)){throw 'Fresh task evidence required'}
$mem=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory;$os=Get-CimInstance Win32_OperatingSystem
if(($mem.CommitLimit-$mem.CommittedBytes)/1GB -lt 7 -or $os.FreePhysicalMemory*1KB/1GB -lt 4){throw 'One small native test needs 1 GiB own forecast and 6 GiB reserve'}
New-Item -ItemType Directory -Path $directory|Out-Null
Add-Type -TypeDefinition @'
using System;using System.Runtime.InteropServices;
public static class TestIntegrationOwn08{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}
'@
$helper=@'
Option Explicit
' ==================== ДЛЯ ТЕСТОВ ====================
' Читает полный файловый снимок без изменения НДС и отдаёт проверяемый текст.
Public Function TestExportNM() As String
 Dim store As New CCalculationSnapshotStore, snapshot As CCalculationSnapshot
 Set snapshot = store.LoadFile(ThisWorkbook)
 Dim key As Variant, values As Variant, r As Long, c As Long, text As String, value As Variant, token As String
 For Each key In snapshot.Keys
  If Left$(CStr(key), 12) <> "TorsionPost." Then
   values = snapshot.Table(CStr(key))
   text = text & CStr(key) & "|" & CStr(UBound(values, 1)) & "|" & CStr(UBound(values, 2)) & vbCrLf
   For r = 1 To UBound(values, 1)
    For c = 1 To UBound(values, 2)
     value = values(r, c): token = CStr(value)
     If IsNumeric(value) Then
      If CDbl(value) = snapshot.RunID Then token = "[RunID]"
     End If
     If CStr(key) = "Summary" And r = 4 And c = 1 Then token = "[Elapsed]"
     text = text & CStr(r) & ":" & CStr(c) & "=" & token & vbTab
    Next c
    text = text & vbCrLf
   Next r
  End If
 Next key
 TestExportNM = text
End Function
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает сохранённые самостоятельные строки T (кручение) после чтения файла.
Public Function TestReadTorsion() As Variant
 Dim store As New CCalculationSnapshotStore, snapshot As CCalculationSnapshot
 Set snapshot = store.LoadFile(ThisWorkbook)
 TestReadTorsion = snapshot.Table("TorsionPost.Data")
End Function
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает все дополнительные матрицы после закрытия и повторного открытия файла.
Public Function TestTorsionKeys() As String
 Dim store As New CCalculationSnapshotStore, snapshot As CCalculationSnapshot, key As Variant, text As String
 Set snapshot = store.LoadFile(ThisWorkbook)
 For Each key In Array("TorsionPost.Headers", "TorsionPost.Data", "TorsionPost.Bars", "TorsionPost.MNT")
  If Not snapshot.HasTable(CStr(key)) Then Err.Raise 5, , "Missing " & CStr(key)
  text = text & CStr(key) & ";"
 Next key
 TestTorsionKeys = text
End Function
'@
$receipt=[ordered]@{Status='RUNNING';Own=$null;Checks=@();Messages=@();Baseline=$null;Candidate=$CandidatePath}
$app=$null;$book=$null;$owned=0;$ticks=0;$failure=$null
function Test-Check([string]$Name,[bool]$Passed){$receipt.Checks+=,[ordered]@{Name=$Name;Passed=$Passed};if(-not $Passed){throw ('FAIL '+$Name)}}
function Test-Set([string]$Key,$Value){$t=$book.Names.Item('rngSystemSettings').RefersToRange;for($i=2;$i -le $t.Rows.Count;$i++){if([string]$t.Cells.Item($i,1).Value2 -ceq $Key){if($Value -is [string]){$t.Cells.Item($i,2).Value2=[string]$Value}else{$t.Cells.Item($i,2).Value2=[double]$Value};return}};throw ('Missing setting '+$Key)}
function Test-TSet([string]$Key,$Value){$t=$book.Names.Item('rngTorsionSettings').RefersToRange;for($i=2;$i -le $t.Rows.Count;$i++){if([string]$t.Cells.Item($i,1).Value2 -ceq $Key){if($Value -is [string]){$t.Cells.Item($i,2).Value2=[string]$Value}else{$t.Cells.Item($i,2).Value2=[double]$Value};return}};throw ('Missing torsion setting '+$Key)}
function Test-Run([string]$Name){$app.Calculate();$msg=[string]$app.Run("'$($book.Name)'!modWorkbookCalculation.RunSectionCalculationForWorkbook",$book,$false);$receipt.Messages+=,[ordered]@{Name=$Name;Message=$msg};$msg|Set-Content -LiteralPath (Join-Path $directory ($Name+'_Primary.txt')) -Encoding UTF8;if($Name -eq 'Candidate_Excel'){return ''};return [string]$app.Run("'$($book.Name)'!TestTorsionIntegration08.TestExportNM")}
function Test-ReadT{return ,$app.Run("'$($book.Name)'!TestTorsionIntegration08.TestReadTorsion")}
try{
 $app=New-Object -ComObject Excel.Application;[uint32]$n=0;[void][TestIntegrationOwn08]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$n)
 $owned=[int]$n;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
 $receipt.Own=[ordered]@{PID=$owned;StartTicksUTC=$ticks;QuitRequested=$false;Forced=$false;Absent=$false}
 $receipt|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $directory 'Review.json') -Encoding UTF8
 $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=1
 foreach($spec in @(@{Name='Baseline';Path=$CandidatePath},@{Name='Candidate';Path=$CandidatePath})){
  $path=Join-Path $directory ($spec.Name+'.xlsm');Copy-Item -LiteralPath $spec.Path -Destination $path
  $book=$app.Workbooks.Open($path,0,$false)
  if($spec.Name -eq 'Baseline'){
   foreach($name in @('CSystemSettingsReader','modCalculationSnapshot','modSnapshotTables','modWorkbookCalculation')){
    $relative='src/Excel/'+$name+$(if($name -eq 'CSystemSettingsReader'){'.cls'}else{'.bas'})
    $code=(& git -c ('safe.directory='+$repo.Replace('\','/')) show ('1315d2fd28df06465c318d70af822c3a60ac4bda:'+$relative)) -join "`r`n"
    if($LASTEXITCODE -ne 0 -or $code.IndexOf('Option Explicit') -lt 0){throw 'Approved main source absent'}
    $component=$book.VBProject.VBComponents.Item($name);$component.CodeModule.DeleteLines(1,$component.CodeModule.CountOfLines);$component.CodeModule.AddFromString($code.Substring($code.IndexOf('Option Explicit')))
    [IO.File]::WriteAllText((Join-Path $directory ($name+'_Main.txt')),$code,[Text.UTF8Encoding]::new($true))
   }
  }
  $mainLoads=$book.Names.Item('rngLoadCombinations').RefersToRange
  $mainLoads.Cells.Item(2,1).Value2='1';$mainLoads.Cells.Item(2,2).Value2=10.0;$mainLoads.Cells.Item(2,3).Value2=1.0;$mainLoads.Cells.Item(2,4).Value2=0.0;$mainLoads.Cells.Item(2,5).Value2='PR1'
  $mainLoads.Cells.Item(3,1).Value2='2';$mainLoads.Cells.Item(3,2).Value2=20.0;$mainLoads.Cells.Item(3,3).Value2=2.0;$mainLoads.Cells.Item(3,4).Value2=0.0;$mainLoads.Cells.Item(3,5).Value2='PR1'
  $mainLoads.Cells.Item(4,1).Value2='3';$mainLoads.Cells.Item(4,2).Value2=0.0;$mainLoads.Cells.Item(4,3).Value2=0.0;$mainLoads.Cells.Item(4,4).Value2=0.0
  $module=$book.VBProject.VBComponents.Add(1);$module.Name='TestTorsionIntegration08';$module.CodeModule.AddFromString($helper)
  Test-Set 'Plot.Enabled' 'No';Test-Set 'Results.SnapshotStorage' 'File';Test-Set 'Results.SnapshotName' $spec.Name
  $form=[string]$app.Run("'$($book.Name)'!modWorkbookCalculation.FormSectionForWorkbook",$book)
  $form|Set-Content -LiteralPath (Join-Path $directory ($spec.Name+'_Form.txt')) -Encoding UTF8
  $text=Test-Run ($spec.Name+'_EmptyT');$text|Set-Content -LiteralPath (Join-Path $directory ($spec.Name+'_NM.txt')) -Encoding UTF8
  if($spec.Name -eq 'Baseline'){$baseline=$text;$receipt.Baseline=$path;$book.Save();$book.Close($false);$book=$null
   $app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app);$app=$null
   $old=Get-Process -Id $owned -ErrorAction SilentlyContinue;if($old -and $old.ProcessName -eq 'EXCEL' -and $old.StartTime.ToUniversalTime().Ticks -eq $ticks){Stop-Process -Id $owned}
   for($w=0;$w -lt 10 -and (Get-Process -Id $owned -ErrorAction SilentlyContinue);$w++){Start-Sleep -Milliseconds 200}
   $receipt['BaselineOwn']=[ordered]@{PID=$owned;UTCticks=$ticks;QuitRequested=$true;Absent=($null -eq (Get-Process -Id $owned -ErrorAction SilentlyContinue))}
   $memoryNow=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory;if(($memoryNow.CommitLimit-$memoryNow.CommittedBytes)/1GB -lt 7){throw 'Insufficient reserve for next serial test'}
   $app=New-Object -ComObject Excel.Application;[uint32]$n=0;[void][TestIntegrationOwn08]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$n)
   $owned=[int]$n;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
   $receipt.Own=[ordered]@{PID=$owned;StartTicksUTC=$ticks;QuitRequested=$false;Forced=$false;Absent=$false}
   $receipt|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $directory 'Review.json') -Encoding UTF8
   $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=1
   continue
  }
  Test-Check 'Empty T NM equals approved main' ($text -ceq $baseline)
  $loads=$book.Names.Item('rngTorsionCombinations').RefersToRange
  $loads.Cells.Item(2,2).Value2=1.0;$loads.Cells.Item(3,2).Value2=4.0;$loads.Cells.Item(4,2).Value2=20.0
  $active=Test-Run 'Candidate_ActiveT';Test-Check 'Active T NM equals approved main' ($active -ceq $baseline)
  $data=Test-ReadT
  Test-Check 'Three requested T rows saved' ($data.GetLength(0) -eq 3)
  Test-Check 'Missing saved state is addressed' ([string]$data[3,36] -ceq 'N/A' -and [string]$data[3,39] -match 'Torsion.StateType' -and [string]$data[3,39] -match 'Config!')
  Test-Check 'Independent algebraic sum' ([Math]::Abs(([double]$data[1,33]) - (([double]$data[1,30]) + ([double]$data[1,31]))) -lt 1e-8)
  $table=$book.Names.Item('rngTorsionPostSummary').RefersToRange
  Test-Check 'Formatted data survives' ([string]$table.Cells.Item(5,1).Value2 -eq [string]$data[1,1] -and [Math]::Abs(([double]$table.Cells.Item(5,33).Value2) - ([double]$data[1,33])) -lt 1e-8)
  $keys=[string]$app.Run("'$($book.Name)'!TestTorsionIntegration08.TestTorsionKeys");Test-Check 'Four optional matrices roundtrip' ($keys.Split(';').Count -eq 5)
  $file=Join-Path $directory 'Candidate.ndm';if(-not(Test-Path -LiteralPath $file)){$file=Join-Path $directory 'Candidate.ndms'};$before=(Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash
  foreach($key in @('rngTorsionPostSummary','rngTorsionBarForces','rngTorsionMNTSummary')){$book.Names.Item($key).RefersToRange.ClearContents()}
  [void]$app.Run("'$($book.Name)'!modCalculationSnapshot.UpdateSectionResultsForWorkbook",$book)
  Test-Check 'File refresh restores T data' ([string]$table.Cells.Item(5,1).Value2 -eq [string]$data[1,1])
  Test-Check 'File refresh does not write snapshot' ((Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash -ceq $before)
  $loads.Cells.Item(2,2).Formula="='Расчет'!ZZ1";$book.Worksheets.Item('Расчет').Range('ZZ1').Value2='bad torque'
  $invalid=Test-Run 'Candidate_InvalidT';Test-Check 'Invalid T NM unchanged' ($invalid -ceq $baseline);$data=Test-ReadT
  Test-Check 'Invalid T clear status and source chain' ([string]$data[1,38] -ceq 'InputErr' -and [string]$data[1,39] -match 'ZZ1' -and [string]$data[1,39] -match 'числом')
  $loads.Cells.Item(2,2).Value2=1.0;$loads.Cells.Item(3,1).Value2=[string]$loads.Cells.Item(2,1).Value2
  [void](Test-Run 'Candidate_DuplicateT');$data=Test-ReadT;Test-Check 'Duplicate active IDs both rejected' ([string]$data[1,38] -ceq 'InputErr' -and [string]$data[2,38] -ceq 'InputErr')
  $loads.Cells.Item(3,2).Value2=0.0;[void](Test-Run 'Candidate_ZeroDuplicate');$data=Test-ReadT
  Test-Check 'Zero T fully disables duplicate row' ($data.GetLength(0) -eq 2 -and [string]$data[1,38] -cne 'InputErr')
  Test-TSet 'Torsion.ConcreteMethod' 'None';$ts=$book.Names.Item('rngTorsionSettings').RefersToRange
  foreach($key in @('Torsion.Rb','Torsion.ConcreteEfficiency','Torsion.WallThickness','Torsion.GorskyCoefficient')){for($r=2;$r -le $ts.Rows.Count;$r++){if([string]$ts.Cells.Item($r,1).Value2 -ceq $key){$ts.Cells.Item($r,2).Formula='=1/0'}}}
  [void](Test-Run 'Candidate_InactiveConcrete');$data=Test-ReadT;Test-Check 'Inactive concrete errors not read' ([string]$data[1,28] -ceq 'N/A' -and [string]$data[1,38] -cne 'InputErr')
  Test-TSet 'Torsion.HoopSpacing' 0.0;[void](Test-Run 'Candidate_InvalidSpacing');$data=Test-ReadT
  Test-Check 'Invalid setting is addressed' ([string]$data[1,38] -ceq 'InputErr' -and [string]$data[1,39] -match 'Torsion.HoopSpacing' -and [string]$data[1,39] -match 'Config!')
  Test-TSet 'Torsion.HoopSpacing' 100.0;Test-Set 'Results.SnapshotStorage' 'Excel'
  [void]$app.Run("'$($book.Name)'!modWorkbookCalculation.FormSectionForWorkbook",$book)
  [void](Test-Run 'Candidate_Excel')
  Test-Check 'Excel output populated' ([string]$table.Cells.Item(5,1).Value2 -eq [string]$loads.Cells.Item(2,1).Value2 -and [string]$table.Cells.Item(5,28).Value2 -ceq 'N/A')
  $book.Save();$book.Close($false);$book=$app.Workbooks.Open($path,0,$false)
  $table=$book.Names.Item('rngTorsionPostSummary').RefersToRange
  Test-Check 'Saved reopened Excel output intact' ([string]$table.Cells.Item(5,1).Value2 -eq '1' -and [string]$table.Cells.Item(5,28).Value2 -ceq 'N/A')
  [void]$app.Run("'$($book.Name)'!modWorkbookCalculation.FormSectionForWorkbook",$book)
  Test-Check 'Form clears only stale torsion results' ([string]$table.Cells.Item(5,1).Value2 -eq '')
 }
 $receipt.Status='DIRECTED_INTEGRATION_PASS_NOT_FULL_SUITE'
}catch{$failure=$_;$receipt.Status='FAILED';$receipt['Error']=[string]$_}
finally{
 if($book){try{$book.Close($false)}catch{};$book=$null}
 if($app){try{$receipt.Own.QuitRequested=$true;$app.Quit()}catch{};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app);$app=$null}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
 if($owned){$p=Get-Process -Id $owned -ErrorAction SilentlyContinue;if($p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.ToUniversalTime().Ticks -eq $ticks){$receipt.Own.Forced=$true;Stop-Process -Id $owned};for($i=0;$i -lt 10 -and (Get-Process -Id $owned -ErrorAction SilentlyContinue);$i++){Start-Sleep -Milliseconds 200};$receipt.Own.Absent=($null -eq (Get-Process -Id $owned -ErrorAction SilentlyContinue))}
 $receipt|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $directory 'Review.json') -Encoding UTF8
}
if($failure){throw $failure}
[ordered]@{Status=$receipt.Status;Checks=$receipt.Checks.Count;Own=$receipt.Own}|ConvertTo-Json -Depth 5
