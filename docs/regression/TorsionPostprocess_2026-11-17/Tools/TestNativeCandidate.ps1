param([Parameter(Mandatory=$true)][string]$CandidatePath,[Parameter(Mandatory=$true)][string]$EvidenceDirectory,[switch]$WaitForListUI)
$ErrorActionPreference='Stop'
Import-Module C:/Windows/System32/WindowsPowerShell/v1.0/Modules/Microsoft.PowerShell.Utility/Microsoft.PowerShell.Utility.psd1 -Global
$source=(Resolve-Path -LiteralPath $CandidatePath).Path
$directory=[IO.Path]::GetFullPath($EvidenceDirectory)
$allowed='C:\Users\avidzba\.codex\visualizations\2026\10\12\01a12ece-8192-7141-89b3-f83f1176184c\TorsionPostprocess_2026-11-17\'
if(-not $directory.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $directory)){throw 'Fresh task evidence directory required'}
New-Item -ItemType Directory -Path $directory|Out-Null
$os=Get-CimInstance Win32_OperatingSystem;$memory=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
if(($memory.CommitLimit-$memory.CommittedBytes)/1GB -lt 7 -or $os.FreePhysicalMemory*1KB/1GB -lt 4){throw 'Directed smoke requires 1 GiB own peak plus 6 GiB reserve; this is not a full suite'}
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class TestTorsionPostNativeIdentity {
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);
 public static string Normalize(string text) {
  int start=text.IndexOf("Option Explicit",StringComparison.Ordinal);
  if(start<0)throw new Exception("Option Explicit missing");
  var result=new System.Text.StringBuilder();
  foreach(string raw in text.Substring(start).Replace("\r\n","\n").Split('\n')) {
   string line=raw.TrimEnd(); bool quoted=false,comment=false;
   for(int i=0;i<line.Length;i++) {
    char c=line[i];
    if(comment){result.Append(c);continue;}
    if(c=='"') {result.Append(c);if(quoted && i+1<line.Length && line[i+1]=='"'){result.Append(line[++i]);}else quoted=!quoted;continue;}
    if(c=='\'' && !quoted){while(result.Length>0 && result[result.Length-1]==' ')result.Length--;comment=true;result.Append(c);continue;}
    if(!quoted && (char.IsDigit(c) || (c=='.' && i+1<line.Length && char.IsDigit(line[i+1]))) && (i==0 || (!char.IsLetterOrDigit(line[i-1]) && line[i-1]!='_'))) {
     var number=System.Text.RegularExpressions.Regex.Match(line.Substring(i), @"^(?:\d*\.\d+|\d+)(?:[eE][+-]?\d+)?");
     if(number.Success){
      double parsed=double.Parse(number.Value,System.Globalization.CultureInfo.InvariantCulture);
      bool fractional=number.Value.IndexOfAny(new char[]{'.','e','E'})>=0;
      int next=i+number.Length;bool explicitDouble=next<line.Length && line[next]=='#';
      bool otherSuffix=next<line.Length && "!@%&".IndexOf(line[next])>=0;
      result.Append(parsed.ToString("G15",System.Globalization.CultureInfo.InvariantCulture));
      if(explicitDouble || (fractional && !otherSuffix))result.Append('#');
      i+=number.Length-1;if(explicitDouble)i++;continue;
     }
    }
    result.Append(quoted?c:char.ToLowerInvariant(c));
   }
   result.Append('\n');
  }
  return result.ToString().TrimEnd();
 }

}
'@
$bookPath=Join-Path $directory 'RC_Section_NDM.xlsm'
$sourceHash=(Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
Copy-Item -LiteralPath $source -Destination $bookPath
if((Get-FileHash -LiteralPath $bookPath -Algorithm SHA256).Hash -ne $sourceHash){throw 'Candidate copy changed'}
$app=$null;$book=$null;$ownedId=0;$ownedTicks=0;$quitRequested=$false;$bookClosed=$false;$failure=$null
$receipt=[ordered]@{Status='RUNNING';Source=$source;SourceSHA256=$sourceHash;StartedUTC=[DateTime]::UtcNow.ToString('o');Math=$null;Lists=@();SourceModules=@();Own=$null}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Сверяет именно тот Excel, который создан этим тестом.
function Test-OwnExcel {
 $p=Get-Process -Id $ownedId -ErrorAction SilentlyContinue
 return $p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.ToUniversalTime().Ticks -eq $ownedTicks
}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Отделяет текст исходного компонента от импортных атрибутов класса.
function Test-CodeBody([string]$Text) {
 return [TestTorsionPostNativeIdentity]::Normalize($Text)
}
try {
 $app=New-Object -ComObject Excel.Application
 [uint32]$number=0;[void][TestTorsionPostNativeIdentity]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$number)
 $ownedId=[int]$number;$ownedTicks=(Get-Process -Id $ownedId).StartTime.ToUniversalTime().Ticks
 $receipt.Own=[ordered]@{PID=$ownedId;StartTicksUTC=$ownedTicks;QuitRequested=$false;BookClosed=$false;Absent=$false;Forced=$false}
 $receipt|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $directory 'Review.json') -Encoding UTF8
 $app.Caption='NDM Torsion Native Acceptance';$app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=1
 $book=$app.Workbooks.Open($bookPath,0,$false)
 $math=[string]$app.Run("'$($book.Name)'!modTestTorsionPost.RunTorsionPostTests")
 $math|Set-Content -LiteralPath (Join-Path $directory 'MathPrimary.txt') -Encoding UTF8
 if($math -notmatch 'Passed=([0-9]+); Failed=0' -or $math -match 'RUNTIME:|FAIL '){throw ('Analytical tests failed: '+$math)}
 $receipt.Math=[ordered]@{Passed=[int]$Matches[1];Primary='MathPrimary.txt';EntirePrimaryRead=$true}
 $output=[string]$app.Run("'$($book.Name)'!modTestTorsionPost.RunTorsionPostOutputTests")
 $output|Set-Content -LiteralPath (Join-Path $directory 'OutputPrimary.txt') -Encoding UTF8
 if($output -notmatch 'Passed=([0-9]+); Failed=0' -or $output -match 'RUNTIME:|FAIL '){throw ('Native output regression failed: '+$output)}
 $receipt['Output']=[ordered]@{Passed=[int]$Matches[1];Primary='OutputPrimary.txt';EntirePrimaryRead=$true}
 $guard=[string]$app.Run("'$($book.Name)'!modTestTorsionPost.RunTorsionPostGuardTests")
 $guard|Set-Content -LiteralPath (Join-Path $directory 'GuardPrimary.txt') -Encoding UTF8
 if($guard -notmatch 'Passed=([0-9]+); Failed=0' -or $guard -match 'RUNTIME:|FAIL '){throw ('Native boundary guard failed: '+$guard)}
 $receipt['Guards']=[ordered]@{Passed=[int]$Matches[1];Primary='GuardPrimary.txt';EntirePrimaryRead=$true}
 foreach($spec in @(
   @{Key='Torsion.StateType';Values=@('StrengthState','CapacityState','PreCrackState','PostCrackState','CrackedState')},
   @{Key='Torsion.ConcreteMethod';Values=@('Strut','GorskyRectangle','None')}
 )) {
  $table=$book.Names.Item('rngTorsionSettings').RefersToRange;$cell=$null
  for($r=2;$r -le $table.Rows.Count;$r++){if([string]$table.Cells.Item($r,1).Value2 -ceq $spec.Key){$cell=$table.Cells.Item($r,2);break}}
  if($null -eq $cell){throw ('Setting missing: '+$spec.Key)}
  $original=$cell.Formula;$selections=@()
  foreach($choice in $spec.Values){$cell.Value2=[string]$choice;$accepted=[bool]$cell.Validation.Value;if(-not $accepted){throw ('Native list rejected '+$spec.Key+'='+$choice)};$selections+=,$choice}
  $cell.Value2='__INVALID_TORSION_OPTION__';if([bool]$cell.Validation.Value){throw ('Native list accepts invalid option '+$spec.Key)}
  $cell.Value2=($spec.Values -join [string]$app.International(5));if([bool]$cell.Validation.Value){throw ('Native list accepts joined text '+$spec.Key)}
  $cell.Formula=$original
  $receipt.Lists+=,[ordered]@{Key=$spec.Key;Address=$cell.Address();Formula1=[string]$cell.Validation.Formula1;NativeSeparator=[string]$app.International(5);Accepted=$selections;InvalidRejected=$true;JoinedRejected=$true;UIItemsObserved=$false}
 }
 $book.Save();$book.Close($false);$book=$null
 $book=$app.Workbooks.Open($bookPath,0,$true)
 foreach($record in $receipt.Lists){$table=$book.Names.Item('rngTorsionSettings').RefersToRange;for($r=2;$r -le $table.Rows.Count;$r++){if([string]$table.Cells.Item($r,1).Value2 -ceq $record.Key){$cell=$table.Cells.Item($r,2);break}};if([string]$cell.Validation.Formula1 -cne $record.Formula1){throw ('Saved list changed '+$record.Key)};foreach($choice in $record.Accepted){$cell.Value2=[string]$choice;if(-not [bool]$cell.Validation.Value){throw ('Reopened list rejects '+$choice)}};$record['SavedReopenedAccepted']=$true}
 $repo='C:\Users\avidzba\Desktop\программы\В_РАБОТЕ!\RC_Section_NDM'
 $sources=@(Get-ChildItem -LiteralPath (Join-Path $repo 'src') -Recurse -File|Where-Object {$_.Extension -in '.bas','.cls'};Get-ChildItem -LiteralPath (Join-Path $repo 'tests') -Recurse -File|Where-Object {$_.Extension -in '.bas','.cls'})
 $sourceDifferences=@()
 foreach($file in $sources){$name=$file.BaseName;try{$component=$book.VBProject.VBComponents.Item($name)}catch{throw ('Native module absent '+$name)};$native=[string]$component.CodeModule.Lines(1,$component.CodeModule.CountOfLines);$expected=[IO.File]::ReadAllText($file.FullName,[Text.Encoding]::UTF8);if((Test-CodeBody $native) -cne (Test-CodeBody $expected)){$native|Set-Content -LiteralPath (Join-Path $directory ($name+'_Native.txt')) -Encoding UTF8;$expected|Set-Content -LiteralPath (Join-Path $directory ($name+'_Expected.txt')) -Encoding UTF8;$sourceDifferences+=,$name};$receipt.SourceModules+=,$name}
 $receipt['SourceDifferences']=$sourceDifferences
 $export=New-Object Text.StringBuilder
 [void]$export.AppendLine('VBA PROJECT EXPORT');[void]$export.AppendLine('Workbook: '+$book.Name);[void]$export.AppendLine('Date: '+[DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss'))
 foreach($component in $book.VBProject.VBComponents){[void]$export.AppendLine('='*100);[void]$export.AppendLine('COMPONENT: '+$component.Name);[void]$export.AppendLine('TYPE: '+$component.Type);[void]$export.AppendLine('LINES: '+$component.CodeModule.CountOfLines);[void]$export.AppendLine('='*100);if($component.CodeModule.CountOfLines){[void]$export.AppendLine($component.CodeModule.Lines(1,$component.CodeModule.CountOfLines))};[void]$export.AppendLine()}
 [IO.File]::WriteAllText((Join-Path $directory 'VBA_All_Code.txt'),$export.ToString(),[Text.UTF8Encoding]::new($true))
 if($WaitForListUI){
  $app.Visible=$true;$app.ScreenUpdating=$true;$app.WindowState=-4137;$book.Windows.Item(1).Visible=$true;$book.Windows.Item(1).WindowState=-4137;$book.Worksheets.Item('Config').Activate()
  foreach($ui in @(@{Name='State';Address='AG91'},@{Name='Method';Address='AG99'})){
   $book.Worksheets.Item('Config').Range($ui.Address).Select();$app.ActiveWindow.ScrollRow=89;$app.ActiveWindow.ScrollColumn=32
   $receipt['UIActiveCell']=[string]$app.ActiveCell.Address();$receipt['UIWorksheet']=[string]$app.ActiveSheet.Name;$receipt['UIWindowState']=[int]$app.WindowState;$receipt['UIScreenUpdating']=[bool]$app.ScreenUpdating;$receipt['UIStage']=$ui.Name;$receipt|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $directory 'Review.json') -Encoding UTF8
   $until=[DateTime]::UtcNow.AddMinutes(3);$signal=Join-Path $directory ('UI_Continue_'+$ui.Name)
   while(-not(Test-Path -LiteralPath $signal)){if([DateTime]::UtcNow -gt $until){throw ('UI observation timeout '+$ui.Name)};Start-Sleep -Milliseconds 500}
  }
 }
 $receipt.Status='MATH_SOURCE_NATIVE_LIST_VALUES_PASS_UI_PENDING_WORKFLOW_PENDING'
 if($sourceDifferences.Count){$receipt.Status='NATIVE_MATH_GUARD_LIST_VALUES_DONE_SOURCE_REVIEW_REQUIRED'}
} catch {$failure=$_;$receipt.Status='FAILED';$receipt['Error']=[string]$_}
finally {
 if($null -ne $book){try{$book.Close($false);$bookClosed=$true}catch{$receipt['CloseError']=[string]$_};$book=$null}
 if($null -ne $app -and (Test-OwnExcel)){try{$quitRequested=$true;$app.Quit()}catch{$receipt['QuitError']=[string]$_}}
 if($null -ne $app){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app);$app=$null}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
 if($ownedId){$receipt.Own.QuitRequested=$quitRequested;$receipt.Own.BookClosed=$bookClosed;for($wait=0;$wait -lt 10 -and (Test-OwnExcel);$wait++){Start-Sleep -Milliseconds 500};if(Test-OwnExcel){$receipt.Own.Forced=$true;Stop-Process -Id $ownedId -ErrorAction Stop};for($wait=0;$wait -lt 10 -and (Test-OwnExcel);$wait++){Start-Sleep -Milliseconds 200};$receipt.Own.Absent=(-not(Test-OwnExcel))}
 $receipt['FinishedUTC']=[DateTime]::UtcNow.ToString('o');$receipt|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $directory 'Review.json') -Encoding UTF8
}
if($failure){throw $failure}
[ordered]@{Status=$receipt.Status;Math=$receipt.Math;SourceModules=$receipt.SourceModules.Count;Own=$receipt.Own;Receipt=(Join-Path $directory 'Review.json')}|ConvertTo-Json -Depth 5
