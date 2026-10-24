# Проверяет реальное Save/Close/Open, Save As и обе кнопки диалога большого снимка.
# Все книги и файлы принадлежат тесту; исходная книга открывается только через копию.
param([Parameter(Mandatory=$true)][string]$WorkbookPath,
      [Parameter(Mandatory=$true)][string]$EvidenceDirectory)
$ErrorActionPreference='Stop'
$source=(Resolve-Path -LiteralPath $WorkbookPath).Path
$hash=(Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$directory=[IO.Path]::GetFullPath((Join-Path $root $EvidenceDirectory))
$allowed=[IO.Path]::GetFullPath((Join-Path $root 'docs/regression'))+[IO.Path]::DirectorySeparatorChar
if(-not $directory.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $directory)){throw 'A fresh private evidence directory is required.'}
New-Item -ItemType Directory -Path $directory|Out-Null
Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
public static class SnapshotLifecycleTestWindow {
 public delegate bool EnumProc(IntPtr h,IntPtr p);
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);
 [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc p,IntPtr l);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr h,StringBuilder t,int n);
 [DllImport("user32.dll")] static extern IntPtr GetDlgItem(IntPtr h,int id);
 [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr h,uint m,IntPtr w,IntPtr l);
 public static Task<bool> Answer(uint owned,int button) {
  return Task.Run(()=>{
   DateTime end=DateTime.UtcNow.AddSeconds(25);
   while(DateTime.UtcNow<end) {
    bool found=false;
    EnumWindows((h,p)=>{uint id;GetWindowThreadProcessId(h,out id);if(id!=owned)return true;
     var title=new StringBuilder(256);GetWindowText(h,title,256);
     if(title.ToString()!="Хранение большого снимка")return true;
     IntPtr control=GetDlgItem(h,button);if(control==IntPtr.Zero)return true;
     SendMessage(control,0xF5,IntPtr.Zero,IntPtr.Zero);found=true;return false;
    },IntPtr.Zero);
    if(found)return true;
    Thread.Sleep(50);
   }
   return false;
  });
 }
}
'@
$app=$null;$book=$null;$owned=0;$ticks=0;$records=@()

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Ограничивает закрытие точным процессом, созданным этим тестом.
function Test-OwnedSnapshotExcel {
 $process=Get-Process -Id $owned -ErrorAction SilentlyContinue
 return $process -and $process.ProcessName -eq 'EXCEL' -and $process.StartTime.ToUniversalTime().Ticks -eq $ticks
}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Хеширует возвращённые точные биты всех таблиц, схемы и данных экспорта.
function Test-SnapshotFingerprint([string]$EntryPoint) {
 $text=[string]$app.Run("'$($book.Name)'!modTestCalculationSnapshot.$EntryPoint")
 $digest=[Security.Cryptography.SHA256]::Create()
 try{return [BitConverter]::ToString($digest.ComputeHash([Text.Encoding]::UTF8.GetBytes($text))).Replace('-','')}
 finally{$digest.Dispose()}
}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Возвращает существующую ячейку значения по ключу; новых настроек не создаёт.
function Test-SnapshotSetting([string]$Key) {
 $table=$book.Names.Item('rngSystemSettings').RefersToRange;$values=$table.Value2
 for($i=1;$i -le $values.GetLength(0);$i++){if([string]$values[$i,1] -ceq $Key){return $table.Cells.Item($i,2)}}
 throw ('Setting missing: '+$Key)
}

try {
 $app=New-Object -ComObject Excel.Application
 [uint32]$id=0;[void][SnapshotLifecycleTestWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id)
 $owned=[int]$id;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
 [ordered]@{PID=$owned;StartTicks=$ticks;SourceSHA256=$hash}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $directory 'Ownership.json') -Encoding UTF8
 $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$true;$app.AutomationSecurity=1
 foreach($mode in @('Excel','File')) {
  $folder=Join-Path $directory $mode;New-Item -ItemType Directory -Path $folder|Out-Null
  $path=Join-Path $folder 'RC_Section_NDM.xlsm';Copy-Item -LiteralPath $source -Destination $path
  $book=$app.Workbooks.Open($path,0,$false)
  $initial=[string]$app.Run("'$($book.Name)'!modTestCalculationSnapshot.TestPrepareSnapshotReopen",$mode)
  $digest=[Security.Cryptography.SHA256]::Create()
  try{$expected=[BitConverter]::ToString($digest.ComputeHash([Text.Encoding]::UTF8.GetBytes($initial))).Replace('-','')}finally{$digest.Dispose()}
  $book.Save();$book.Close($false);$book=$null
  $book=$app.Workbooks.Open($path,0,$false)
  $actual=Test-SnapshotFingerprint 'TestRestoreSnapshotAfterReopen'
  if($actual -cne $expected){throw ('Reopen changed data: '+$mode)}
  $records+=,[ordered]@{Mode=$mode;Stage='Save/Close/Open/Restore';ExactSHA256=$actual;SolverCalls=0}
  $moved=Join-Path $folder 'SaveAs';New-Item -ItemType Directory -Path $moved|Out-Null
  $movedPath=Join-Path $moved 'RC_Section_NDM.xlsm'
  $fileHash=$null
  if($mode -eq 'File'){$file=Get-ChildItem -LiteralPath $folder -Filter '*.ndms';if(@($file).Count -ne 1){throw 'Expected one snapshot file.'};$fileHash=(Get-FileHash -LiteralPath $file.FullName).Hash}
  $book.SaveAs($movedPath,52);$book.Close($false);$book=$null
  if($mode -eq 'File'){
   $movedFile=Get-ChildItem -LiteralPath $moved -Filter '*.ndms'
   if(@($movedFile).Count -ne 1 -or (Get-FileHash -LiteralPath $movedFile.FullName).Hash -cne $fileHash){throw 'Save As did not preserve the exact single snapshot.'}
   if((Get-FileHash -LiteralPath $file.FullName).Hash -cne $fileHash){throw 'Save As modified the source snapshot.'}
  }
  $book=$app.Workbooks.Open($movedPath,0,$false)
  $actual=Test-SnapshotFingerprint 'TestRestoreSnapshotAfterReopen'
  if($actual -cne $expected){throw ('Save As changed data: '+$mode)}
  $records+=,[ordered]@{Mode=$mode;Stage='SaveAs/Close/Open/Restore';ExactSHA256=$actual;ExternalSHA256=$fileHash;SolverCalls=0}
  if($mode -eq 'File'){
   # Удаляется только проверенный собственный файл из первой тестовой папки.
   if(-not $file.FullName.StartsWith($directory+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw 'Original test snapshot is outside the evidence directory.'}
   Remove-Item -LiteralPath $file.FullName
   $secondFolder=Join-Path $moved 'SecondSaveAs';New-Item -ItemType Directory -Path $secondFolder|Out-Null
   $secondPath=Join-Path $secondFolder 'RC_Section_NDM.xlsm'
   $book.SaveAs($secondPath,52);$book.Close($false);$book=$null
   $secondFile=Get-ChildItem -LiteralPath $secondFolder -Filter '*.ndms'
   if(@($secondFile).Count -ne 1 -or (Get-FileHash -LiteralPath $secondFile.FullName).Hash -cne $fileHash){throw 'Repeated Save As lost the snapshot after removing its original copy.'}
   $book=$app.Workbooks.Open($secondPath,0,$false)
   $actual=Test-SnapshotFingerprint 'TestRestoreSnapshotAfterReopen'
   if($actual -cne $expected){throw 'Repeated Save As changed the full snapshot.'}
   $records+=,[ordered]@{Mode=$mode;Stage='SecondSaveAs/OriginalRemoved/Reopen/Restore';ExactSHA256=$actual;ExternalSHA256=$fileHash;SolverCalls=0}
  }
  $book.Close($false);$book=$null
 }
 $promptPath=Join-Path $directory 'Prompt.xlsm';Copy-Item -LiteralPath $source -Destination $promptPath
 $book=$app.Workbooks.Open($promptPath,0,$false)
 foreach($case in @(@{Rows=60000;Button=0;Expected='Excel'},@{Rows=60001;Button=7;Expected='Excel'},@{Rows=60001;Button=6;Expected='File'})){
  (Test-SnapshotSetting 'Results.SnapshotStorage').Value2='Excel'
  $answer=$null;if($case.Button){$answer=[SnapshotLifecycleTestWindow]::Answer([uint32]$owned,[int]$case.Button)}
  [void]$app.Run("'$($book.Name)'!modCalculationSnapshot.OfferExternalCalculationSnapshot",$book,[double]$case.Rows,$true)
  if($answer -and -not $answer.GetAwaiter().GetResult()){throw 'The expected own dialog was not answered.'}
  $actual=[string](Test-SnapshotSetting 'Results.SnapshotStorage').Value2
  if($actual -cne $case.Expected){throw ('Incorrect prompt choice: '+$actual)}
  $records+=,[ordered]@{Stage='LargeSnapshotPrompt';EstimatedRows=$case.Rows;Button=$case.Button;Storage=$actual;Continued=$true}
 }
 if((Get-FileHash -LiteralPath $source).Hash -cne $hash){throw 'Source workbook changed.'}
 [ordered]@{Status='PASS';SourceSHA256=$hash;Checks=$records}|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $directory 'Lifecycle.json') -Encoding UTF8
 Write-Output 'CALCULATION_SNAPSHOT_LIFECYCLE_PASS'
} finally {
 if(Test-OwnedSnapshotExcel){try{if($book){$book.Close($false)}}catch{};try{if($app){$app.Quit()}}catch{}}
 if($app){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
 if(Test-OwnedSnapshotExcel){Stop-Process -Id $owned -Force -ErrorAction SilentlyContinue}
}
