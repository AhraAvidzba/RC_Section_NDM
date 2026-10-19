# Снимает расположение из сохранённой книги, не меняя её и пользовательские данные.
param([Parameter(Mandatory=$true)][string]$WorkbookPath,[string]$LayoutPath='tools/build_workbook/WorkbookLayout.json')
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'WorkbookLayout.ps1')
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$path=(Resolve-Path -LiteralPath (Join-Path $root $WorkbookPath)).Path
$output=[IO.Path]::GetFullPath((Join-Path $root $LayoutPath))
if(-not $output.StartsWith($root+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw 'Описание расположения должно оставаться в проекте.'}
$hash=(Get-FileHash -LiteralPath $path).Hash;$xml=Get-WorkbookLayoutXmlSheets $path
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class LayoutCaptureWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
$app=$null;$book=$null;$owned=0;$ticks=0
try{
 $app=New-Object -ComObject Excel.Application;$app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=3
 [uint32]$id=0;[void][LayoutCaptureWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id);$owned=[int]$id;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
 $book=$app.Workbooks.Open($path,0,$true)
 if(-not $book.ReadOnly){throw 'Книга должна открываться только для чтения.'}
 $layout=Get-WorkbookLayoutSnapshot $book $xml
 $layout | Add-Member Provenance ([ordered]@{Workbook=[IO.Path]::GetFileName($path);SHA256=$hash;Captured=(Get-Date).ToString('s')})
 $layout | ConvertTo-Json -Depth 15 | Set-Content -LiteralPath $output -Encoding UTF8
 Write-Output ('LAYOUT_CAPTURED: names='+$layout.Names.Count+'; sheets='+$layout.Sheets.Count+'; objects='+(($layout.Sheets | ForEach-Object {$_.Objects.Count} | Measure-Object -Sum).Sum))
}finally{
 if($book){$book.Close($false)}
 if($app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
 if($owned){$p=Get-Process -Id $owned -ErrorAction SilentlyContinue;if($p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.ToUniversalTime().Ticks -eq $ticks){Stop-Process -Id $owned -Force}}
}
if((Get-FileHash -LiteralPath $path).Hash -cne $hash){throw 'Исходная книга изменилась во время чтения.'}
