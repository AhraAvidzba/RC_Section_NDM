# Адресная замена текста сообщения в частной копии; пользовательская книга не открывается.
$ErrorActionPreference='Stop'
$root=(Get-Location).Path
$dir=Join-Path $root 'docs/regression/CompactFilterMessage_2026-11-27'
$memory=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
if($memory.AvailableMBytes -lt 4096 -or ($memory.CommitLimit-$memory.CommittedBytes) -lt 7GB){throw 'Insufficient RAM/commit before private Excel'}
$memory|Select-Object AvailableMBytes,CommitLimit,CommittedBytes|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $dir 'Memory.json') -Encoding UTF8
$native=Join-Path $dir 'Native'
New-Item -ItemType Directory -Path $native -Force|Out-Null
$path=Join-Path $native 'RC_Section_NDM.xlsm'
Copy-Item -LiteralPath (Join-Path $dir 'Source/RC_Section_NDM.xlsm') -Destination $path
Add-Type 'using System;using System.Runtime.InteropServices;public static class CompactMessageWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr w,out uint id);}'
$existing=@(Get-Process EXCEL -ErrorAction SilentlyContinue|ForEach-Object Id)
$excel=$null;$book=$null;$owned=0;$ticks=0
try{
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3
    [uint32]$id=0;[void][CompactMessageWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $owned=[int]$id;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
    if($owned -in $existing){throw 'Private Excel must not reuse a user process'}
    Write-Output "OWNED_EXCEL: $owned; ticks=$ticks"
    $book=$excel.Workbooks.Open($path,0,$false)
    if($book.ReadOnly){throw 'Private candidate is read-only'}
    $module=$book.VBProject.VBComponents.Item('modWorkbookCalculation').CodeModule
    $before=$module.Lines(1,$module.CountOfLines)
    $source=[IO.File]::ReadAllText((Join-Path $root 'src/Excel/modWorkbookCalculation.bas'),[Text.Encoding]::UTF8)
    $pattern='(?ms)^Private Function FilteredCombinationMessage\b.*?^End Function\r?$'
    $old=[regex]::Matches($before,$pattern);$new=[regex]::Matches($source,$pattern)
    if($old.Count -ne 1 -or $new.Count -ne 1){throw 'Expected exactly one message formatter'}
    $after=$before.Replace($old[0].Value,$new[0].Value)
    $module.DeleteLines(1,$module.CountOfLines);$module.AddFromString($after)
    $excel.VBE.ActiveVBProject=$book.VBProject
    $compile=$excel.VBE.CommandBars.FindControl(1,578)
    if($compile.Enabled){$compile.Execute()}
    if($compile.Enabled){throw 'VBA compilation failed'}
    if($book.VBProject.VBComponents.Count -ne 209){throw 'Component count changed'}
    $text=[Text.StringBuilder]::new()
    [void]$text.AppendLine('VBA PROJECT EXPORT');[void]$text.AppendLine('Workbook: RC_Section_NDM.xlsm')
    foreach($component in $book.VBProject.VBComponents){
        [void]$text.AppendLine('='*100);[void]$text.AppendLine('COMPONENT: '+$component.Name)
        [void]$text.AppendLine('TYPE: '+$component.Type);$lines=$component.CodeModule.CountOfLines
        [void]$text.AppendLine('LINES: '+$lines);[void]$text.AppendLine('='*100)
        if($lines -gt 0){[void]$text.AppendLine($component.CodeModule.Lines(1,$lines))}
        [void]$text.AppendLine()
    }
    [IO.File]::WriteAllText((Join-Path $dir 'VBA_All_Code.txt'),$text.ToString(),[Text.UTF8Encoding]::new($true))
    $book.Save();$book.Close($false);$book=$null
    [ordered]@{Compiled=$true;Components=209;ChangedFormatter='FilteredCombinationMessage';UserWorkbookOpened=$false;PID=$owned;StartTicks=$ticks}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $dir 'Preparation.json') -Encoding UTF8
    Write-Output 'COMPACT_MESSAGE_COMPILED'
}finally{
    if($book){$book.Close($false)}
    if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($owned -and $owned -notin $existing){$p=Get-Process -Id $owned -ErrorAction SilentlyContinue;if($p -and $p.StartTime.ToUniversalTime().Ticks -eq $ticks){Stop-Process -Id $owned -Force}}
}
