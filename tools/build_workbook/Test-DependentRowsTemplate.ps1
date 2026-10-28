# Проверяет прямое создание новых форм Config, без перехода со старого формата.
param([string]$EvidenceDirectory='docs/regression/DependentRowOffsets_2026-10-17',[string]$Stage='DirectTemplate02')
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
$directory=(Resolve-Path -LiteralPath $EvidenceDirectory).Path
if($Stage -notmatch '^[A-Za-z][A-Za-z0-9_]*$'){throw 'Invalid local stage name.'}
$path=Join-Path $directory ($Stage+'.xlsx')
if(Test-Path -LiteralPath $path){throw 'Template evidence exists.'}
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class DirectRowsWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hWnd,out uint id);}'
$excel=$null;$book=$null;$ownedPid=0;$ownedTicks=0;$checks=@()
try{
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false
    [uint32]$identity=0;[void][DirectRowsWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$identity)
    $ownedPid=[int]$identity;$ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    Write-Output "OWNED_EXCEL: $ownedPid"
    $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1);$sheet.Name='Config'
    Add-UnitSettingsTable $book $sheet 3 1
    $top=30
    foreach($item in (Get-GeometrySettingsCatalog)){
        $rows=Add-GeometrySettingsTable $book $sheet $item $top 8
        $range=$book.Names.Item($item.RangeName).RefersToRange
        if($range.Columns.Count -ne 11 -or $range.Rows.Count -ne $rows){throw "Wrong range: $($item.RangeName)"}
        if($range.FormatConditions.Count -ne 0){throw 'Conditional formatting is not permitted.'}
        $blocks=@(Get-GeometryTableBlocks $range)
        if($blocks.Count -ne $item.Blocks.Count){throw 'Wrong block count.'}
        $extra=@($blocks | Where-Object Key -eq 'Extra')[0]
        if([string]$range.Cells.Item($extra.HeaderRow,4).Value2 -cne 'Смещение 2' -or
            [string]$range.Cells.Item($extra.HeaderRow,8).Value2 -cne 'Смещение 3'){throw 'Wrong offset captions.'}
        foreach($block in $blocks){
            for($r=$block.FirstRow;$r -le $block.LastRow;$r++){
                $rowKey=[string]$range.Cells.Item($r,1).Value2
                if(-not $rowKey -or $rowKey -eq 'Units'){continue}
                $comment=$range.Cells.Item($r,[int]$block.Columns['note'])
                if($comment.HorizontalAlignment -ne -4152 -or $comment.WrapText){throw 'Wrong comment alignment.'}
                foreach($key in @($block.Columns.Keys)){
                    if($key -in @('key','name','unit','note')){continue}
                    $cell=$range.Cells.Item($r,[int]$block.Columns[$key])
                    if($cell.Interior.Color -ne 16777215 -or $cell.HorizontalAlignment -ne -4108){throw 'Wrong input format.'}
                    if($key -like 'bind*' -or $key -like 'pos*' -or $rowKey -eq 'Type' -or $rowKey -eq 'Shape'){
                        if($cell.Validation.Type -ne 3 -or -not $cell.Validation.Value){throw 'Wrong selector.'}
                    }
                }
            }
        }
        $checks+=,[ordered]@{Name=$item.RangeName;Address=$range.Address();Blocks=$blocks.Count;Passed=$true}
        $top+=$rows+4
    }
    $beforeFormat=$sheet.UsedRange.Formula | ConvertTo-Json -Depth 4 -Compress
    Apply-ConfigNamedRangeBorders $book
    Apply-ConfigUserInputAlignment $book
    Apply-ConfigCommentColumnAlignment $book
    Apply-ConfigShortColumnAlignment $book
    $afterFormat=$sheet.UsedRange.Formula | ConvertTo-Json -Depth 4 -Compress
    if($beforeFormat -cne $afterFormat){throw 'Formatting changed values, comments or formulas.'}
    $book.SaveAs($path,51);$book.Close($false);$book=$null
    $book=$excel.Workbooks.Open($path,0,$true)
    foreach($item in $checks){if($book.Names.Item($item.Name).RefersToRange.Address() -cne $item.Address){throw 'Reopen changed table.'}}
    [ordered]@{Passed=$true;Tables=$checks;Reopened=$true;MigrationRequired=$false} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory ($Stage+'.json')) -Encoding UTF8
    Write-Output 'DIRECT_TEMPLATE: four forms, selectors, offsets and reopen passed.'
}finally{
    if($book){$book.Close($false)}
    if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($ownedPid){$p=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue;if($p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.ToUniversalTime().Ticks -eq $ownedTicks){Stop-Process -Id $ownedPid -Force}}
}
