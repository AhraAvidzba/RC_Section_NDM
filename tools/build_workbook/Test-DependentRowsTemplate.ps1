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
    Reset-ConfigValidationSources $sheet
    Add-UnitSettingsTable $book $sheet 3 1
    $top=30
    foreach($item in (Get-GeometrySettingsCatalog)){
        if($item.CircleTable){Add-CircleSettingsTable $book $sheet $item.RangeName $top 8 $item.Title $item.Rows;$rows=11}
        elseif($item.RoundedRectangleTable){Add-RoundedRectangleSettingsTable $book $sheet $item.RangeName $top 8 $item.Title $item.Rows;$rows=22}
        elseif($item.HollowRectangleTable){Add-HollowRectangleSettingsTable $book $sheet $item.RangeName $top 8 $item.Title $item.Rows;$rows=28}
        else{Add-RectSetFaceSettingsTable $book $sheet $item.RangeName $top 8 $item.Title $item.Rows;$rows=28}
        $range=$book.Names.Item($item.RangeName).RefersToRange
        if($range.Columns.Count -ne 11 -or $range.Rows.Count -ne $rows){throw "Wrong range: $($item.RangeName)"}
        $header=0
        for($r=1;$r -le $range.Rows.Count;$r++){if([string]$range.Cells.Item($r,2).Value2 -eq 'd2'){$header=$r}}
        $expected=@('грань','d2','отступ','привязка','положение','d3','отступ','привязка','положение','ед','комментарий')
        for($c=1;$c -le 11;$c++){if([string]$range.Cells.Item($header,$c).Value2 -cne $expected[$c-1]){throw 'Wrong caption.'}}
        for($r=$header+1;$r -le $range.Rows.Count;$r++){
            foreach($c in 3,7){if($range.Cells.Item($r,$c).Value2 -ne 0 -or $range.Cells.Item($r,$c).Validation.Type -ne 2){throw 'Wrong offset input.'}}
            foreach($c in 4,5,8,9){if($range.Cells.Item($r,$c).Validation.Type -ne 3 -or -not $range.Cells.Item($r,$c).Validation.Value){throw 'Wrong selector.'}}
        }
        $checks+=,[ordered]@{Name=$item.RangeName;Address=$range.Address();Header=$header;Passed=$true}
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
