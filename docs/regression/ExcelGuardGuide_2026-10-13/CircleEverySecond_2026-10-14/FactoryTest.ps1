# ДЛЯ ТЕСТОВ: штатная сборка только таблицы круга в новой несохраняемой книге.
$ErrorActionPreference='Stop'
. (Join-Path (Get-Location).Path 'tools/build_workbook/SettingsCatalog.ps1')
$taskDirectory=Join-Path (Get-Location).Path 'docs/regression/ExcelGuardGuide_2026-10-13/CircleEverySecond_2026-10-14'
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class BindingFactoryIdentity{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hWnd,out uint id);}'
$excel=$null;$book=$null;$ownedId=0;$ownedTicks=0;$checks=0
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false
    [uint32]$id=0;[void][BindingFactoryIdentity]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $ownedId=[int]$id;$ownedTicks=(Get-Process -Id $ownedId).StartTime.ToUniversalTime().Ticks
    $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1);$sheet.Name='Config'
    Reset-ConfigValidationSources $sheet
    $catalog=@(Get-GeometrySettingsCatalog | Where-Object {$_.RangeName -eq 'rngCircleGeometry'})[0]
    Add-GeometrySettingsTable $book $sheet $catalog.RangeName 3 1 $catalog.Title $catalog.Rows
    $range=$book.Names.Item('rngCircleGeometry').RefersToRange
    foreach($pair in @(@('Rebar.Bind2row','rngCircleBind2Choices'),@('Rebar.Bind3row','rngCircleBind3Choices'))) {
        $found=$false
        for($row=1;$row -le $range.Rows.Count;$row++) {
            if([string]$range.Cells.Item($row,1).Value2 -cne $pair[0]){continue}
            $found=$true;$cell=$range.Cells.Item($row,2)
            if([string]$cell.Value2 -cne 'EachBar'){throw 'Factory changed the default'};$checks++
            if([string]$cell.Validation.Formula1 -cne ('='+$pair[1])){throw 'Factory list is not independently named'};$checks++
            foreach($value in @('EachBar','EverySecondBar')){$cell.Value2=$value;if(-not $cell.Validation.Value){throw 'Factory list rejects option'};$checks++}
        }
        if(-not $found){throw 'Factory omitted a circle binding key'};$checks++
    }
    if($book.Names.Item('rngCircleBind2Choices').RefersToRange.Address -ceq $book.Names.Item('rngCircleBind3Choices').RefersToRange.Address){throw 'Factory shares binding lists'};$checks++
    [ordered]@{Passed=$true;Checks=$checks;CircleRows=$range.Rows.Count;IndependentLists=$true;BothOptionsAccepted=$true} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $taskDirectory 'FactoryTest.json') -Encoding UTF8
    Write-Output ('CIRCLE_FACTORY_PASS: checks='+$checks)
} finally {
    if($book){$book.Close($false)}
    if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($ownedId){$owned=Get-Process -Id $ownedId -ErrorAction SilentlyContinue;if($owned -and $owned.ProcessName -eq 'EXCEL' -and $owned.StartTime.ToUniversalTime().Ticks -eq $ownedTicks){Stop-Process -Id $ownedId -Force}}
}
