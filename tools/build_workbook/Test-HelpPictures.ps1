# -------------------- ДЛЯ ТЕСТОВ --------------------
# Проверяет встроенные рисунки справки в отдельном Excel и восстанавливает карту
# строк по действующему сборщику. Исходная книга открывается только для чтения.
param([Parameter(Mandatory=$true)][string]$WorkbookPath,[Parameter(Mandatory=$true)][string]$EvidencePath,[switch]$Strict)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
. (Join-Path $PSScriptRoot 'HelpOutline.ps1')
$path=(Resolve-Path -LiteralPath $WorkbookPath).Path
$output=[IO.Path]::GetFullPath($EvidencePath)
[IO.Directory]::CreateDirectory($output) | Out-Null
$script:FormulaImageDir=Join-Path $output 'FormulaSources'
$before=(Get-FileHash -LiteralPath $path).Hash
$existing=@(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class HelpPictureTestWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
$app=$null;$book=$null;$scratch=$null;$owned=0;$ticks=0
# -------------------- ДЛЯ ТЕСТОВ --------------------
# Проверяет идентичность собственного процесса перед каждым закрытием COM-книг.
function Test-HelpOwnedExcel {
    if(-not $owned){return $false}
    $p=Get-Process -Id $owned -ErrorAction SilentlyContinue
    return ($p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.ToUniversalTime().Ticks -eq $ticks)
}
# -------------------- ДЛЯ ТЕСТОВ --------------------
# Снимает диапазон средствами Excel; временный график используется только как PNG-контейнер.
function Export-TestHelpRange($Source,$Destination,[int]$First,[int]$Last,[string]$Name) {
    $range=$Source.Range("A${First}:F${Last}")
    $book.Activate();$Source.Activate();$app.Goto($range,$true)
    $range.CopyPicture(1,2) | Out-Null
    $chart=$Destination.ChartObjects().Add(0,0,[double]$range.Width,[double]$range.Height)
    try {
        $scratch.Activate();$Destination.Activate();$chart.Activate();$chart.Chart.Paste() | Out-Null
        if(-not $chart.Chart.Export((Join-Path $output ($Name+'.png')),'PNG')){throw 'Excel did not export Help capture.'}
    } finally {$chart.Delete()}
}
try {
    $app=New-Object -ComObject Excel.Application
    [uint32]$id=0;[void][HelpPictureTestWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id)
    if($existing -contains [int]$id){throw 'Excel COM attached to an existing user process.'}
    $owned=[int]$id;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
    $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=3
    $book=$app.Workbooks.Open($path,0,$true)
    if(-not $book.ReadOnly){throw 'Expected read-only workbook.'}
    $sheet=$book.Worksheets.Item('Справка');$rows=Get-HelpMainRows $sheet
    $sheet.Rows.Hidden=$false
    $actual=@(foreach($shape in $sheet.Shapes){[ordered]@{Name=[string]$shape.Name;Row=[int]$shape.TopLeftCell.Row;Height=[double]$shape.Height;Width=[double]$shape.Width;Description=[string]$shape.AlternativeText}})
    $scratch=$app.Workbooks.Add();$temp=$scratch.Worksheets.Item(1)
    for($c=1;$c -le 6;$c++){$temp.Columns.Item($c).ColumnWidth=[double]$sheet.Columns.Item($c).ColumnWidth}
    $script:FormulaImageCounter=0;$script:FormulaImageRows.Clear()
    $null=Add-DirectStateMethodologyGuide $temp ([int]$rows[$script:HelpMainTitles[1]]+1)
    $null=Add-CapacityMethodologyGuide $temp ([int]$rows[$script:HelpMainTitles[2]]+1)
    $null=Add-CrackWidthMethodologyGuide $temp ([int]$rows[$script:HelpMainTitles[3]]+1)
    $null=Add-StabilityMethodologyGuide $temp ([int]$rows[$script:HelpMainTitles[4]]+1)
    $null=Add-ResultsSheetGuide $temp ([int]$rows[$script:HelpMainTitles[5]]+1)
    $expected=@(foreach($item in $script:FormulaImageRows){$shape=$temp.Shapes.Item($item.Name);[ordered]@{Name=$item.Name;Row=$item.Row;RowHeight=$item.Height;Width=$item.PictureWidth;Height=$item.PictureHeight;Description=[string]$shape.AlternativeText;TargetBlank=([string]$sheet.Cells.Item($item.Row,1).Value2 -eq '');TargetMerged=[bool]$sheet.Cells.Item($item.Row,1).MergeCells}})
    if($Strict){
        if($actual.Count -ne $expected.Count){throw 'Help picture count differs from builder.'}
        foreach($picture in $expected){
            $shape=$sheet.Shapes.Item($picture.Name)
            if($shape.TopLeftCell.Row -ne $picture.Row -or [Math]::Abs($shape.Height-$picture.Height) -gt 0.4 -or [Math]::Abs($shape.Width-$picture.Width) -gt 0.4 -or $shape.AlternativeText -cne $picture.Description){throw ('Invalid Help picture: '+$picture.Name)}
        }
        $outline=@()
        $book.Activate();$sheet.Activate() # Excel ShowDetail требует активного листа при открытой вспомогательной книге.
        for($i=0;$i -lt $script:HelpMainTitles.Count;$i++){
            $first=[int]$rows[$script:HelpMainTitles[$i]];$last=Get-HelpContentLastRow $sheet
            if($i+1 -lt $script:HelpMainTitles.Count){$last=[int]$rows[$script:HelpMainTitles[$i+1]]-1}
            $body=$sheet.Rows.Item(($first+1).ToString()+':'+$last)
            try{$sheet.Rows.Item($first).ShowDetail=$false}catch{
                $failure=$_.Exception.Message
                $details=@(foreach($shape in $sheet.Shapes){[ordered]@{Name=$shape.Name;Top=$shape.Top;Height=$shape.Height;From=$shape.TopLeftCell.Row;To=$shape.BottomRightCell.Row;Placement=$shape.Placement}})
                $details | ConvertTo-Json -Depth 4 | Set-Content (Join-Path $output 'OutlineFailure.json') -Encoding UTF8
                throw ('Cannot collapse Help section '+$script:HelpMainTitles[$i]+' at row '+$first+': '+$failure)
            }
            if($body.Hidden -ne $true){throw 'Help section did not collapse completely.'}
            $sheet.Rows.Item($first).ShowDetail=$true
            if($body.Hidden -ne $false){throw 'Help section contains hidden rows after expansion.'}
            $outline+=,[ordered]@{First=$first;Last=$last;Collapsed=$true;Expanded=$true}
        }
        foreach($picture in $expected){if([Math]::Abs($sheet.Shapes.Item($picture.Name).Height-$picture.Height) -gt 0.4){throw 'Outline toggle damaged Help picture.'}}
        $setting=$sheet.Cells.Find('General.ReformatResultsEnabled',$sheet.Cells.Item(1,1),-4163,1)
        $first=[int]$setting.Row
        $entry=Get-SettingsInstructionCatalog | Where-Object Key -eq 'General.ReformatResultsEnabled'
        $paragraphCount=@($entry.Lines | Where-Object {-not [string]::IsNullOrWhiteSpace([string]$_)}).Count
        $sheet.Range("A$($first+1):F$($first+$paragraphCount)").EntireRow.AutoFit()
        if($setting.MergeArea.Address() -cne ('$A$'+$first+':$F$'+$first)){throw 'Setting heading is not merged A:F.'}
        $heights=@()
        for($r=$first+1;$r -le $first+$paragraphCount;$r++){
            if([string]$sheet.Cells.Item($r,1).Value2 -cne [string]($r-$first) -or $sheet.Cells.Item($r,2).MergeCells -or $sheet.Cells.Item($r,2).Font.Size -ne 10){throw 'Setting body differs from common numbered layout.'}
            $heights+=,[ordered]@{Row=$r;Height=[double]$sheet.Rows.Item($r).RowHeight}
        }
        foreach($link in $sheet.Hyperlinks){if([string]$link.SubAddress -match "^'?Config'?!"){throw 'Reverse Help-to-Config link remains.'}}
        Export-TestHelpRange $sheet $temp $first ($first+$paragraphCount) 'ReformatSetting'
        Export-TestHelpRange $sheet $temp 96 103 'DirectFormulas'
        Export-TestHelpRange $sheet $temp 183 188 'CrackFraction'
        Export-TestHelpRange $sheet $temp 305 309 'SP35Diagram'
        $book.Activate();$sheet.Activate()
        $book.SaveCopyAs((Join-Path $output 'ExpandedNative.xlsm'))
        foreach($row in $rows.Values){if([int]$row -ne [int]$rows[$script:HelpMainTitles[-1]]){$sheet.Rows.Item([int]$row).ShowDetail=$false}}
        $roundTrip=Join-Path $output 'RoundTrip.xlsm';$book.SaveCopyAs($roundTrip)
        $probe=$app.Workbooks.Open($roundTrip,0,$true)
        try {
            $probe.Worksheets.Item('Справка').Rows.Hidden=$false
            foreach($picture in $expected){$shape=$probe.Worksheets.Item('Справка').Shapes.Item($picture.Name);if($shape.TopLeftCell.Row -ne $picture.Row -or [Math]::Abs($shape.Height-$picture.Height) -gt 0.4){throw ('Saved collapsed picture was lost: '+$picture.Name)}}
        } finally {if(Test-HelpOwnedExcel){$probe.Close($false)}}
        [ordered]@{Status='PASS';Pictures=$expected.Count;Outline=$outline;ReformatRow=$first;ReformatHeights=$heights;SavedCollapsedReopen=$true} | ConvertTo-Json -Depth 10 | Set-Content (Join-Path $output 'NativeChecks.json') -Encoding UTF8
    }
    [ordered]@{WorkbookSHA256=$before;OwnedPID=$owned;StartTicks=[string]$ticks;MainRows=$rows;Actual=$actual;Expected=$expected} | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $output 'Pictures.json') -Encoding UTF8
    Write-Output ('HELP_PICTURES: actual='+$actual.Count+'; expected='+$expected.Count+'; zeroHeightExpanded='+@($actual | Where-Object Height -lt 1).Count+'; nonblankTargets='+@($expected | Where-Object {-not $_.TargetBlank}).Count)
} finally {
    if(Test-HelpOwnedExcel){if($scratch){$scratch.Close($false)};if($book){$book.Close($false)};if($app){$app.Quit()}}
    if($app){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if(Test-HelpOwnedExcel){Stop-Process -Id $owned -Force}
}
if((Get-FileHash -LiteralPath $path).Hash -cne $before){throw 'Read-only source workbook changed.'}
