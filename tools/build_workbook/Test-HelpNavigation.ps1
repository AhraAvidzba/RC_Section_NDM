# -------------------- ДЛЯ ТЕСТОВ --------------------
# Проверяет штатные события открытия и ссылок на отдельной read-only копии,
# затем применяет только макет Справки для контроля следующей сборки.
param([Parameter(Mandatory=$true)][string]$WorkbookPath,[Parameter(Mandatory=$true)][string]$EvidencePath)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'HelpOutline.ps1')
. (Join-Path $PSScriptRoot 'WorkbookLayout.ps1')
$path=(Resolve-Path -LiteralPath $WorkbookPath).Path
$output=[IO.Path]::GetFullPath($EvidencePath);[IO.Directory]::CreateDirectory($output) | Out-Null
$before=(Get-FileHash -LiteralPath $path).Hash
$existing=@(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class HelpNavigationTestWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
$app=$null;$book=$null;$owned=0;$ticks=0
# -------------------- ДЛЯ ТЕСТОВ --------------------
# Отличает собственный Excel по PID и времени старта перед закрытием объектов.
function Test-HelpNavigationOwnedExcel {
    if(-not $owned){return $false}
    $p=Get-Process -Id $owned -ErrorAction SilentlyContinue
    return ($p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.ToUniversalTime().Ticks -eq $ticks)
}
try {
    $app=New-Object -ComObject Excel.Application
    [uint32]$id=0;[void][HelpNavigationTestWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id)
    if($existing -contains [int]$id){throw 'Excel COM attached to a user process.'}
    $owned=[int]$id;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
    $app.Visible=$false;$app.DisplayAlerts=$false;$app.AutomationSecurity=1;$app.EnableEvents=$true
    $book=$app.Workbooks.Open($path,0,$true)
    if(-not $book.ReadOnly){throw 'Expected read-only workbook.'}
    $sheet=$book.Worksheets.Item('Справка');$rows=Get-HelpMainRows $sheet
    $config=[int]$rows[$script:HelpMainTitles[-1]];$last=Get-HelpContentLastRow $sheet
    if($sheet.Rows.Item('2:85').Hidden -ne $true -or $sheet.Rows.Item(($config+1).ToString()+':'+$last).Hidden -ne $false){throw 'Workbook_Open did not set Help defaults.'}
    $setting=$book.Worksheets.Item('Config').Cells.Find('General.ReformatResultsEnabled',$book.Worksheets.Item('Config').Cells.Item(1,1),-4163,1)
    $links=@($book.Worksheets.Item('Config').Cells.Item($setting.Row,5).Hyperlinks.Item(1))
    foreach($name in @('rngCircleGeometry','rngRectanglesGeometry','rngPointedSectionGeometry','rngHollowRectangleGeometry')){
        $range=$book.Names.Item($name).RefersToRange
        $links+=,$range.Worksheet.Cells.Item($range.Row-1,$range.Column).Hyperlinks.Item(1)
    }
    $checks=@()
    foreach($link in $links){
        $book.Activate();$sheet.Activate();$sheet.Outline.ShowLevels(1) | Out-Null
        $target=[string]$link.SubAddress
        if($target -notmatch "^'?Справка'?!([A-Z]+)([0-9]+)$"){throw 'Unexpected Help link target.'}
        $targetRow=[int]$Matches[2]
        $link.Follow()
        if($app.ActiveSheet.Name -cne 'Справка' -or $app.ActiveWindow.ScrollRow -ne $targetRow -or $sheet.Rows.Item(($config+1).ToString()+':'+$last).Hidden -ne $false){throw ('Help navigation failed: '+$target)}
        $checks+=,[ordered]@{Target=$target;ScrollRow=[int]$app.ActiveWindow.ScrollRow;WholeSectionExpanded=$true}
    }
    $layout=Read-WorkbookLayout
    $helpOnly=[pscustomobject]@{Names=@();OtherTables=@();Sheets=@($layout.Sheets | Where-Object Name -eq 'Справка');SectionHeaders=$layout.SectionHeaders}
    $app.EnableEvents=$false
    Apply-WorkbookLayout $book $helpOnly
    $expected=Get-Content (Join-Path $output '../FinalPackage.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach($p in $expected.Pictures){$shape=$sheet.Shapes.Item($p.Name);if($shape.TopLeftCell.Row -ne $p.Row -or [Math]::Abs($shape.Height-$p.Height) -gt 0.4){throw ('Build layout damaged picture: '+$p.Name)}}
    [ordered]@{Status='PASS';WorkbookSHA256=$before;WorkbookOpenDefaults=$true;Links=$checks;BuildLayoutPictures=$expected.Pictures.Count;OwnedPID=$owned;StartTicks=[string]$ticks} | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $output 'Navigation.json') -Encoding UTF8
    Write-Output ('HELP_NAVIGATION_PASS: links='+$checks.Count+'; buildLayoutPictures='+$expected.Pictures.Count)
} finally {
    if(Test-HelpNavigationOwnedExcel){if($book){$book.Close($false)};if($app){$app.Quit()}}
    if($app){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if(Test-HelpNavigationOwnedExcel){Stop-Process -Id $owned -Force}
}
if((Get-FileHash -LiteralPath $path).Hash -cne $before){throw 'Source workbook changed during read-only test.'}
