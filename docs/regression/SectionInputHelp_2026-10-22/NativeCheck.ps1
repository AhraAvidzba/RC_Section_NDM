# Проверяет реальные переходы по всем ссылкам сечений и читает фактический VBA.
# Книга открывается только для чтения; расчёт и формирование сечений не запускаются.
param([string]$WorkbookPath='docs/regression/SectionInputHelp_2026-10-22/Release/RC_Section_NDM.xlsm',
      [string]$OutputStem='NativeCheck')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path
$path=(Resolve-Path -LiteralPath (Join-Path $root $WorkbookPath)).Path
$hash=(Get-FileHash -LiteralPath $path).Hash
$edit=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'Native02/NativeEdit.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$definitions=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'Native02/Catalog.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$geometry=Get-Content -LiteralPath (Join-Path $root 'tools/build_workbook/GeometryTables.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$checks=@();$excel=$null;$book=$null;$ownedPid=0;$ownedTicks=0
$native=New-Object Text.StringBuilder
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class SectionHelpCheckWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint id);}'
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$true;$excel.AutomationSecurity=1
    [uint32]$id=0;[void][SectionHelpCheckWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $ownedPid=[int]$id;$ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    Write-Output "OWNED_EXCEL: id=$ownedPid; utcStartTicks=$ownedTicks"
    $book=$excel.Workbooks.Open($path,0,$true)
    if (-not $book.ReadOnly) {throw 'Verification must be read-only.'}
    $sheet=$book.Worksheets.Item('Справка')
    $main=$sheet.Range($edit.MainHeadingAddress)
    if ([string]$main.Value2 -cne 'Справка по настройкам Config и листа «Сечения»') {throw 'Main heading differs.'}
    if ([bool]$sheet.Rows.Item(($main.Row+1)).Hidden) {throw 'Workbook_Open did not reveal the renamed settings section.'}
    foreach ($definition in $geometry) {
        $item=@($definitions | Where-Object Title -CEQ $definition.Title)[0]
        $block=@($edit.Blocks | Where-Object Key -eq $item.Key)[0]
        $range=$book.Names.Item($definition.RangeName).RefersToRange
        $header=$range.Worksheet.Cells.Item(($range.Row-1),$range.Column)
        if ($header.Hyperlinks.Count -ne 1) {throw 'A section header has no unique link.'}
        $expected="'Справка'!A"+$block.First
        if ([string]$header.Hyperlinks.Item(1).SubAddress -cne $expected) {throw 'Wrong help destination.'}
        $start=([string]$header.Value2).IndexOf('(Подробнее)')+1
        if ($start -lt 1 -or $header.Characters($start,11).Font.Color -ne 16711680) {throw 'Section help link is not visibly blue.'}
        $sheet.Outline.ShowLevels(1)
        $range.Worksheet.Activate();$header.Hyperlinks.Item(1).Follow()
        if ($excel.ActiveSheet.Name -cne 'Справка' -or $excel.ActiveCell.Row -ne $block.First -or [bool]$sheet.Rows.Item($block.First).Hidden) {throw 'Header click did not reveal its actual help block.'}
        $checks+=,[ordered]@{Type=$definition.Type;Link=$header.Address($false,$false);Target=$expected;Followed=$true;TargetVisible=$true}
        $common=$sheet.Cells.Item($block.Last,2)
        if ($common.Hyperlinks.Count -ne 1 -or $common.Font.Color -ne 16711680 -or $common.Font.Underline -ne 2) {throw 'Common help link is missing or not styled.'}
        $common.Hyperlinks.Item(1).Follow()
        $commonBlock=@($edit.Blocks | Where-Object Key -eq 'DependentRebarRows')[0]
        if ($excel.ActiveCell.Row -ne $commonBlock.First -or [bool]$sheet.Rows.Item($commonBlock.First).Hidden) {throw 'Common help link failed.'}
        $checks+=,[ordered]@{Type=$definition.Type;Link=$common.Address($false,$false);Target=[string]$common.Hyperlinks.Item(1).SubAddress;Followed=$true;TargetVisible=$true}
    }
    [void]$native.AppendLine('VBA PROJECT EXPORT')
    [void]$native.AppendLine('Workbook: '+$book.Name)
    [void]$native.AppendLine('Date: '+(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
    $count=[int]$book.VBProject.VBComponents.Count
    for ($i=1;$i -le $count;$i++) {
        $component=$book.VBProject.VBComponents.Item($i);$lines=[int]$component.CodeModule.CountOfLines
        [void]$native.AppendLine('='*100)
        [void]$native.AppendLine('COMPONENT: '+$component.Name)
        [void]$native.AppendLine('TYPE: '+$component.Type)
        [void]$native.AppendLine('LINES: '+$lines)
        [void]$native.AppendLine('='*100)
        if ($lines) {[void]$native.AppendLine($component.CodeModule.Lines(1,$lines))}
        [void]$native.AppendLine()
    }
    $book.Close($false);$book=$null
    if ((Get-FileHash -LiteralPath $path).Hash -cne $hash) {throw 'Verification changed the workbook.'}
    [IO.File]::WriteAllText((Join-Path $PSScriptRoot ($OutputStem+'VBA.txt')),$native.ToString(),(New-Object Text.UTF8Encoding($true)))
    [ordered]@{Passed=$true;ReadOnly=$true;WorkbookSHA256=$hash;WorkbookOpenReveal=$true;FollowedLinks=$checks;Components=$count;OwnPID=$ownedPid;OwnStartTicks=$ownedTicks;SourceUnchanged=$true} |
        ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $PSScriptRoot ($OutputStem+'.json')) -Encoding UTF8
    Write-Output ('HELP_NATIVE_PASS: links='+$checks.Count+'; components='+$count+'; workbook unchanged=True')
} finally {
    if ($book) {$book.Close($false)}
    if ($excel) {$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if ($ownedPid) {
        $owned=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue
        if ($owned -and $owned.ProcessName -eq 'EXCEL' -and $owned.StartTime.ToUniversalTime().Ticks -eq $ownedTicks) {Stop-Process -Id $ownedPid -Force}
    }
}
