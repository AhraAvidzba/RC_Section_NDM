# Оформляет уже готовую новую схему Config и обновляет Справку в отдельной копии.
# Старые таблицы не преобразует и пользовательскую рабочую книгу не открывает.
param([string]$SourceWorkbook='docs/regression/DependentRowOffsets_2026-10-17/Accepted.xlsm',[string]$Stage='PreparedRelease',[switch]$CaptureOnly)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
. (Join-Path $PSScriptRoot 'SectionRecipeGuide.ps1')
. (Join-Path $PSScriptRoot 'StandaloneHelpText.ps1')
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$directory=(Resolve-Path -LiteralPath 'docs/regression/DependentRowOffsets_2026-10-17').Path
$source=(Resolve-Path -LiteralPath $SourceWorkbook).Path
if(-not $source.StartsWith($directory+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase) -or $Stage -notmatch '^[A-Za-z][A-Za-z0-9_]*$'){throw 'Only local evidence copies are allowed.'}
$path=Join-Path $directory ($Stage+'.xlsm')
if(-not $CaptureOnly){
    if(Test-Path -LiteralPath $path){throw 'Prepared copy exists.'}
    Copy-Item -LiteralPath $source -Destination $path
}elseif(-not (Test-Path -LiteralPath $path)){throw 'No prepared file to capture.'}
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class RowsReleaseWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hWnd,out uint id);}'
function Picture($Book,$Range,[string]$Path){
    $Range.Worksheet.Activate();$Range.EntireRow.Hidden=$false;$Range.CopyPicture(1,-4147)
    $scratch=$Book.Worksheets.Add()
    try{$chart=$scratch.ChartObjects().Add(0,0,[double]$Range.Width,[double]$Range.Height);$chart.Activate();$chart.Chart.Paste();if(-not $chart.Chart.Export($Path,'PNG')){throw 'Picture export failed.'}}
    finally{$scratch.Delete()}
}
$excel=$null;$book=$null;$ownedPid=0;$ownedTicks=0;$records=@()
try{
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3
    [uint32]$identity=0;[void][RowsReleaseWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$identity)
    $ownedPid=[int]$identity;$ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    Write-Output "OWNED_EXCEL: $ownedPid"
    $book=$excel.Workbooks.Open($path,0,$false);$excel.Calculation=-4135
    $sheet=$book.Worksheets.Item('Config')
    if(-not $CaptureOnly){
    foreach($name in @('rngCircleGeometry','rngRectanglesGeometry','rngPointedSectionGeometry','rngHollowRectangleGeometry')){
        $range=$book.Names.Item($name).RefersToRange
        if($range.Columns.Count -ne 11){throw "Already updated schema required: $name"}
    }
    if([string]$book.Names.Item('rngCircleGeometry').RefersToRange.Cells.Item(6,2).Value2 -cne 'as'){throw 'Final Circle main-row schema is required.'}
    # Весь перечень входных формул/значений снимается до обновления только Справки.
    $inputs=@{}
    foreach($name in (Get-ConfigNamedRangeNames)){$inputs[$name]=$book.Names.Item($name).RefersToRange.Formula | ConvertTo-Json -Depth 4 -Compress}
    $sheet.Columns.Item(18).ColumnWidth=45
    Apply-ConfigNamedRangeBorders $book
    Apply-ConfigUserInputAlignment $book
    Apply-ConfigCommentColumnAlignment $book
    Apply-ConfigShortColumnAlignment $book
    try{$book.Names.Item('rngSectionRecipeGuide').Delete()}catch{}
    Add-SettingsInstructions $book $sheet $book.Worksheets.Item('Справка')
    Update-StandaloneHelpText $book | Out-Null
    foreach($name in $inputs.Keys){
        $after=$book.Names.Item($name).RefersToRange.Formula | ConvertTo-Json -Depth 4 -Compress
        if($after -cne $inputs[$name]){throw "Help changed input table: $name"}
    }
    $book.Worksheets.Item('Расчет').Activate()
    $book.Save();$book.Close($false);$book=$null
    $guideDirectory=Join-Path $directory ($Stage+'_Guide')
    & C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Update-SectionRecipeGuide.ps1') -WorkbookPath $path.Substring($root.Length+1) -ReportDirectory $guideDirectory.Substring($root.Length+1)
    if($LASTEXITCODE -ne 0){throw 'Guide verification failed.'}
    $guide=Get-Content -LiteralPath (Join-Path $guideDirectory 'NativeMap.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    }else{$book.Close($false);$book=$null}
    $book=$excel.Workbooks.Open($path,0,$true);$excel.ScreenUpdating=$true
    foreach($name in @('rngCircleGeometry','rngRectanglesGeometry','rngPointedSectionGeometry','rngHollowRectangleGeometry')){
        $range=$book.Names.Item($name).RefersToRange
        Picture $book ($range.Offset(-1,0).Resize($range.Rows.Count+1,11)) (Join-Path $directory ($Stage+'_'+$name+'.png'))
        $records+=,[ordered]@{Name=$name;Address=$range.Address();Rows=$range.Rows.Count;Columns=$range.Columns.Count}
    }
    $help=$book.Worksheets.Item('Справка');$found=$help.Cells.Find('Отступы дополнительных рядов арматуры',$help.Cells.Item(1,1),-4163,1)
    if($null -eq $found){throw 'New help section is missing.'}
    Picture $book ($help.Cells.Item($found.Row,1).Resize(9,6)) (Join-Path $directory ($Stage+'_Help.png'))
    [ordered]@{Passed=$true;Workbook=$path;SHA256=(Get-FileHash -LiteralPath $path).Hash;InputsPreserved=$(-not $CaptureOnly);CaptureOnly=[bool]$CaptureOnly;Reopened=$true;Tables=$records;Guide=$guide} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $directory ($Stage+'.json')) -Encoding UTF8
    Write-Output 'PREPARED: Config, Help, input preservation, save/reopen and native pictures passed.'
}catch{
    Write-Output $_.ScriptStackTrace
    Write-Output $_.InvocationInfo.PositionMessage
    throw
}finally{
    if($book){$book.Close($false)}
    if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($ownedPid){$p=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue;if($p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.ToUniversalTime().Ticks -eq $ownedTicks){Stop-Process -Id $ownedPid -Force}}
}
