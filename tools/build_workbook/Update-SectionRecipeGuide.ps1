# Направленная native-приёмка общего updater на отдельной книге.
param([Parameter(Mandatory=$true)][string]$WorkbookPath,
    [Parameter(Mandatory=$true)][string]$ReportDirectory,[switch]$Verify,[string]$PreviousLayoutPath='')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
. (Join-Path $PSScriptRoot 'SectionRecipeGuide.ps1')
$path=(Resolve-Path -LiteralPath (Join-Path $root $WorkbookPath)).Path
$allowed=[IO.Path]::GetFullPath((Join-Path $root 'docs/regression'))+[IO.Path]::DirectorySeparatorChar
if (-not $path.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)) {throw 'Only isolated workbooks may be updated.'}
$directory=[IO.Path]::GetFullPath((Join-Path $root $ReportDirectory))
if (-not ($directory+[IO.Path]::DirectorySeparatorChar).StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)) {throw 'Evidence must stay in docs/regression.'}
New-Item -ItemType Directory -Path $directory -Force | Out-Null
$production=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$productionHash=(Get-FileHash -LiteralPath $production).Hash
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class RecipeGuideWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hWnd,out uint id);}'
function States($application) {return [ordered]@{Calculation=[int]$application.Calculation;ScreenUpdating=[bool]$application.ScreenUpdating;
    EnableEvents=[bool]$application.EnableEvents;DisplayAlerts=[bool]$application.DisplayAlerts}}
function Assert-States($before,$application) {if (($before | ConvertTo-Json -Compress) -cne ((States $application) | ConvertTo-Json -Compress)) {throw 'Excel states were not restored.'}}
function Shape-Snapshot($sheet) {
    return @($sheet.Shapes | ForEach-Object {[ordered]@{Name=$_.Name;Type=$_.Type;Top=$_.Top;Left=$_.Left;Width=$_.Width;Height=$_.Height;
        Row=[int]$_.TopLeftCell.Row;Offset=[double]$_.Top-[double]$sheet.Rows.Item($_.TopLeftCell.Row).Top}})
}
function Export-GuideRange($book,$sheet,[int]$start,[int]$end,[string]$target) {
    $sheet.Activate();$range=$sheet.Range($sheet.Cells.Item($start,1),$sheet.Cells.Item($end,6))
    $range.EntireRow.Hidden=$false
    $range.CopyPicture(1,-4147)
    $scratch=$book.Worksheets.Add()
    try {
        $chart=$scratch.ChartObjects().Add(0,0,[double]$range.Width,[double]$range.Height)
        $chart.Activate();$chart.Chart.Paste()
        if (-not $chart.Chart.Export($target,'PNG')) {throw 'Guide PNG export failed.'}
    } finally {$book.Application.DisplayAlerts=$false;$scratch.Delete();$sheet.Activate()}
}
$excel=$null;$book=$null;$ownedPid=0;$ownedTicks=0
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=1
    [uint32]$id=0;[void][RecipeGuideWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $ownedPid=[int]$id;$ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    Write-Output ('OWNED_EXCEL: id='+$ownedPid+'; utcStartTicks='+$ownedTicks)
    $book=$excel.Workbooks.Open($path,0,$false)
    if ($book.ReadOnly) {throw 'Private workbook is read-only.'}
    $sheet=$book.Worksheets.Item('Справка')
    $shapesBefore=Shape-Snapshot $sheet
    $namesBefore=@($book.Names | ForEach-Object {$_.Name})
    $excel.Calculation=-4135;$excel.ScreenUpdating=$true;$excel.EnableEvents=$false;$excel.DisplayAlerts=$true
    $state=States $excel
    $first=Update-SectionRecipeGuide $book $root $PreviousLayoutPath
    Assert-States $state $excel
    $first | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $directory 'NativeMap.json') -Encoding UTF8
    $errorStateTests=@()
    if ($Verify) {
        $excel.Calculation=-4105;$excel.ScreenUpdating=$false;$excel.EnableEvents=$true;$excel.DisplayAlerts=$false
        $state=States $excel
        $probe=$sheet.Cells.Item($first.StartRow+1,1);$original=[string]$probe.Value2;$probe.Value2='TEST USER EDIT'
        $reason=''
        try {Update-SectionRecipeGuide $book $root | Out-Null} catch {$reason=$_.Exception.Message}
        Assert-States $state $excel
        if ($reason -notlike '*content was edited*' -or [string]$probe.Value2 -cne 'TEST USER EDIT') {throw 'Guide user edit guard failed.'}
        $probe.Value2=$original
        $errorStateTests+=,[ordered]@{Scenario='User edit';Restored=$true;Preserved=$true;Reason=$reason}
        $sheet.Protect('NDM_GUIDE_TEST')
        $reason=''
        try {Update-SectionRecipeGuide $book $root | Out-Null} catch {$reason=$_.Exception.Message}
        Assert-States $state $excel
        $sheet.Unprotect('NDM_GUIDE_TEST')
        if (-not $reason) {throw 'Protected guide unexpectedly updated.'}
        $errorStateTests+=,[ordered]@{Scenario='Protected sheet';Restored=$true;Reason=$reason}
    }
    $state=States $excel
    $second=Update-SectionRecipeGuide $book $root
    Assert-States $state $excel
    if ($first.StartRow -ne $second.StartRow -or $first.EndRow -ne $second.EndRow -or $first.BodySHA256 -cne $second.BodySHA256) {throw 'Guide updater is not idempotent.'}
    $shapesAfter=Shape-Snapshot $sheet
    if($shapesBefore.Count -ne $shapesAfter.Count){throw 'Old Help shape count changed.'}
    foreach($savedShape in $shapesBefore){
        $actual=@($shapesAfter | Where-Object Name -ceq $savedShape.Name)
        if($actual.Count -ne 1 -or $actual[0].Type -ne $savedShape.Type){throw 'Old Help shape identity changed.'}
        $row=$savedShape.Row
        foreach($move in $first.Outline.Moves){$row=Convert-HelpMovedRow $row $move.First $move.Last $move.Destination}
        $expectedTop=[double]$sheet.Rows.Item($row).Top+$savedShape.Offset
        if($actual[0].Row -ne $row -or [Math]::Abs($actual[0].Top-$expectedTop) -gt 0.05){throw ('Help shape moved relative to its paragraph: '+$savedShape.Name)}
        foreach($field in @('Left','Width','Height')){if([Math]::Abs($actual[0][$field]-$savedShape[$field]) -gt 0.001){throw ('Help shape geometry changed: '+$savedShape.Name)}}
    }
    $namesAfter=@($book.Names | ForEach-Object {$_.Name})
    $added=@($namesAfter | Where-Object {$_ -notin $namesBefore})
    $expectedAdded=1
    if ('rngSectionRecipeGuide' -in $namesBefore) {$expectedAdded=0}
    if ($added.Count -ne $expectedAdded -or ($added.Count -eq 1 -and $added[0] -cne 'rngSectionRecipeGuide')) {throw 'Unexpected new workbook names.'}
    $book.Save();$book.Close($false);$book=$null
    $book=$excel.Workbooks.Open($path,0,$true);$sheet=$book.Worksheets.Item('Справка')
    $range=$book.Names.Item('rngSectionRecipeGuide').RefersToRange
    $lines=@(Get-RecipeGuideNativeLines $sheet $second.StartRow $second.EndRow)
    if ((Get-RecipeGuideTextHash $lines) -cne $second.BodySHA256) {throw 'Guide text changed after reopen.'}
    $navigation=@()
    if ($Verify) {
        $excel.EnableEvents=$true
        $layout=Get-Content -LiteralPath (Join-Path $root 'docs/geometry_refactor/SECTION_RECIPE_GUIDE_LAYOUT.json') -Raw -Encoding UTF8 | ConvertFrom-Json
        # Адресные переходы после reopen: полный код, длинные сигнатуры, options.
        foreach ($title in @('Полный рабочий пример','CSectionDefinitionBuilder','CRebarRuleOptions')) {
            $section=@($layout.sections | Where-Object title -ceq $title)[0]
            if ($null -eq $section) {throw ('Navigation section is missing: '+$title)}
            $row=$second.StartRow+[int]$section.index
            $sheet.Range($sheet.Cells.Item($row+1,1),$sheet.Cells.Item($row+3,1)).EntireRow.Hidden=$true
            $link=@($sheet.Hyperlinks | Where-Object { $_.SubAddress -ceq ("'Справка'!A"+$row) })[0]
            if ($null -eq $link) {throw ('TOC link missing: '+$title)}
            $link.Follow($false) | Out-Null
            $visible=-not [bool]$sheet.Rows.Item($row+1).Hidden
            if (-not $visible) {throw ('Navigation did not reveal body: '+$title)}
            $navigation+=,[ordered]@{Title=$title;TargetRow=$row;BodyVisible=$visible}
        }
        Export-GuideRange $book $sheet $second.StartRow ($second.StartRow+32) (Join-Path $directory 'GuideToc.png')
        $apiRow=[int]$second.ApiRows['api-csectiondefinitionbuilder-rebarprojected']
        if ($apiRow -lt $second.StartRow) {throw 'Signature anchor is missing.'}
        Export-GuideRange $book $sheet $apiRow ($apiRow+21) (Join-Path $directory 'GuideSignature.png')
        $example=@($layout.rows | Where-Object { $_.source -ceq 'ADDING_SECTION_TYPE.md' -and $_.text -ceq 'VERSION 1.0 CLASS' })[0]
        $exampleIndex=[Array]::IndexOf([object[]]$layout.rows,$example)
        Export-GuideRange $book $sheet ($second.StartRow+$exampleIndex+16) ($second.StartRow+$exampleIndex+43) (Join-Path $directory 'GuideCode.png')
        $options=@($layout.sections | Where-Object title -ceq 'CRebarRuleOptions')[0]
        $optionsRow=$second.StartRow+[int]$options.index
        Export-GuideRange $book $sheet $optionsRow ($optionsRow+23) (Join-Path $directory 'GuideOptions.png')
    }
    $book.Close($false);$book=$null
    if ((Get-FileHash -LiteralPath $production).Hash -cne $productionHash) {throw 'Production workbook changed.'}
    [ordered]@{Status='DIRECTED_NATIVE_GUIDE_PASS';ExcelVersion=$excel.Version;ExcelBuild=$excel.Build;
        WorkbookSHA256=(Get-FileHash -LiteralPath $path).Hash;ProductionSHA256=$productionHash;ProductionChanged=$false;
        Idempotent=$true;StatesSuccess=$true;StatesFailure=$errorStateTests;ReopenTextMatch=$true;OldShapesPreserved=$true;
        AddedNames=$added;First=$first;Second=$second;Navigation=$navigation;ApiStatus=$second.ApiStatus} |
        ConvertTo-Json -Depth 7 | Set-Content -LiteralPath (Join-Path $directory 'GuideEvidence.json') -Encoding UTF8
    Write-Output ('NATIVE_GUIDE_PASS: rows='+$second.Rows+'; members='+$second.Members+'; current API status='+$second.ApiStatus)
} catch {
    [ordered]@{Status='FAIL';Reason=$_.Exception.Message;ScriptStackTrace=$_.ScriptStackTrace;
        Native=$_.ToString();ProductionUnchanged=((Get-FileHash -LiteralPath $production).Hash -ceq $productionHash)} |
        ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $directory ('FailedProbe_'+(Get-Date -Format 'yyyyMMdd_HHmmss')+'.json')) -Encoding UTF8
    throw
} finally {
    if ($book) {$book.Close($false)}
    if ($excel) {$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if ($ownedPid) {
        $owned=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue
        if ($owned -and $owned.ProcessName -eq 'EXCEL' -and $owned.StartTime.ToUniversalTime().Ticks -eq $ownedTicks) {Stop-Process -Id $ownedPid -Force}
    }
}
