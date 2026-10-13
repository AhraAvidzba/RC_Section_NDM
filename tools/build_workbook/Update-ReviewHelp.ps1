# Адресное обновление Справки в отдельном кандидате; пользовательский output не открывает.
param([Parameter(Mandatory=$true)][string]$WorkbookPath,[Parameter(Mandatory=$true)][string]$ReportDirectory,[switch]$HalveDefaultRebarLabelOffset)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
. (Join-Path $PSScriptRoot 'SectionRecipeGuide.ps1')
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
. (Join-Path $PSScriptRoot 'StandaloneHelpText.ps1')
$path=(Resolve-Path -LiteralPath (Join-Path $root $WorkbookPath)).Path
$allowed=[IO.Path]::GetFullPath((Join-Path $root 'docs/regression/GeometryReview_2026-10-12'))+[IO.Path]::DirectorySeparatorChar
$endpointAllowed=[IO.Path]::GetFullPath((Join-Path $root 'docs/regression/RoundedEndpoint_2026-10-12'))+[IO.Path]::DirectorySeparatorChar
$fullAllowed=[IO.Path]::GetFullPath((Join-Path $root 'docs/regression/GeometryFull_2026-10-13'))+[IO.Path]::DirectorySeparatorChar
$directory=[IO.Path]::GetFullPath((Join-Path $root $ReportDirectory))
$printAreas=@(Get-WorkbookPrintAreas $path)
foreach ($target in @($path,($directory+[IO.Path]::DirectorySeparatorChar))) {
    if (-not $target.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase) -and -not $target.StartsWith($endpointAllowed,[StringComparison]::OrdinalIgnoreCase) -and -not $target.StartsWith($fullAllowed,[StringComparison]::OrdinalIgnoreCase)) {throw 'Only review evidence paths are allowed.'}
}
New-Item -ItemType Directory -Path $directory -Force | Out-Null
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class ReviewHelpWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hWnd,out uint id);}'

# Сохраняет изображение настоящего Excel-диапазона для визуальной проверки.
function Export-ReviewHelpPicture($Book,$Sheet,[int]$First,[int]$Last,[string]$Path) {
    $Sheet.Activate();$range=$Sheet.Range($Sheet.Cells.Item($First,1),$Sheet.Cells.Item($Last,6))
    $range.CopyPicture(1,-4147)
    $scratch=$Book.Worksheets.Add()
    try {
        $chart=$scratch.ChartObjects().Add(0,0,[double]$range.Width,[double]$range.Height)
        $chart.Activate();$chart.Chart.Paste()
        if (-not $chart.Chart.Export($Path,'PNG')) {throw 'Help PNG export failed.'}
    } finally {$Book.Application.DisplayAlerts=$false;$scratch.Delete();$Sheet.Activate()}
}
$excel=$null;$book=$null;$ownedPid=0;$ownedTicks=0
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=1
    [uint32]$id=0;[void][ReviewHelpWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $ownedPid=[int]$id;$ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    Write-Output ('OWNED_EXCEL: id='+$ownedPid+'; utcStartTicks='+$ownedTicks)
    $book=$excel.Workbooks.Open($path,0,$false)
    if ($book.ReadOnly) {throw 'Candidate is read-only.'}
    $initialSheet=[string]$book.ActiveSheet.Name
    if ($HalveDefaultRebarLabelOffset) {
        # Прямое указание пользователя: прежний стандартный Offset 60 -> 30 мм.
        # Пользовательское значение, отличающееся от прежнего default, не меняем.
        $annotationTable=$book.Names.Item('rngPlotAnnotationSettings').RefersToRange
        $offsetCell=$null
        for ($r=2;$r -le $annotationTable.Rows.Count;$r++) {
            if ([string]$annotationTable.Cells.Item($r,1).Value2 -ceq 'Offset') {$offsetCell=$annotationTable.Cells.Item($r,2);break}
        }
        if ($null -eq $offsetCell) {throw 'Offset is missing in rngPlotAnnotationSettings.'}
        $beforeOffset=$offsetCell.Value2
        if ([double]$beforeOffset -eq 60) {$offsetCell.Value2=30}
        [ordered]@{Key='Plot.RebarLabels.Offset';Address=('Config!'+$offsetCell.Address());Before=$beforeOffset;After=$offsetCell.Value2;
            Reason='User explicitly requested half of the default: 60 -> 30 mm.'} |
            ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'DefaultOffsetChange.json') -Encoding UTF8
    }
    # Последнее решение пользователя: распространяется только книга.
    Update-StandaloneHelpText $book | Out-Null
    $code=[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'WorkbookHelpEvents.vba'),[Text.Encoding]::UTF8)
    $module=$book.VBProject.VBComponents.Item($book.CodeName).CodeModule
    $module.DeleteLines(1,$module.CountOfLines);$module.AddFromString($code)
    foreach ($document in $book.VBProject.VBComponents) {
        if ($document.Type -eq 100 -and $document.Name -ne $book.CodeName -and $document.CodeModule.CountOfLines -eq 0) {
            $document.CodeModule.AddFromString("Option Explicit`r`n`r`n' Лист книги: отображение пользовательских данных и результатов.")
        }
    }
    $first=Update-SectionRecipeGuide $book $root
    $first | ConvertTo-Json -Depth 7 | Set-Content -LiteralPath (Join-Path $directory 'NativeMap.json') -Encoding UTF8
    Write-Output ('GUIDE_UPDATED: '+$first.StartRow+'..'+$first.EndRow)
    $second=Update-SectionRecipeGuide $book $root
    if ($first.StartRow -ne $second.StartRow -or $first.EndRow -ne $second.EndRow) {throw 'Guide update is not idempotent.'}
    $excel.VBE.ActiveVBProject=$book.VBProject
    $compile=$excel.VBE.CommandBars.FindControl(1,578)
    if ($null -eq $compile) {throw 'Compile command unavailable.'}
    if ($compile.Enabled) {$compile.Execute()}
    if ($compile.Enabled) {throw 'Updated workbook events failed native compilation.'}
    # Сохраняем намеренно раскрытую Справку: reopen должен доказать работу события.
    $book.Worksheets.Item('Справка').Outline.ShowLevels(2) | Out-Null
    $book.Save();$book.Close($false);$book=$null
    $excel.EnableEvents=$true
    $book=$excel.Workbooks.Open($path,0,$false)
    $sheet=$book.Worksheets.Item('Справка');$rows=Get-HelpMainRows $sheet
    $probe=$sheet.Cells.Find('Справка по настройкам листа Config',$sheet.Cells.Item(1,1),-4163,1)
    Write-Output ('OPEN_CONFIG_TITLE: '+$probe.Row+'; codeLines='+$book.VBProject.VBComponents.Item($book.CodeName).CodeModule.CountOfLines)
    $states=@()
    foreach ($title in $script:HelpMainTitles) {
        $r=[int]$rows[$title];$expanded=$title -eq $script:HelpMainTitles[-1]
        Write-Output ('CHECK_OUTLINE: '+$title+'; row='+$r+'; protected='+$sheet.ProtectContents+'; summaryRow='+$sheet.Outline.SummaryRow)
        $next=@($rows.Values | Where-Object {$_ -gt $r} | Sort-Object)[0]
        $end=Get-HelpContentLastRow $sheet
        if ($null -ne $next) {$end=[int]$next-1}
        $body=$sheet.Range($sheet.Cells.Item($r+1,1),$sheet.Cells.Item($end,6)).EntireRow
        Write-Output ('BODY: '+($r+1)+'..'+$end+'; hidden='+$body.Hidden+'; outline='+$body.OutlineLevel+'; firstHidden='+$sheet.Rows.Item($r+1).Hidden+'; events='+$excel.EnableEvents)
        if ($null -eq $body.Hidden -or [bool]$body.Hidden -eq $expanded -or [int]$body.OutlineLevel -ne 2) {throw ('Partial or invalid outline body: '+$title)}
        if ([int]$sheet.Rows.Item($r).OutlineLevel -ne 1 -or [int]$sheet.Rows.Item($r+1).OutlineLevel -ne 2 -or [bool]$sheet.Rows.Item($r+1).Hidden -eq $expanded) {throw ('Invalid open outline state: '+$title)}
        if ($expanded) {$sheet.Rows.Item($r).ShowDetail=$false}
        $sheet.Rows.Item($r).ShowDetail=$true
        if ($null -eq $body.Hidden -or [bool]$body.Hidden) {throw 'Full-body expand failed.'}
        $sheet.Rows.Item($r).ShowDetail=$false
        if ($null -eq $body.Hidden -or -not [bool]$body.Hidden) {throw 'Full-body collapse failed.'}
        $states+=,[ordered]@{Title=$title;Row=$r;OpenExpanded=$expanded;Toggle=$true}
    }
    Set-HelpSectionOrderAndOutline $book | Out-Null
    $guide=$book.Names.Item('rngSectionRecipeGuide').RefersToRange
    $lines=@(Get-RecipeGuideNativeLines $sheet $guide.Row ($guide.Row+$guide.Rows.Count-2))
    if ((Get-RecipeGuideTextHash $lines) -cne $first.BodySHA256) {throw 'Reopen guide text differs.'}
    $config=[int]$rows[$script:HelpMainTitles[-1]]
    $settings=@()
    foreach ($key in @('SLS.Crack.ToleranceLambda','Solver.PlateauSlope.Enabled')) {
        $anchor=$sheet.Cells.Find($key,$sheet.Cells.Item(1,1),-4163,1)
        if ($null -eq $anchor -or $anchor.Row -le $config -or [int]$anchor.EntireRow.OutlineLevel -ne 2) {throw ('Setting Help is outside Config: '+$key)}
        $settings+=,[ordered]@{Key=$key;HelpRow=$anchor.Row;InsideConfig=$true}
    }
    $navigation=@()
    $system=$book.Names.Item('rngSystemSettings').RefersToRange
    foreach ($key in @('Solver.PlateauSlope.Enabled','SLS.Crack.ToleranceLambda','Geometry.Type')) {
        $cell=$null
        for ($row=2;$row -le $system.Rows.Count;$row++) {
            if ([string]$system.Cells.Item($row,1).Value2 -ceq $key) {$cell=$system.Cells.Item($row,5);break}
        }
        if ($null -eq $cell -or $cell.Hyperlinks.Count -ne 1) {throw ('Setting link is missing: '+$key)}
        $link=$cell.Hyperlinks.Item(1)
        $address=([string]$link.SubAddress -split '!')[-1]
        $target=$sheet.Range($address)
        $sheet.Rows.Item($config).ShowDetail=$false
        $link.Follow($false) | Out-Null
        if ([bool]$target.EntireRow.Hidden -or -not [bool]$sheet.Rows.Item([int]$rows['Добавление нового типа сечения']+1).Hidden) {throw ('Navigation opened an invalid section: '+$key)}
        $navigation+=,[ordered]@{Key=$key;Address=$address;OnlyTargetSectionOpened=$true}
    }
    $excel.ScreenUpdating=$true
    $sheet.Rows.Item($config).ShowDetail=$false
    $sheet.Rows.Item([int]$rows['Добавление нового типа сечения']).ShowDetail=$true
    Export-ReviewHelpPicture $book $sheet $guide.Row ($guide.Row+24) (Join-Path $directory 'GuideSteps.png')
    Export-ReviewHelpPicture $book $sheet ($guide.Row+25) ($guide.Row+$guide.Rows.Count-2) (Join-Path $directory 'GuideExamples.png')
    Set-HelpSectionOrderAndOutline $book | Out-Null
    Export-ReviewHelpPicture $book $sheet 1 ($config+20) (Join-Path $directory 'HelpDefault.png')
    $sheet.Range('A1').Select()
    $excel.ActiveWindow.ScrollRow=1;$excel.ActiveWindow.ScrollColumn=1;$excel.ActiveWindow.DisplayOutline=$true
    $book.Worksheets.Item($initialSheet).Activate()
    $book.Save()
    [ordered]@{Status='DIRECTED_REVIEW_HELP_PASS';Idempotent=$true;ReopenText=$true;States=$states;Settings=$settings;Navigation=$navigation;Native=$first} |
        ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $directory 'HelpEvidence.json') -Encoding UTF8
    Write-Output 'HELP_REOPEN_OUTLINE_PASS'
} catch {
    [ordered]@{Status='FAIL';Reason=$_.Exception.Message;Stack=$_.ScriptStackTrace;Title=$title;Row=$r;CompletedStates=$states} | ConvertTo-Json -Depth 4 |
        Set-Content -LiteralPath (Join-Path $directory ('Failed_'+(Get-Date -Format 'yyyyMMdd_HHmmss')+'.json')) -Encoding UTF8
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
Restore-WorkbookPrintAreas $path $printAreas
