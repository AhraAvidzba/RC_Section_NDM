# Исправляет только оформление двух блоков Справки и главного заголовка на копии книги.
param([string]$ReportDirectory='docs/regression/HelpPresentation_2026-10-13',
      [string]$SourceWorkbookPath='workbook/output/RC_Section_NDM.xlsm')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
$source=(Resolve-Path -LiteralPath (Join-Path $root $SourceWorkbookPath)).Path
$directory=[IO.Path]::GetFullPath((Join-Path $root $ReportDirectory))
$allowed=(Join-Path $root 'docs/regression')+[IO.Path]::DirectorySeparatorChar
if (-not $directory.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)) {throw 'Evidence directory must stay inside docs/regression.'}
if (Test-Path -LiteralPath $directory) {throw 'Existing evidence must not be overwritten.'}
New-Item -ItemType Directory -Path $directory | Out-Null
$sourceHash=(Get-FileHash -LiteralPath $source).Hash
$candidate=Join-Path $directory 'RC_Section_NDM.xlsm'
Copy-Item -LiteralPath $source -Destination (Join-Path $directory 'Before.xlsm')
Copy-Item -LiteralPath $source -Destination $candidate
$printAreas=@(Get-WorkbookPrintAreas $candidate)
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class HelpPresentationWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hWnd,out uint id);}'

function Export-HelpPicture($Book,$Sheet,[int]$First,[int]$Last,[string]$Path) {
    $Sheet.Activate()
    $range=$Sheet.Range($Sheet.Cells.Item($First,1),$Sheet.Cells.Item($Last,6))
    $range.CopyPicture(1,-4147)
    # Превью в отдельной книге не меняет состав листов и бинарный VBA-проект кандидата.
    $scratchBook=$Book.Application.Workbooks.Add()
    $scratch=$scratchBook.Worksheets.Item(1)
    try {
        $chart=$scratch.ChartObjects().Add(0,0,[double]$range.Width,[double]$range.Height)
        $chart.Activate();$chart.Chart.Paste()
        if (-not $chart.Chart.Export($Path,'PNG')) {throw 'Native Help image export failed.'}
    } finally {$scratchBook.Close($false);$Sheet.Activate()}
}

$excel=$null;$book=$null;$ownedPid=0;$ownedTicks=0
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3
    [uint32]$excelId=0;[void][HelpPresentationWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$excelId)
    $ownedPid=[int]$excelId;$ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    Write-Output ('OWNED_EXCEL: id='+$ownedPid+'; utcStartTicks='+$ownedTicks)
    $book=$excel.Workbooks.Open($candidate,0,$false)
    if ($book.ReadOnly) {throw 'Candidate is read-only.'}
    $excel.CalculateBeforeSave=$false
    $initialSheet=[string]$book.ActiveSheet.Name
    $sheet=$book.Worksheets.Item('Справка')
    $blocks=@()
    foreach ($key in @('SLS.Crack.ToleranceLambda','General.PlateauSlopeFactor')) {
        $anchor=$sheet.Cells.Find($key,$sheet.Cells.Item(1,1),-4163,1,1,1,$false,$false,$false)
        if ($null -eq $anchor) {throw ('Missing Help entry: '+$key)}
        $first=[int]$anchor.Row;$last=$first
        while ([string]$sheet.Cells.Item($last+1,1).Value2 -ne '' -and -not $sheet.Cells.Item($last+1,1).Font.Bold) {$last++}
        $textColumn=if ([string]$sheet.Cells.Item($first+1,2).Value2 -ne '') {2} else {1}
        $oldLines=@(for ($r=$first+1;$r -le $last;$r++) {[string]$sheet.Cells.Item($r,$textColumn).Value2})
        $definition=@(Get-SettingsInstructionCatalog | Where-Object Key -CEQ $key)
        if ($definition.Count -ne 1) {throw ('Ambiguous Help definition: '+$key)}
        $lines=@($definition[0].Lines)
        # В книге могла быть согласована более поздняя редакция текста.
        # Объединяем прежние переносы только при дословном совпадении с каталогом.
        if (($oldLines -join ' ') -cne ($lines -join ' ')) {$lines=$oldLines}
        $blocks+=,[ordered]@{Key=$key;First=$first;Last=$last;Lines=$lines}
    }
    $recipe=$sheet.Cells.Find('Добавление нового типа сечения',$sheet.Cells.Item(1,1),-4163,1)
    $reference=$sheet.Cells.Find('Справка по результатам расчета листа Results',$sheet.Cells.Item(1,1),-4163,1)
    $config=$sheet.Cells.Find('Справка по настройкам Config и листа «Сечения»',$sheet.Cells.Item(1,1),-4163,1)
    $recipeRow=[int]$recipe.Row;$referenceRow=[int]$reference.Row;$configRow=[int]$config.Row
    Export-HelpPicture $book $sheet ($blocks[0].First-6) $blocks[1].Last (Join-Path $directory 'BeforeSettings.png')
    Export-HelpPicture $book $sheet $referenceRow $configRow (Join-Path $directory 'BeforeSections.png')
    foreach ($block in $blocks) {
        Set-SettingInstructionBlock $sheet $block.First $block.Key $block.Lines $block.Last | Out-Null
    }
    $sheet.Range($sheet.Cells.Item($recipeRow,1),$sheet.Cells.Item($recipeRow,6)).Interior.Color=$reference.Interior.Color
    $book.Worksheets.Item($initialSheet).Activate()
    $book.Save();$book.Close($false);$book=$null
    Restore-WorkbookPrintAreas $candidate $printAreas

    # Excel при Save может перестроить служебные потоки VBA даже без правки кода.
    # Здесь меняется только Справка, состав/CodeName листов остаётся прежним:
    # сохраняем весь исходный VBA-проект, включая его скомпилированные потоки.
    $originalZip=[IO.Compression.ZipFile]::OpenRead($source)
    $vbaBytes=New-Object IO.MemoryStream
    try {
        $entryStream=$originalZip.GetEntry('xl/vbaProject.bin').Open()
        try {$entryStream.CopyTo($vbaBytes)} finally {$entryStream.Dispose()}
    } finally {$originalZip.Dispose()}
    $candidateZip=[IO.Compression.ZipFile]::Open($candidate,[IO.Compression.ZipArchiveMode]::Update)
    try {
        $candidateZip.GetEntry('xl/vbaProject.bin').Delete()
        $entryStream=$candidateZip.CreateEntry('xl/vbaProject.bin').Open()
        try {$vbaBytes.Position=0;$vbaBytes.CopyTo($entryStream)} finally {$entryStream.Dispose()}
    } finally {$candidateZip.Dispose();$vbaBytes.Dispose()}

    # Проверяется уже сохранённая книга; изображения создаются без её повторного сохранения.
    $book=$excel.Workbooks.Open($candidate,0,$true)
    $sheet=$book.Worksheets.Item('Справка')
    $checks=@()
    foreach ($block in $blocks) {
        $anchor=$sheet.Cells.Item([int]$block.First,1)
        if ($anchor.Font.Size -ne 10 -or -not $anchor.Font.Bold -or $anchor.Interior.Color -ne 15921906) {throw 'Setting heading style differs.'}
        if ($anchor.EntireRow.RowHeight -gt 15) {throw 'Oversized setting heading.'}
        if ($anchor.MergeArea.Borders.Item(8).Color -ne 14277081 -or $anchor.MergeArea.Borders.Item(9).Color -ne 14277081) {throw 'Unexpected dark heading frame.'}
        $heights=@()
        for ($i=0;$i -lt $block.Lines.Count;$i++) {
            $r=[int]$block.First+$i+1;$cell=$sheet.Cells.Item($r,2)
            if ([double]$sheet.Cells.Item($r,1).Value2 -ne ($i+1) -or [string]$cell.Value2 -cne $block.Lines[$i]) {throw 'Numbered Help text mismatch.'}
            if ($cell.MergeCells -or $cell.Font.Size -ne 10 -or -not $cell.WrapText) {throw 'Body style differs.'}
            if ($cell.Borders.Item(8).Color -ne 14277081 -or $cell.Borders.Item(9).Color -ne 14277081) {throw 'Unexpected dark row frame.'}
            $heights+=[double]$cell.EntireRow.RowHeight
        }
        $checks+=,[ordered]@{Key=$block.Key;First=$block.First;Last=$block.Last;Paragraphs=$block.Lines.Count;RowHeights=$heights;Pass=$true}
    }
    $fill=[double]$sheet.Cells.Item($referenceRow,1).Interior.Color
    if ($sheet.Cells.Item($recipeRow,1).Interior.Color -ne $fill -or $sheet.Cells.Item($configRow,1).Interior.Color -ne $fill) {throw 'Main section fills differ.'}
    Export-HelpPicture $book $sheet ($blocks[0].First-6) ($blocks[1].First+$blocks[1].Lines.Count) (Join-Path $directory 'AfterSettings.png')
    Export-HelpPicture $book $sheet $referenceRow $configRow (Join-Path $directory 'AfterSections.png')
    $book.Close($false);$book=$null
    if ((Get-FileHash -LiteralPath $source).Hash -cne $sourceHash) {throw 'Source changed during preparation.'}
    [ordered]@{AllPassed=$true;SourceSHA256=$sourceHash;CandidateSHA256=(Get-FileHash -LiteralPath $candidate).Hash;Blocks=$checks;HeaderRow=$recipeRow;HeaderFill=$fill;VisualReview='Pending'} |
        ConvertTo-Json -Depth 7 | Set-Content -LiteralPath (Join-Path $directory 'NativeChecks.json') -Encoding UTF8
    Write-Output ('CANDIDATE_READY: '+$candidate)
} finally {
    if ($null -ne $book) {$book.Close($false)}
    if ($null -ne $excel) {$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if ($ownedPid) {
        $owned=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue
        if ($null -ne $owned -and $owned.StartTime.ToUniversalTime().Ticks -eq $ownedTicks) {Stop-Process -Id $ownedPid -Force}
    }
}
