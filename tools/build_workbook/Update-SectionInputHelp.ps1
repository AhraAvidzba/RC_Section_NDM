# Обновляет краткую справку сечений только в отдельной копии книги.
# Excel нужен для сохранения VBA-событий, проверки компиляции и нативного вида справки.
param([string]$EvidenceDirectory='docs/regression/SectionInputHelp_2026-10-22')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
$directory=[IO.Path]::GetFullPath((Join-Path $root $EvidenceDirectory))
$allowed=(Join-Path $root 'docs/regression')+[IO.Path]::DirectorySeparatorChar
if (-not $directory.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)) {throw 'Evidence must stay in docs/regression.'}
if (Test-Path -LiteralPath $directory) {throw 'Use a new evidence directory.'}
New-Item -ItemType Directory -Path $directory | Out-Null
$source=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$baseline=Join-Path $directory 'Baseline.xlsm'
$candidate=Join-Path $directory 'NativeCandidate.xlsm'
$sourceHash=(Get-FileHash -LiteralPath $source).Hash
Copy-Item -LiteralPath $source -Destination $baseline
Copy-Item -LiteralPath $source -Destination $candidate
Copy-Item -LiteralPath (Join-Path $root 'workbook/output/VBA_All_Code.txt') -Destination (Join-Path $directory 'BaselineVBA.txt')
$keys=@('DependentRebarRows','CircleGeometry','PointedSectionGeometry','HollowRectangleGeometry','RectanglesGeometry')
$definitions=@(Get-SettingsInstructionCatalog | Where-Object {$_.Key -in $keys})
if ($definitions.Count -ne 5) {throw 'Expected common help and four section entries.'}
$definitions | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $directory 'Catalog.json') -Encoding UTF8
$oldTitle='Справка по настройкам листа Config'
$newTitle='Справка по настройкам Config и листа «Сечения»'
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class SectionHelpWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint id);}'

# Копия рисунка переносится в временную книгу; объекты исходной книги не меняются.
function Export-HelpPreview($Book,$Sheet,[int]$First,[int]$Last,[string]$Path) {
    $Sheet.Activate()
    $range=$Sheet.Range($Sheet.Cells.Item($First,1),$Sheet.Cells.Item($Last,6))
    $range.CopyPicture(1,-4147)
    $scratchBook=$Book.Application.Workbooks.Add()
    try {
        $chart=$scratchBook.Worksheets.Item(1).ChartObjects().Add(0,0,[double]$range.Width,[double]$range.Height)
        $chart.Activate();$chart.Chart.Paste()
        if (-not $chart.Chart.Export($Path,'PNG')) {throw 'Help preview export failed.'}
    } finally {$scratchBook.Close($false);$Sheet.Activate()}
}

$excel=$null;$book=$null;$ownedPid=0;$ownedTicks=0
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3
    [uint32]$id=0;[void][SectionHelpWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $ownedPid=[int]$id;$ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    Write-Output "OWNED_EXCEL: id=$ownedPid; utcStartTicks=$ownedTicks"
    $book=$excel.Workbooks.Open($candidate,0,$false)
    if ($book.ReadOnly) {throw 'Private candidate is read-only.'}
    $excel.Calculation=-4135;$excel.CalculateBeforeSave=$false
    $sheet=$book.Worksheets.Item('Справка')
    $anchors=@{};$layout=@();$changes=@()
    foreach ($item in $definitions) {
        $title=if ($item.Key -eq 'DependentRebarRows') {'Смещения дополнительных рядов арматуры'} else {[string]$item.Title}
        $cell=$sheet.Cells.Find($title,$sheet.Cells.Item(1,1),-4163,1,1,1,$false,$false,$false)
        if ($null -eq $cell) {throw "Help block not found: $title"}
        $anchors[$item.Key]=[int]$cell.Row
        for ($i=1;$i -le $item.Lines.Count;$i++) {
            if ([string]$sheet.Cells.Item(([int]$cell.Row+$i),1).Value2 -cne [string]$i) {throw 'Existing paragraph count differs; do not move unrelated help.'}
        }
        $sheet.Rows.Item("$($cell.Row):$([int]$cell.Row+$item.Lines.Count)").Hidden=$false
    }
    foreach ($key in @('DependentRebarRows','HollowRectangleGeometry')) {
        $item=@($definitions | Where-Object Key -eq $key)[0]
        Export-HelpPreview $book $sheet $anchors[$key] ($anchors[$key]+$item.Lines.Count) (Join-Path $directory "Before_$key.png")
    }
    foreach ($item in $definitions) {
        $first=[int]$anchors[$item.Key]
        $heading=$sheet.Cells.Item($first,1)
        $changes+=,[ordered]@{Address=$heading.Address($false,$false);Before=[string]$heading.Value2;After=[string]$item.Title}
        $heading.Value2=[string]$item.Title
        for ($i=0;$i -lt $item.Lines.Count;$i++) {
            $cell=$sheet.Cells.Item(($first+$i+1),2)
            $changes+=,[ordered]@{Address=$cell.Address($false,$false);Before=[string]$cell.Value2;After=[string]$item.Lines[$i]}
            $cell.Value2=[string]$item.Lines[$i]
        }
        $sheet.Rows.Item("$($first+1):$($first+$item.Lines.Count)").AutoFit() | Out-Null
    }
    Add-RelatedInstructionLinks $sheet $anchors
    $heading=$sheet.Cells.Find($oldTitle,$sheet.Cells.Item(1,1),-4163,1,1,1,$false,$false,$false)
    if ($null -eq $heading) {throw 'Main settings heading not found.'}
    $changes+=,[ordered]@{Address=$heading.Address($false,$false);Before=$oldTitle;After=$newTitle}
    $mainHeadingAddress=$heading.Address($false,$false)
    $heading.Value2=$newTitle
    $events=[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'WorkbookHelpEvents.vba'),[Text.Encoding]::UTF8)
    $module=$book.VBProject.VBComponents.Item([string]$book.CodeName).CodeModule
    $beforeEvents=[string]$module.Lines(1,$module.CountOfLines)
    $updatedEvents=$beforeEvents.Replace($oldTitle,$newTitle)
    # Редактор VBA меняет регистр идентификаторов согласно своей таблице имён.
    # В самой книге заменяем только один строковый литерал, сохраняя остальной код.
    if ($updatedEvents.Trim().Replace("`r`n","`n") -ine $events.Trim().Replace("`r`n","`n") -or $updatedEvents -ceq $beforeEvents) {throw 'Workbook events differ beyond the approved title change.'}
    $module.DeleteLines(1,$module.CountOfLines);$module.AddFromString($updatedEvents)
    $excel.VBE.ActiveVBProject=$book.VBProject
    $compile=$excel.VBE.CommandBars.FindControl(1,578)
    if ($null -eq $compile) {throw 'Compile command unavailable.'}
    if ($compile.Enabled) {$compile.Execute()}
    if ($compile.Enabled) {throw 'VBA compilation failed.'}
    foreach ($item in $definitions) {
        $first=[int]$anchors[$item.Key];$heights=@{}
        for ($i=1;$i -le $item.Lines.Count;$i++) {$heights[[string]($first+$i)]=[double]$sheet.Rows.Item(($first+$i)).RowHeight}
        $layout+=,[ordered]@{Key=$item.Key;First=$first;Last=($first+$item.Lines.Count);Heights=$heights;RelatedKey=$item.RelatedKey}
        Export-HelpPreview $book $sheet $first ($first+$item.Lines.Count) (Join-Path $directory "After_$($item.Key).png")
    }
    $book.Save();$book.Close($false);$book=$null
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip=[IO.Compression.ZipFile]::OpenRead($candidate)
    $digest=[Security.Cryptography.SHA256]::Create()
    try {
        $stream=$zip.GetEntry('xl/vbaProject.bin').Open()
        try {$compiledHash=[BitConverter]::ToString($digest.ComputeHash($stream)).Replace('-','')}
        finally {$stream.Dispose()}
    } finally {$digest.Dispose();$zip.Dispose()}
    [ordered]@{SourceSHA256=$sourceHash;Compiled=$true;CompiledVbaSHA256=$compiledHash;OwnedPID=$ownedPid;OwnedStartTicks=$ownedTicks;MainHeadingAddress=$mainHeadingAddress;Blocks=$layout;Changes=$changes} |
        ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $directory 'NativeEdit.json') -Encoding UTF8
    if ((Get-FileHash -LiteralPath $source).Hash -cne $sourceHash) {throw 'Published workbook changed during private editing.'}
} finally {
    if ($book) {$book.Close($false)}
    if ($excel) {$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if ($ownedPid) {
        $owned=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue
        if ($owned -and $owned.ProcessName -eq 'EXCEL' -and $owned.StartTime.ToUniversalTime().Ticks -eq $ownedTicks) {Stop-Process -Id $ownedPid -Force}
    }
}
