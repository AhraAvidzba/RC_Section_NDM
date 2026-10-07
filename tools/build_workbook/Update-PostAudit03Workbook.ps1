# Готовит кандидат выпуска из сохраненной пользовательской baseline-книги.
# Основной output, Config и пользовательский execution report не перезаписываются.
param([string]$ReportDirectory = 'docs/regression/PostAudit03/ReleaseCandidate',
      [switch]$ReuseCandidate, [string[]]$ComponentNames = @())
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$directory = [IO.Path]::GetFullPath((Join-Path $root $ReportDirectory))
$allowed = [IO.Path]::GetFullPath((Join-Path $root 'docs/regression/PostAudit03')) + [IO.Path]::DirectorySeparatorChar
if (-not $directory.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase)) {throw 'Candidate must remain inside PostAudit03 evidence.'}
New-Item -ItemType Directory -Path $directory -Force | Out-Null
$path = Join-Path $directory 'RC_Section_NDM.xlsm'
if (-not $ReuseCandidate) {
    Copy-Item -LiteralPath (Join-Path $root 'docs/regression/PostAudit03/Baseline/RC_Section_NDM.xlsm') -Destination $path -Force
}
$ComponentNames = @($ComponentNames | ForEach-Object {$_ -split ','})
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
$printAreas = @(Get-WorkbookPrintAreas $path)
$excel = $null; $book = $null
function Get-FormulaSignature([object]$Range) {
    $payload = ConvertTo-Json -InputObject @($Range.Formula) -Depth 8 -Compress
    $sha = [Security.Cryptography.SHA256]::Create()
    try {return [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($payload))).Replace('-', '')} finally {$sha.Dispose()}
}
# Читает тот же COM-объект через IDispatch без повторного открытия или подстановки.
# Недоступное свойство остается ошибкой и не маскирует потерю книги.
function Get-RequiredComProperty([object]$Target, [string]$Property) {
    if ($null -eq $Target) {throw "COM target unavailable: $Property"}
    $value=$Target.GetType().InvokeMember($Property,[Reflection.BindingFlags]::GetProperty,$null,$Target,$null)
    if ($null -eq $value) {throw "COM property unavailable: $Property"}
    return ,$value
}
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false; $excel.AutomationSecurity=1
    $books=Get-RequiredComProperty $excel 'Workbooks'
    $book=$books.Open($path)
    $sheets=Get-RequiredComProperty $book 'Worksheets'
    $config=$sheets.Item('Config').UsedRange
    $configAddress=$config.Address(); $beforeConfig=Get-FormulaSignature $config
    $results=$sheets.Item('Results'); $widths=@()
    for ($c=1;$c -le 100;$c++) {$widths += [double]$results.Columns.Item($c).ColumnWidth}
    $project=$book.GetType().InvokeMember('VBProject',[Reflection.BindingFlags]::GetProperty,$null,$book,$null)
    if ($null -eq $project) {throw 'Workbook VBProject is unavailable.'}
    $components=$project.GetType().InvokeMember('VBComponents',[Reflection.BindingFlags]::GetProperty,$null,$project,$null)
    if ($null -eq $components) {throw 'Workbook VBComponents is unavailable.'}
    foreach ($file in Get-ChildItem -LiteralPath (Join-Path $root 'src'), (Join-Path $root 'tests') -Recurse -File | Where-Object Extension -in '.bas','.cls') {
        if ($ComponentNames.Count -gt 0 -and $ComponentNames -notcontains $file.BaseName) {continue}
        Write-Output ('IMPORT: ' + $file.BaseName)
        $component=$null
        for ($index=1;$index -le $components.Count;$index++) {
            $existing=$components.Item($index)
            if ([string]$existing.Name -eq $file.BaseName) {$component=$existing; break}
        }
        $body=[regex]::Match([IO.File]::ReadAllText($file.FullName,[Text.Encoding]::UTF8),'(?ms)^Option Explicit.*').Value
        if (-not $body) {throw "Missing VBA body: $($file.FullName)"}
        $type=1; if ($file.Extension -eq '.cls') {$type=2}
        if ($null -ne $component) {
            if ([int]$component.Type -ne $type) {throw "VBA component type differs: $($file.BaseName)"}
            if ($component.CodeModule.CountOfLines -gt 0) {$component.CodeModule.DeleteLines(1,$component.CodeModule.CountOfLines)}
        } else {
            $component=$components.Add($type)
            try {$component.Name=$file.BaseName} catch {throw "VBA component creation failed: $($file.BaseName); $($_.Exception.Message)"}
        }
        $component.CodeModule.AddFromString(($body -split "`r?`n") -join "`r`n")
    }
    Write-Output 'MIGRATION: start'
    $migration=[string]$excel.Run("'$($book.Name)'!MigrateSavedSectionContours")
    $repeat=[string]$excel.Run("'$($book.Name)'!MigrateSavedSectionContours")
    if ($repeat -notmatch 'already current') {throw 'Migration not idempotent.'}
    $sheets=Get-RequiredComProperty $book 'Worksheets'
    # Правит только сохраненную шапку: Es остается отдельным параметром,
    # объединение ширины начинается с acrc. Строки расчетных данных не меняются.
    $results=$sheets.Item('Results')
    $anchor=$results.Range('rngCrackSummaryAnchor')
    $headerRefreshed=$false
    if ([string]$anchor.Offset(-1,62).Value2 -like 'Es,*' -and
        [string]$anchor.Offset(-1,71).Value2 -eq 'статус') {
        $dataSignature=Get-FormulaSignature $anchor.Resize(30,72)
        $heading=$anchor.Offset(-3,62).Resize(1,5)
        $heading.UnMerge(); $heading.ClearContents()
        $heading.Borders.LineStyle=1; $heading.Borders.Weight=2
        $widthHeading=$anchor.Offset(-3,63).Resize(1,4)
        $widthHeading.Merge()
        $widthHeading.Cells.Item(1,1).Value2='ширина раскрытия нормальных трещин'
        if ((Get-FormulaSignature $anchor.Resize(30,72)) -ne $dataSignature) {throw 'Crack data changed during header formatting.'}
        $headerRefreshed=$true
    }
    if ((Get-FormulaSignature $sheets.Item('Config').Range($configAddress)) -ne $beforeConfig) {throw 'Config formula/value changed during migration.'}
    for ($c=1;$c -le 100;$c++) {if ([double]$results.Columns.Item($c).ColumnWidth -ne $widths[$c-1]) {throw "Results column width changed: $c"}}
    $book.Save(); $book.Close($false); $book=$null
    $books=Get-RequiredComProperty $excel 'Workbooks'
    $book=$books.Open($path)
    $sheets=Get-RequiredComProperty $book 'Worksheets'
    if ((Get-FormulaSignature $sheets.Item('Config').Range($configAddress)) -ne $beforeConfig) {throw 'Config differs after reopen.'}
    $reopen=[string]$excel.Run("'$($book.Name)'!MigrateSavedSectionContours")
    if ($reopen -notmatch 'already current') {throw 'Reopened snapshot not current.'}
    [ordered]@{Migration=$migration; Repeat=$repeat; Reopen=$reopen; ConfigPreserved=$true; ResultsWidthsPreserved=$true; CrackHeaderRefreshed=$headerRefreshed; CrackDataPreserved=$true; SourceGitSHA=(& git -C $root rev-parse HEAD); SolveExecuted=$false; ImportExecuted=$false} |
        ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'Migration.json') -Encoding UTF8
    Write-Output $migration
} finally {
    if ($book) {try {$book.Close($false)} catch {}; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}
    if ($excel) {try {$excel.Quit()} catch {}; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
Restore-WorkbookPrintAreas $path $printAreas
Write-Output "CANDIDATE: $path; Config=True; widths=True; no solve/import"
