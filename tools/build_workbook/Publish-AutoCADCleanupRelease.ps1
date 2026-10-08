# Проверяет узкий выпуск очистки AutoCAD и справки перед заменой рабочей книги.
# Настройки, весь Results, имена и ширины сохраняются; отчет пользователя не трогаем.
param(
    [string]$Directory = 'docs/regression/Performance/CleanupGeometry_2026-10-08',
    [switch]$VerifyOnly,
    [ValidateSet('GeometryOwnership', 'PolylineImport')]
    [string]$ChangeKind = 'GeometryOwnership'
)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$directoryPath = [IO.Path]::GetFullPath((Join-Path $root $Directory))
$allowed = [IO.Path]::GetFullPath((Join-Path $root 'docs/regression/Performance')) + [IO.Path]::DirectorySeparatorChar
if (-not $directoryPath.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase)) { throw 'Evidence must remain inside Performance.' }
$baseline = Join-Path $directoryPath 'Baseline/RC_Section_NDM.xlsm'
$candidate = Join-Path $directoryPath 'Candidate/RC_Section_NDM.xlsm'
$export = Join-Path $directoryPath 'Candidate/VBA_All_Code.txt'
$outputBook = Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$outputExport = Join-Path $root 'workbook/output/VBA_All_Code.txt'
$outputReport = Join-Path $root 'workbook/output/RC_Section_NDM_execution_report.txt'

# Читает структурированное доказательство независимого прогона текущей книги.
function Read-Evidence([string]$Path) {
    $parsed = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
    return $parsed
}

# Не допускает публикацию, если исходник или проверенный артефакт изменился.
function Assert-Hash([string]$Path, [string]$Expected) {
    if ((Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash -cne $Expected) { throw "Artifact changed: $Path" }
}

# Получает устойчивый SHA256 текстового снимка значений, формул или metadata.
function Text-Hash([string]$Text) {
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Text)))).Replace('-', '') }
    finally { $sha.Dispose() }
}

# Сравнивает все значения/формулы пользовательских листов, все именованные
# диапазоны и ширины Results. Изменения допускаются только на листе справки.
function Book-Fingerprint([object]$Book) {
    $sheets = @()
    foreach ($sheet in $Book.Worksheets) {
        if ([string]$sheet.Name -eq 'Справка') { continue }
        $range = $sheet.UsedRange
        $sheets += [ordered]@{Name = [string]$sheet.Name; Address = $range.Address();
            FormulaHash = (Text-Hash ($range.Formula | ConvertTo-Json -Depth 6 -Compress));
            ValueHash = (Text-Hash ($range.Value2 | ConvertTo-Json -Depth 6 -Compress))}
    }
    $names = @(); foreach ($name in $Book.Names) { $names += ([string]$name.Name + '|' + [string]$name.RefersTo) }
    $results = $Book.Worksheets.Item('Results')
    $widths = @(); for ($column = 1; $column -le 260; $column++) { $widths += [double]$results.Columns.Item($column).ColumnWidth }
    return [ordered]@{Sheets = $sheets; Names = (Text-Hash ((@($names | Sort-Object)) -join "`n"));
        ResultsWidths = (Text-Hash ($widths | ConvertTo-Json -Compress))}
}

$preparation = Read-Evidence (Join-Path $directoryPath 'Candidate/Preparation.json')
$candidateHash = (Get-FileHash -LiteralPath $candidate -Algorithm SHA256).Hash
Assert-Hash $baseline $preparation.BaselineSHA256
Assert-Hash $outputBook $preparation.BaselineSHA256
Assert-Hash $candidate $preparation.CandidateSHA256
if (-not $preparation.VBACompileCompleted -or -not $preparation.HelpRefreshed) { throw 'Candidate preparation is incomplete.' }
$equality = Read-Evidence (Join-Path $directoryPath 'Candidate/SourceEquality.json')
$repositoryComponents = @(Get-ChildItem -LiteralPath (Join-Path $root 'src'), (Join-Path $root 'tests') -Recurse -File |
    Where-Object { $_.Extension -in '.bas', '.cls' }).Count
if (-not $equality.AllPassed -or $equality.RepositoryComponents -ne $repositoryComponents -or $equality.ActualComponents -ne ($repositoryComponents + 5) -or
    @($equality.GeneratedDocumentChecks).Count -ne 5) { throw 'Source agreement is incomplete.' }
foreach ($check in $equality.Checks) {
    $files = @(Get-ChildItem -LiteralPath (Join-Path $root 'src'), (Join-Path $root 'tests') -Recurse -File |
        Where-Object { $_.BaseName -eq $check.Component -and $_.Extension -in '.bas', '.cls' })
    if ($files.Count -ne 1) { throw 'Source identity is ambiguous.' }
    Assert-Hash $files[0].FullName $check.SourceSHA256
}
Assert-Hash (Join-Path $PSScriptRoot 'Build-Workbook.ps1') $equality.BuildScriptSHA256
foreach ($folder in @('Candidate', 'FreshCompiled')) {
    $validation = @(Read-Evidence (Join-Path $directoryPath "$folder/Validation.json"))
    if ($validation.Count -ne 27 -or @($validation | Where-Object { -not $_.Passed }).Count) { throw "Validation failed: $folder" }
}
foreach ($gate in @(
    @{Path = 'Focused.txt'; Totals = @('TOTAL_AUTOCAD_CONFIG', 'TOTAL_AUTOCAD_CONTOURS', 'TOTAL_POSTAUDIT03_CONTOURS', 'TOTAL_CONFIG_PRESENTATION', 'TOTAL:', 'TOTAL_GENERATED_CONTOURS', 'TOTAL_SP35_PREPARATION', 'TOTAL_CRACK_GEOMETRY')},
    @{Path = 'GeometryTests.txt'; Totals = @('TOTAL:', 'TOTAL_HOLLOW_CELL_MESH')},
    @{Path = 'FreshConfigPresentation.txt'; Totals = @('TOTAL_CONFIG_PRESENTATION')}
)) {
    $text = Get-Content -LiteralPath (Join-Path $directoryPath $gate.Path) -Raw
    if ($text -match '(?m)^FAIL:|^SCRIPT ERROR|^RUNTIME ERROR' -or $text -notmatch 'SOURCE_UNCHANGED: True') { throw "Gate failed: $($gate.Path)" }
    foreach ($total in $gate.Totals) {
        if ($text -notmatch "(?m)^${total}[^\r\n]*passed=[1-9][0-9]*; failed=0") { throw "Gate missing: $total" }
    }
}
$native = Get-Content -LiteralPath (Join-Path $directoryPath 'Native/NativeCAD.txt') -Raw
$identity = Read-Evidence (Join-Path $directoryPath 'Native/NativeCADIdentity.json')
if ($native -match '(?m)^FAIL:' -or $native -notmatch 'TOTAL_REAL_AUTOCAD_CONTOURS: passed=[1-9][0-9]*; failed=0' -or
    $identity.UserDocumentsUsed -or $identity.Executable -cne 'C:\Program Files\Autodesk\AutoCAD 2023\acad.exe') { throw 'Native CAD gate failed.' }
Assert-Hash (Join-Path $directoryPath 'Native/RC_Section_NDM.xlsm') $candidateHash
$reportHash = (Get-FileHash -LiteralPath $outputReport -Algorithm SHA256).Hash
$oldExportHash = (Get-FileHash -LiteralPath $outputExport -Algorithm SHA256).Hash
$exportHash = (Get-FileHash -LiteralPath $export -Algorithm SHA256).Hash
$excel = $null; $book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false; $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($baseline, 0, $true)
    $before = Book-Fingerprint $book
    $book.Close($false); $book = $null
    $book = $excel.Workbooks.Open($candidate, 0, $true)
    $after = Book-Fingerprint $book
    if (($before | ConvertTo-Json -Depth 8 -Compress) -cne ($after | ConvertTo-Json -Depth 8 -Compress)) { throw 'User data, named ranges or widths changed.' }
    $guide = $book.Worksheets.Item('Справка')
    $used = $guide.UsedRange; $values = $used.Value2; $startRow = 0; $endRow = 0
    for ($row = 1; $row -le $values.GetLength(0); $row++) {
        if ([string]$values[$row, 1] -ceq '2. Точный контур и несколько отверстий') { $startRow = $used.Row + $row - 1 }
        if ([string]$values[$row, 1] -ceq '3. Единицы, поворот и плоскость') { $endRow = $used.Row + $row - 1 }
    }
    if ($startRow -le 0 -or $endRow -le $startRow) { throw 'Geometry guide section is missing.' }
    $guide.Outline.ShowLevels(8)
    $range = $guide.Range($guide.Cells.Item($startRow, 1), $guide.Cells.Item($endRow - 1, 6))
    $guide.PageSetup.PrintArea = $range.Address()
    $guide.PageSetup.Orientation = 2; $guide.PageSetup.PaperSize = 8
    $guide.PageSetup.Zoom = $false; $guide.PageSetup.FitToPagesWide = 1; $guide.PageSetup.FitToPagesTall = 1
    $guide.ExportAsFixedFormat(0, (Join-Path $directoryPath 'GeometryGuide.pdf'), 0, $true, $false)
    $book.Close($false); $book = $null
} finally {
    if ($book) { try { $book.Close($false) } catch {} }
    if ($excel) { try { $excel.Quit() } catch {}; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
Assert-Hash $baseline $preparation.BaselineSHA256
Assert-Hash $outputBook $preparation.BaselineSHA256
Assert-Hash $candidate $candidateHash
Assert-Hash $outputExport $oldExportHash
Assert-Hash $outputReport $reportHash
$accepted = [ordered]@{Passed = $true; WorkbookSHA256 = $candidateHash; ExportSHA256 = $exportHash;
    BaselineWorkbookSHA256 = $preparation.BaselineSHA256; Fingerprint = $after;
    UserReportSHA256 = $reportHash; UserReportPreserved = $true; UserDataNamesWidthsPreserved = $true;
    NativeCADPassed = $true; FullOnOffRepeated = $false; CrackDomainContractChanged = ($ChangeKind -eq 'GeometryOwnership');
    GeneratedContourOwnershipChanged = ($ChangeKind -eq 'GeometryOwnership'); MultipleOpeningsFromGeometryPassed = $true;
    PolylineImportChanged = ($ChangeKind -eq 'PolylineImport');
    ImportSummaryChanged = ($ChangeKind -eq 'PolylineImport'); LockedContourGuardChanged = ($ChangeKind -eq 'PolylineImport');
    SolverSearchMaterialsChanged = $false}
$accepted | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $directoryPath 'Acceptance.json') -Encoding UTF8
if ($VerifyOnly) { Write-Output 'ACCEPTED: current code, native CAD and help gates passed; user state preserved.'; exit }
Copy-Item -LiteralPath $candidate -Destination $outputBook
Copy-Item -LiteralPath $export -Destination $outputExport
Assert-Hash $outputBook $candidateHash
Assert-Hash $outputExport $exportHash
Assert-Hash $outputReport $reportHash
[ordered]@{Published = $true; UTC = [DateTime]::UtcNow.ToString('o'); Acceptance = $accepted} |
    ConvertTo-Json -Depth 9 | Set-Content -LiteralPath (Join-Path $directoryPath 'Publication.json') -Encoding UTF8
Write-Output "PUBLISHED: book=$candidateHash; VBA=$exportHash; user state and report unchanged."
