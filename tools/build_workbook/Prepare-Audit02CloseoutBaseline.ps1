# Сохраняет фактическую книгу и код до доработки; пользовательские данные не заменяются.
$ErrorActionPreference='Stop'
$root=(Get-Location).Path
$relative='docs/regression/ExcelGuardGuide_2026-10-13/Audit02Closeout_2026-10-14'
$directory=Join-Path $root $relative
New-Item -ItemType Directory -Path $directory -Force | Out-Null
$production=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$beforePath=Join-Path $directory 'Before.json'
if(Test-Path -LiteralPath $beforePath){throw 'Baseline already recorded; do not replace.'}
$lock=$null
try {$lock=[IO.File]::Open($production,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::None)} catch [IO.IOException] {}
if($lock){$lock.Dispose()} else {
    # Пользователь ранее разрешил сохранить и закрыть именно основную книгу.
    $book=[Runtime.InteropServices.Marshal]::BindToMoniker($production)
    try {
        if([string]$book.FullName -cne $production -or -not $book.Application.Ready -or $book.ReadOnly){throw 'Leave the unavailable or unexpected workbook alone.'}
        [ordered]@{FullName=[string]$book.FullName;Saved=[bool]$book.Saved;OtherWorkbooks=$book.Application.Workbooks.Count-1} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'LiveStateBeforeClose.json') -Encoding UTF8
        if(-not $book.Saved){$book.Save()}
        $book.Close($false)
    } finally {[void][Runtime.InteropServices.Marshal]::ReleaseComObject($book)}
}
foreach($name in @('Baseline.xlsm','Candidate.xlsm')){Copy-Item -LiteralPath $production -Destination (Join-Path $directory $name)}
$sourceArchive=Join-Path $directory 'BaselineSource.zip'
& git archive --format=zip --output=$sourceArchive HEAD src tests tools/build_workbook docs/geometry_refactor/examples
if($LASTEXITCODE -ne 0){throw 'Read-only source archive failed.'}
[ordered]@{Head=([string](git rev-parse HEAD));WorkbookSHA256=(Get-FileHash -LiteralPath $production).Hash;AllCodeSHA256=(Get-FileHash -LiteralPath 'workbook/output/VBA_All_Code.txt').Hash;ExecutionReportSHA256=(Get-FileHash -LiteralPath 'workbook/output/RC_Section_NDM_execution_report.txt').Hash;SourceArchiveSHA256=(Get-FileHash -LiteralPath $sourceArchive).Hash;AuditSpecSHA256=(Get-FileHash -LiteralPath 'docs/geometry_refactor/аудит_02/NDM_Recipe_API_Audit02.md').Hash;PreparedAt=(Get-Date).ToString('o')} | ConvertTo-Json | Set-Content -LiteralPath $beforePath -Encoding UTF8
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File tools/build_workbook/Export-Audit03VbaSnapshot.ps1 -WorkbookPath "$relative/Baseline.xlsm" -OutputPath "$relative/BaselineNativeVBA.txt"
if($LASTEXITCODE -ne 0){throw 'Baseline native export failed.'}
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File tools/build_workbook/Verify-PostAudit03VbaEquality.ps1 -ExportPath "$relative/BaselineNativeVBA.txt" -ReportPath "$relative/BaselineSourceEquality.json"
if($LASTEXITCODE -ne 0){throw 'Actual baseline code differs from current sources.'}
if(-not (Get-Content -LiteralPath (Join-Path $directory 'BaselineSourceEquality.json') -Raw -Encoding UTF8 | ConvertFrom-Json).Passed){throw 'Baseline differs from sources.'}
Write-Output 'ACTUAL_BASELINE_READY: main workbook, source archive and native VBA recorded.'
