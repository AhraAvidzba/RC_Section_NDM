# Сохраняет последнюю пользовательскую книгу; не закрывает чужие книги или Excel.
$ErrorActionPreference='Stop'
$root=(Get-Location).Path
$directory=Join-Path $root 'docs/regression/Performance/Overload_2026-10-15/Release'
New-Item -ItemType Directory -Path $directory -Force | Out-Null
$path=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$baseline=Join-Path $directory 'Baseline.xlsm'
if(Test-Path -LiteralPath $baseline){throw 'Release baseline already exists.'}
$lock=$null
try{$lock=[IO.File]::Open($path,'Open','Read','None')}catch [IO.IOException]{}
$savedOpenBook=$false
if($lock){$lock.Dispose()}else{
    $book=[Runtime.InteropServices.Marshal]::BindToMoniker($path)
    try{
        if([string]$book.FullName -cne $path -or $book.ReadOnly -or -not $book.Application.Ready){throw 'Unexpected or unavailable open workbook.'}
        if(-not $book.Saved){$book.Save();$savedOpenBook=$true}
        $book.Close($false)
    }finally{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($book)}
}
Copy-Item -LiteralPath $path -Destination $baseline
[ordered]@{
    CapturedAt=(Get-Date).ToString('o');SavedOpenWorkbook=$savedOpenBook
    WorkbookSHA256=(Get-FileHash -LiteralPath $path).Hash
    AllCodeSHA256=(Get-FileHash -LiteralPath (Join-Path $root 'workbook/output/VBA_All_Code.txt')).Hash
    ExecutionReportSHA256=(Get-FileHash -LiteralPath (Join-Path $root 'workbook/output/RC_Section_NDM_execution_report.txt')).Hash
} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'Before.json') -Encoding UTF8
Write-Output 'ACTUAL_USER_RELEASE_BASELINE_SAVED'
