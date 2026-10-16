# Сохраняет последние пользовательские данные перед точечным обновлением VBA.
# Закрывает только основную книгу NDM; остальные книги и процессы Excel не меняет.
$ErrorActionPreference='Stop'
$root=(Get-Location).Path;$d=$PSScriptRoot
$path=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$baseline=Join-Path $d 'UserBaseline.xlsm'
if(Test-Path -LiteralPath $baseline){throw 'Latest user snapshot already exists.'}
$lock=$null
try{$lock=[IO.File]::Open($path,'Open','Read','None')}catch [IO.IOException]{}
$saved=$false;$closed=$false
if($lock){$lock.Dispose()}else{
    $book=[Runtime.InteropServices.Marshal]::BindToMoniker($path)
    try{
        if([string]$book.FullName -cne $path -or $book.ReadOnly -or -not $book.Application.Ready){throw 'Unexpected or unavailable open workbook.'}
        if(-not $book.Saved){$book.Save();$saved=$true}
        $book.Close($false);$closed=$true
    }finally{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($book)}
}
Copy-Item -LiteralPath $path -Destination $baseline
$hashes=@{}
foreach($relative in @('workbook/output/RC_Section_NDM.xlsm','workbook/output/VBA_All_Code.txt','workbook/output/RC_Section_NDM_execution_report.txt')){
    $hashes[$relative]=(Get-FileHash -LiteralPath (Join-Path $root $relative)).Hash
}
@{Hashes=$hashes;SavedUnsavedData=$saved;ClosedMainWorkbook=$closed;CapturedAt=(Get-Date).ToString('o')} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $d 'UserSnapshot.json') -Encoding UTF8
Write-Output 'LATEST_USER_BOOK_CAPTURED'
