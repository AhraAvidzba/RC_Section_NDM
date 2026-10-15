# ALL CODE становится точным фактическим экспортом основной книги после readback.
$ErrorActionPreference='Stop'
$root=(Get-Location).Path
$release=Join-Path $root 'docs/regression/Performance/Overload_2026-10-15/Release'
$evidence=Join-Path $root 'docs/regression/ExcelGuardGuide_2026-10-13/Overload_2026-10-15'
$publication=Get-Content -LiteralPath (Join-Path $release 'Publication.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$equality=Get-Content -LiteralPath (Join-Path $evidence 'PublishedSourceEquality.json') -Raw -Encoding UTF8 | ConvertFrom-Json
if(-not $equality.Passed){throw 'Actual source equality failed.'}
$actualPath=Join-Path $evidence 'PublishedNativeVBA.txt'
$actual=[IO.File]::ReadAllText($actualPath,[Text.Encoding]::UTF8)
$expected=[IO.File]::ReadAllText((Join-Path $release 'CandidateNativeVBA.txt'),[Text.Encoding]::UTF8)
$stripDate='(?m)^Date: [^\r\n]*\r?\n'
if([regex]::Replace($actual,$stripDate,'') -cne [regex]::Replace($expected,$stripDate,'')){throw 'Published VBA differs from accepted candidate.'}
if((Get-FileHash -LiteralPath (Join-Path $root 'workbook/output/RC_Section_NDM.xlsm')).Hash -cne $publication.WorkbookSHA256){throw 'Main workbook changed during readback.'}
$code=Join-Path $root 'workbook/output/VBA_All_Code.txt'
[byte[]]$bytes=[IO.File]::ReadAllBytes($actualPath)
$stream=[IO.File]::Open($code,'Open','ReadWrite','None');$digest=[Security.Cryptography.SHA256]::Create()
try{
    $hash=[BitConverter]::ToString($digest.ComputeHash($stream)).Replace('-','')
    if($hash -cne $publication.AllCodeSHA256){throw 'ALL CODE was edited during readback.'}
    $stream.Position=0;$stream.Write($bytes,0,$bytes.Length);$stream.SetLength($bytes.Length);$stream.Flush($true)
}finally{$stream.Dispose();$digest.Dispose()}
$publication.AllCodeSHA256=(Get-FileHash -LiteralPath $code).Hash
$publication.ActualWorkbookReadbackPending=$false
$publication | Add-Member -NotePropertyName AllCodeMatchesActualNativeExport -NotePropertyValue $true
$publication | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $release 'Publication.json') -Encoding UTF8
Write-Output 'OVERLOAD_ACTUAL_OUTPUT_READBACK_PASS'
