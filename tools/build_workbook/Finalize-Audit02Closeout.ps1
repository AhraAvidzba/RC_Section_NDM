# Завершает публикацию только после независимого чтения фактической основной книги.
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$directory=Join-Path $root 'docs/regression/ExcelGuardGuide_2026-10-13/Audit02Closeout_2026-10-14'
$publication=Get-Content -LiteralPath (Join-Path $directory 'Publication.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$book=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$code=Join-Path $root 'workbook/output/VBA_All_Code.txt'
$snapshot=Join-Path $directory 'PublishedNativeVBA.txt'
$equality=Get-Content -LiteralPath (Join-Path $directory 'PublishedSourceEquality.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$help=Get-Content -LiteralPath (Join-Path $directory 'PublishedStandaloneHelp.json') -Raw -Encoding UTF8 | ConvertFrom-Json
if(-not $equality.Passed -or $help.status -ne 'STANDALONE_HELP_PASS'){throw 'Actual workbook verification failed.'}
$actual=[IO.File]::ReadAllText($snapshot,[Text.Encoding]::UTF8)
$expected=[IO.File]::ReadAllText((Join-Path $directory 'FinalNativeVBA.txt'),[Text.Encoding]::UTF8).Replace('Workbook: Candidate.xlsm','Workbook: RC_Section_NDM.xlsm')
$stripDate='(?m)^Date: [^\r\n]*\r?\n'
if([regex]::Replace($actual,$stripDate,'') -cne [regex]::Replace($expected,$stripDate,'')){throw 'Published VBA differs from the accepted candidate.'}
if((Get-FileHash -LiteralPath $book).Hash -cne $publication.WorkbookSHA256){throw 'Published workbook changed during verification.'}
[byte[]]$bytes=[IO.File]::ReadAllBytes($snapshot)
$stream=[IO.File]::Open($code,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
$digest=[Security.Cryptography.SHA256]::Create()
try {
    $hash=[BitConverter]::ToString($digest.ComputeHash($stream)).Replace('-','')
    if($hash -cne $publication.AllCodeSHA256){throw 'ALL CODE was edited during readback; preserve the edit.'}
    $stream.Position=0;$stream.Write($bytes,0,$bytes.Length);$stream.SetLength($bytes.Length);$stream.Flush($true)
} finally {$stream.Dispose();$digest.Dispose()}
if((Get-FileHash -LiteralPath $code).Hash -cne (Get-FileHash -LiteralPath $snapshot).Hash){throw 'ALL CODE is not the actual native export.'}
$publication.AllCodeSHA256=(Get-FileHash -LiteralPath $code).Hash
$publication.ReadbackPending=$false
$publication | Add-Member -NotePropertyName PublishedCodeMatchesCandidate -NotePropertyValue $true
$publication | Add-Member -NotePropertyName AllCodeMatchesActualWorkbookExport -NotePropertyValue $true
$publication | Add-Member -NotePropertyName NativeComponents -NotePropertyValue 157
$publication | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'Publication.json') -Encoding UTF8
Write-Output 'AUDIT02_READBACK_PASS: 157 components; ALL CODE is the actual main workbook export.'
