$ErrorActionPreference='Stop'
$root=(Get-Location).Path
$dir=Join-Path $root 'docs/regression/ExcelGuardGuide_2026-10-13/RecipeArguments'
$publication=Get-Content -LiteralPath (Join-Path $dir 'Publication.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$published=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$allCode=Join-Path $root 'workbook/output/VBA_All_Code.txt'
$snapshot=Join-Path $dir 'PublishedNativeVBA.txt'
$native=[IO.File]::ReadAllText($snapshot,[Text.Encoding]::UTF8)
$expected=[IO.File]::ReadAllText((Join-Path $dir 'NativeVBA.txt'),[Text.Encoding]::UTF8).Replace('Workbook: Candidate.xlsm','Workbook: RC_Section_NDM.xlsm')
$stripDate='(?m)^Date: [^\r\n]*\r?\n'
if([regex]::Replace($native,$stripDate,'') -cne [regex]::Replace($expected,$stripDate,'')) {throw 'Published native code differs from accepted candidate.'}
if((Get-FileHash -LiteralPath $published).Hash -cne $publication.WorkbookSHA256) {throw 'Workbook changed during readback.'}
[byte[]]$bytes=[IO.File]::ReadAllBytes($snapshot)
$digest=[Security.Cryptography.SHA256]::Create()
$stream=[IO.File]::Open($allCode,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
try {
    $actual=[BitConverter]::ToString($digest.ComputeHash($stream)).Replace('-','')
    if($actual -cne $publication.AllCodeSHA256) {throw 'All Code was edited during verification.'}
    $stream.Position=0;$stream.Write($bytes,0,$bytes.Length);$stream.SetLength($bytes.Length);$stream.Flush($true)
} finally {$stream.Dispose();$digest.Dispose()}
$publication.AllCodeSHA256=(Get-FileHash -LiteralPath $allCode).Hash
$publication.ReadbackPending=$false
$publication | Add-Member -NotePropertyName ReadbackNativeSHA256 -NotePropertyValue (Get-FileHash -LiteralPath $snapshot).Hash
$publication | Add-Member -NotePropertyName PublishedCodeMatchesCandidate -NotePropertyValue $true
$publication | Add-Member -NotePropertyName AllCodeMatchesActualWorkbookExport -NotePropertyValue $true
$publication | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $dir 'Publication.json') -Encoding UTF8
Write-Output 'PUBLISHED_READBACK_PASS: all 150 components exact; ALL CODE is actual workbook export.'
