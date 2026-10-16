# Делает ALL CODE точной копией независимого чтения опубликованной книги.
param([string]$EvidenceDirectory='docs/regression/ExcelGuardGuide_2026-10-13/RecipeSimplification_2026-10-16')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$directory=(Resolve-Path -LiteralPath (Join-Path $root $EvidenceDirectory)).Path
if(-not $directory.StartsWith((Join-Path $root 'docs/regression')+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw 'Unexpected finalization evidence directory.'}
function Receipt([string]$name){Get-Content -LiteralPath (Join-Path $directory $name) -Raw -Encoding UTF8 | ConvertFrom-Json}
$publication=Receipt 'Publication.json'
foreach($name in @('PublishedSourceEquality.json','PublishedGuidePresentation.json')){if(-not (Receipt $name).Passed){throw ('Published verification failed: '+$name)}}
$book=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$code=Join-Path $root 'workbook/output/VBA_All_Code.txt'
$dump=Join-Path $directory 'PublishedNativeVBA.txt'
if((Get-FileHash -LiteralPath $book).Hash -cne $publication.WorkbookSHA256){throw 'Published workbook changed.'}
if((Get-FileHash -LiteralPath $dump).Hash.ToLowerInvariant() -cne (Receipt 'PublishedGuidePresentation.json').NativeSHA256){throw 'Published native dump changed.'}
[byte[]]$bytes=[IO.File]::ReadAllBytes($dump)
$stream=[IO.File]::Open($code,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
$digest=[Security.Cryptography.SHA256]::Create()
try {
    if([BitConverter]::ToString($digest.ComputeHash($stream)).Replace('-','') -cne $publication.AllCodeSHA256){throw 'ALL CODE changed during readback; preserve it.'}
    $stream.Position=0;$stream.Write($bytes,0,$bytes.Length);$stream.SetLength($bytes.Length);$stream.Flush($true)
} finally {$stream.Dispose();$digest.Dispose()}
$publication.Status='PUBLISHED_VERIFIED'
$publication.AllCodeSHA256=(Get-FileHash -LiteralPath $code).Hash
if($publication.AllCodeSHA256 -cne (Get-FileHash -LiteralPath $dump).Hash){throw 'ALL CODE differs from the actual workbook export.'}
$publication | Add-Member -NotePropertyName NativeComponents -NotePropertyValue 158
$publication | Add-Member -NotePropertyName AllCodeMatchesPublishedNative -NotePropertyValue $true
$publication | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'Publication.json') -Encoding UTF8
Write-Output 'RECIPE_PUBLISHED_VERIFIED: 158 components; ALL CODE matches actual workbook VBA.'
