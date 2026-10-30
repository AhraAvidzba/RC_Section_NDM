# -------------------- ДЛЯ ТЕСТОВ --------------------
# Публикует только принятую чистую книгу и её нативный экспорт, сохраняя полный резерв заменяемых файлов.
$ErrorActionPreference='Stop'
$p=$PSScriptRoot;$root='C:/Users/avidzba/Desktop/программы/В_РАБОТЕ!/RC_Section_NDM'
& (Join-Path $p 'TestOwnClosures01.ps1') -IncludeNative -IncludeDirected|Out-Null
$native=Get-Content -LiteralPath (Join-Path $p 'Native02/Result.json') -Raw -Encoding UTF8|ConvertFrom-Json
$equality=Get-Content -LiteralPath (Join-Path $p 'SourceEquality03.json') -Raw -Encoding UTF8|ConvertFrom-Json
$backup=Get-Content -LiteralPath (Join-Path $p 'Backup.json') -Raw -Encoding UTF8|ConvertFrom-Json
$catalog=Get-Content -LiteralPath (Join-Path $p 'CatalogChanges.json') -Raw -Encoding UTF8|ConvertFrom-Json
if(-not$native.Passed-or-not$equality.Passed-or$equality.SourceComponents-ne175-or$equality.ExportComponents-ne181){throw 'Native saved/defaults/source verification required'}
if((Get-FileHash -LiteralPath (Join-Path $root 'tools/build_workbook/GeometryTables.json')).Hash-cne$catalog.CatalogSHA256){throw 'Build catalog changed after native acceptance'}
$pairs=@(@('Final03/RC_Section_NDM.xlsm','workbook/output/RC_Section_NDM.xlsm',$backup.BookSHA256,$native.BookSHA256),@('Native02/VBA_All_Code.txt','workbook/output/VBA_All_Code.txt',$backup.AllCodeSHA256,''))
$backups=Join-Path $p 'BeforePublication';if(Test-Path -LiteralPath $backups){throw 'Inspect publication attempt first'}
[void][IO.Directory]::CreateDirectory($backups)
foreach($pair in $pairs){
 $target=Join-Path $root $pair[1]
 if((Get-FileHash -LiteralPath $target).Hash-cne$pair[2]){throw ('Target changed since backup: '+$pair[1])}
 $lock=[IO.File]::Open($target,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None);$lock.Dispose()
 if($pair[3]-and(Get-FileHash -LiteralPath (Join-Path $p $pair[0])).Hash-cne$pair[3]){throw 'Accepted candidate changed'}
}
$receipts=@()
foreach($pair in $pairs){
 $source=Join-Path $p $pair[0];$target=Join-Path $root $pair[1]
 $temp=$target+'.SectionDefaults_new';if(Test-Path -LiteralPath $temp){throw 'Unexpected publication temporary file'}
 Copy-Item -LiteralPath $source -Destination $temp
 $saved=Join-Path $backups ([IO.Path]::GetFileName($target))
 [IO.File]::Replace($temp,$target,$saved)
 $hash=(Get-FileHash -LiteralPath $target).Hash
 if($hash-cne(Get-FileHash -LiteralPath $source).Hash-or(Get-FileHash -LiteralPath $saved).Hash-cne$pair[2]){throw 'Publication copy/hash mismatch'}
 $receipts+=,[ordered]@{Target=$pair[1];SHA256=$hash;PreviousSHA256=$pair[2];Backup=$saved}
}
[ordered]@{Passed=$true;AtUTC=[DateTimeOffset]::UtcNow.ToString('o');CleanRebuild=$true;InputsAndResultsCopied=$false;Files=$receipts}|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $p 'Publication.json') -Encoding utf8
Get-Content -LiteralPath (Join-Path $p 'Publication.json') -Raw -Encoding UTF8
