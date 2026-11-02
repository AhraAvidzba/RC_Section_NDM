# ==================== ДЛЯ ТЕСТОВ ====================
# После whole PASS и двух сохранённых native проверок публикует только точные private копии с новым VBA.
# Обе исходные книги блокируются до записи; их последние байты сохраняются в проверенные резервные файлы.
# При ошибке записи восстанавливает затронутые файлы из этих резервов, удерживая те же исключительные дескрипторы.
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'TestAssertCurrentRun02Pass60.ps1')
$pass=Test-AssertCurrentRun02Pass60 -RequireTaskRemoved -RequireAdditionalIO
$receiptPath=Join-Path $PSScriptRoot 'UserBooksCodeOnlyPublication64.json'
if(Test-Path -LiteralPath $receiptPath){throw 'Publication receipt exists; inspect instead of repeating'}
$prep=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'PreparedUserBooks62.json') -Raw -Encoding UTF8|ConvertFrom-Json
if($prep.Status -cne 'PRIVATE_CODE_ONLY_BOOKS_PREPARED_NATIVE_VALIDATION_REQUIRED' -or $prep.AcceptedSourceSHA256 -cne $pass.SourceSHA256 -or $prep.SourceCommit -cne $pass.SourceCommit -or -not $prep.OfficialFullPass -or @($prep.Books).Count -ne 2 -or $prep.Published){throw 'Exact guarded private preparation required'}
$labels=@('Main','Heavy');$candidates=@();$nativeProofs=@()
foreach($label in $labels){
    $matches=@($prep.Books|Where-Object Label -ceq $label)
    if($matches.Count -ne 1){throw 'Exactly one prepared Main and Heavy required'}
    $item=$matches[0]
    $expected=if($label -ceq 'Main'){Join-Path $pass.Root 'workbook/output/RC_Section_NDM.xlsm'}else{Join-Path $pass.Root 'workbook/output/NDM_HeavyCircle_1232_30_2026-11-01/RC_Section_NDM_HeavyCircle_1232_30.xlsm'}
    if([IO.Path]::GetFullPath([string]$item.Original) -ine [IO.Path]::GetFullPath($expected) -or -not $item.AllOtherUncompressedPartBytesExact -or -not $item.UserDataAndResultsPreserved -or $item.Recalculated -or @($item.ChangedParts).Count -ne 1 -or $item.ChangedParts[0] -cne 'xl/vbaProject.bin' -or $item.PreservedPartCount -ne $item.PartCount-1){throw 'Exact code-only user-data preservation contract required'}
    if((Get-FileHash -LiteralPath $item.Prepared).Hash -cne $item.PreparedSHA256){throw 'Prepared candidate changed after comparison'}
    $directory=Join-Path $PSScriptRoot ('UserBooksNative63/'+$label)
    $accepted=Get-Content -LiteralPath (Join-Path $directory 'Accepted.json') -Raw -Encoding UTF8|ConvertFrom-Json
    $native=Get-Content -LiteralPath (Join-Path $directory 'Native.json') -Raw -Encoding UTF8|ConvertFrom-Json
    $equality=Get-Content -LiteralPath (Join-Path $directory 'SourceEquality.json') -Raw -Encoding UTF8|ConvertFrom-Json
    $owner=Get-Content -LiteralPath (Join-Path $directory 'OwnClosure.json') -Raw -Encoding UTF8|ConvertFrom-Json
    if($accepted.Status -cne 'SAVED_USER_BOOK_CODE_AND_DATA_VERIFIED' -or $accepted.PreparedSHA256 -cne $item.PreparedSHA256 -or $accepted.SourceAcceptedSHA256 -cne $pass.SourceSHA256 -or -not $accepted.OwnIdentityTerminal -or $accepted.NDSRecalculated -or $native.Components -ne 187 -or -not $native.ReadOnly -or -not $native.ManualCalculation -or -not $native.MacrosDisabled -or -not $native.NoVBAMacroRun -or $native.Saved -or $native.NDSRecalculated -or $native.WorkbookSHA256 -cne $item.PreparedSHA256){throw 'Saved native187/manual/readonly/no-recalculation proof required'}
    if($accepted.NativeProofSHA256 -cne (Get-FileHash -LiteralPath (Join-Path $directory 'Native.json')).Hash -or $accepted.SourceEqualitySHA256 -cne (Get-FileHash -LiteralPath (Join-Path $directory 'SourceEquality.json')).Hash -or $accepted.OwnClosureSHA256 -cne (Get-FileHash -LiteralPath (Join-Path $directory 'OwnClosure.json')).Hash){throw 'Native evidence changed after acceptance'}
    if(-not $equality.Passed -or $equality.SourceComponents -ne 181 -or $equality.ExportComponents -ne 187 -or @($equality.Differences).Count -or @($equality.Missing).Count -or -not $owner.ExactIdentityTerminal -or -not $owner.OrdinaryQuit -or $owner.ForceKilled){throw 'Exact source and ordinary terminal own proof required'}
    Test-AssertRun02OwnTerminal60 ([pscustomobject]@{Name='EXCEL';PID=$owner.PID;StartTicksUTC=$owner.StartTicksUTC})
    $candidates+=,$item;$nativeProofs+=,[ordered]@{Label=$label;AcceptedSHA256=(Get-FileHash -LiteralPath (Join-Path $directory 'Accepted.json')).Hash;OwnPID=$owner.PID;OwnStartTicksUTC=$owner.StartTicksUTC;OwnTerminal=$true;Code181Plus6Verified=$true;DataRecalculated=$false}
}

. (Join-Path $PSScriptRoot 'TestWriteProtectedBooks65.ps1')
$baseProof=[ordered]@{AcceptedSourceSHA256=$pass.SourceSHA256;SourceCommit=$pass.SourceCommit;FullPassAppliesOnlyToAcceptedSource=$true;NativeProofs=$nativeProofs;NDSRecalculated=$false;UserDataAndResultsPreserved=$true;AllOtherZipPartsByteExact=$true;NativeSourceBodies181Plus6Verified=$true}
$backup=Join-Path $PSScriptRoot 'BeforeUserBooksCodeOnlyPublication64'
[void](Test-WriteProtectedBooks65 $candidates $backup $receiptPath $baseProof)
Write-Output 'BOTH_USER_BOOKS_CODE_ONLY_PUBLISHED_RESULTS_AND_INPUTS_PRESERVED'
