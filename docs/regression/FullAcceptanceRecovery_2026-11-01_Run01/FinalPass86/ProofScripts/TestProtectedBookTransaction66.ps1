# ==================== ДЛЯ ТЕСТОВ ====================
# Проверяет реальный файловый helper на собственных синтетических байтах: публикация, частичная запись, CAS и блокировки.
# Ни одной настоящей книги и ни одного процесса Excel/CAD; полный PASS не имитируется.
$ErrorActionPreference='Stop'
Import-Module C:/Windows/System32/WindowsPowerShell/v1.0/Modules/Microsoft.PowerShell.Utility/Microsoft.PowerShell.Utility.psd1 -Global
. (Join-Path $PSScriptRoot 'TestWriteProtectedBooks65.ps1')
$directory=Join-Path $PSScriptRoot 'SyntheticFileTransaction66'
if(Test-Path -LiteralPath $directory){throw 'Synthetic test evidence exists; do not overwrite'}
[void][IO.Directory]::CreateDirectory($directory)
$results=@()
foreach($case in @('Success','PartialFirst','PartialSecond','ModifiedOriginal','LockedSecond','ModifiedCandidate')){
    $folder=Join-Path $directory $case;[void][IO.Directory]::CreateDirectory($folder)
    $items=@();$before=@{}
    foreach($label in @('Main','Heavy')){
        $original=Join-Path $folder ($label+'.bin');$candidate=Join-Path $folder ($label+'.candidate.bin')
        $old=New-Object byte[] 4097;$new=New-Object byte[] 7311
        for($i=0;$i -lt $old.Length;$i++){$old[$i]=[byte](($i*17+$label.Length)%256)}
        for($i=0;$i -lt $new.Length;$i++){$new[$i]=[byte](($i*31+$label.Length+2)%256)}
        [IO.File]::WriteAllBytes($original,$old);[IO.File]::WriteAllBytes($candidate,$new)
        $items+=,[pscustomobject]@{Label=$label;Original=$original;OriginalSHA256=(Get-FileHash -LiteralPath $original).Hash;Prepared=$candidate;PreparedSHA256=(Get-FileHash -LiteralPath $candidate).Hash}
    }
    if($case -ceq 'ModifiedOriginal'){[IO.File]::WriteAllBytes($items[0].Original,[byte[]](1,7,3,9,0,255))}
    if($case -ceq 'ModifiedCandidate'){[IO.File]::WriteAllBytes($items[1].Prepared,[byte[]](254,1,2,3,4))}
    foreach($item in $items){$before[$item.Label]=(Get-FileHash -LiteralPath $item.Original).Hash}
    $backup=Join-Path $folder 'Backup';$receipt=Join-Path $folder 'Receipt.json';$lock=$null;$errorText='';$failed=$false
    $fault=if($case -ceq 'PartialFirst'){1}elseif($case -ceq 'PartialSecond'){2}else{0}
    $base=[ordered]@{SimulationOnly=$true;ExcelStarted=0;FullAcceptanceClaim=$false;UserDataAndResultsPreserved=$false;AllOtherZipPartsByteExact=$false;NativeSourceBodies181Plus6Verified=$false}
    try{
        if($case -ceq 'LockedSecond'){$lock=[IO.File]::Open($items[1].Original,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::None)}
        [void](Test-WriteProtectedBooks65 $items $backup $receipt $base $fault)
    }catch{$failed=$true;$errorText=$_.Exception.Message}finally{if($lock){$lock.Dispose()}}
    $proof=Get-Content -LiteralPath $receipt -Raw -Encoding UTF8|ConvertFrom-Json
    if($case -ceq 'Success'){
        if($failed -or $proof.Status -cne 'BOTH_USER_BOOKS_CODE_ONLY_PUBLISHED_WITH_DATA_PRESERVED'){throw ('Synthetic success failed: '+$errorText)}
        foreach($item in $items){if((Get-FileHash -LiteralPath $item.Original).Hash -cne $item.PreparedSHA256){throw 'Synthetic committed payload differs'}}
    }else{
        if(-not $failed -or $proof.Status -cne 'PUBLICATION_FAILED_CURRENT_ORIGINALS_PRESERVED_OR_RESTORED' -or -not $proof.OriginalsRestoredOnFailure){throw ('Expected protected failure missing: '+$case+'; '+$errorText)}
        if($fault -gt 0 -and $errorText -cne 'INJECTED_TEST_PARTIAL_WRITE_FAILURE'){throw 'Partial write injection did not execute'}
        foreach($item in $items){if((Get-FileHash -LiteralPath $item.Original).Hash -cne $before[$item.Label]){throw ('Latest original was not preserved/restored: '+$case+'/'+$item.Label)}}
    }
    if($case -in @('Success','PartialFirst','PartialSecond','ModifiedCandidate')){
        foreach($item in $items){$saved=Join-Path (Join-Path $backup $item.Label) ([IO.Path]::GetFileName($item.Original));if((Get-FileHash -LiteralPath $saved).Hash -cne $before[$item.Label]){throw 'Backup did not preserve source bytes'}}
    }
    # После выхода функции исключительных блокировок не осталось: проверяем новое эксклюзивное открытие обоих файлов.
    foreach($item in $items){$open=[IO.File]::Open($item.Original,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None);$open.Dispose()}
    $results+=,[ordered]@{Case=$case;Passed=$true;ExpectedFailure=$case -cne 'Success';PartialWriteFaultInjected=$fault;LatestOriginalsPreservedOnFailure=$case -cne 'Success';ExclusiveHandlesReleased=$true;ReceiptSHA256=(Get-FileHash -LiteralPath $receipt).Hash}
}
[ordered]@{Status='PASS_SIX_SYNTHETIC_FILE_TRANSACTION_CASES';AtUTC=[DateTimeOffset]::UtcNow.ToString('o');Cases=$results;RealWorkbooksTouched=0;ExcelOrCADStarted=0;FullAcceptanceClaim=$false;HelperSHA256=(Get-FileHash -LiteralPath (Join-Path $PSScriptRoot 'TestWriteProtectedBooks65.ps1')).Hash}|ConvertTo-Json -Depth 7|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'ProtectedBookTransaction66.json') -Encoding UTF8
Get-Content -LiteralPath (Join-Path $PSScriptRoot 'ProtectedBookTransaction66.json') -Raw -Encoding UTF8
