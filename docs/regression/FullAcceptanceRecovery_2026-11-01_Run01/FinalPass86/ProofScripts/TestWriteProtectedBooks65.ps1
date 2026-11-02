# ==================== ДЛЯ ТЕСТОВ ====================
# Файловая транзакция для двух проверенных книг; gate и native-проверки выполняет вызывающая процедура.
# -------------------- ДЛЯ ТЕСТОВ --------------------
# Хэширует именно удерживаемый дескриптор; путь не может быть заменён пользователем до освобождения блокировки.
function Test-HashHeldStream65([IO.FileStream]$stream){
    $stream.Position=0;$hash=[Security.Cryptography.SHA256]::Create()
    try{[BitConverter]::ToString($hash.ComputeHash($stream)).Replace('-','')}finally{$hash.Dispose();$stream.Position=0}
}


# -------------------- ДЛЯ ТЕСТОВ --------------------
# Проверяет CAS под исключительной блокировкой и восстанавливает любые частичные записи из заранее проверенного резерва.
# TestFailDuringWrite используется только синтетической проверкой восстановления, штатный вызов оставляет ноль.
function Test-WriteProtectedBooks65([object[]]$candidates,[string]$backup,[string]$receiptPath,[object]$BaseProof,[int]$TestFailDuringWrite=0){
    if(Test-Path -LiteralPath $receiptPath){throw 'Transaction receipt exists'}
    if(Test-Path -LiteralPath $backup){throw 'Transaction backup directory exists'}
$handles=@();$written=@();$books=@();$proof=$BaseProof
$proof['Status']='PUBLISHING';$proof['AtUTC']=[DateTimeOffset]::UtcNow.ToString('o');$proof['Books']=@();$proof['OriginalsRestoredOnFailure']=$false
try{
    # Сначала захватываются обе книги. Если хотя бы одна открыта пользователем, ни одна не изменяется.
    foreach($item in $candidates){
        $stream=[IO.File]::Open([string]$item.Original,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
        $handles+=,[pscustomobject]@{Item=$item;Stream=$stream;Backup=''}
        if((Test-HashHeldStream65 $stream) -cne $item.OriginalSHA256){throw ('Latest saved user book changed after private preparation: '+$item.Label)}
    }
    [void][IO.Directory]::CreateDirectory($backup)
    foreach($held in $handles){
        $directory=Join-Path $backup $held.Item.Label;[void][IO.Directory]::CreateDirectory($directory)
        $path=Join-Path $directory ([IO.Path]::GetFileName([string]$held.Item.Original))
        $copy=[IO.File]::Open($path,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None)
        try{$held.Stream.Position=0;$held.Stream.CopyTo($copy);$copy.Flush($true)}finally{$copy.Dispose();$held.Stream.Position=0}
        if((Get-FileHash -LiteralPath $path).Hash -cne $held.Item.OriginalSHA256){throw 'Original backup hash differs; nothing will be published'}
        $held.Backup=$path
    }
    $proof.Status='BOTH_CURRENT_ORIGINALS_LOCKED_AND_BACKED_UP'
    $proof|ConvertTo-Json -Depth 7|Set-Content -LiteralPath (Join-Path $backup 'BeforePublication.json') -Encoding UTF8
    foreach($held in $handles){
        $bytes=[IO.File]::ReadAllBytes([string]$held.Item.Prepared)
        $hash=[Security.Cryptography.SHA256]::Create();try{$sha=[BitConverter]::ToString($hash.ComputeHash($bytes)).Replace('-','')}finally{$hash.Dispose()}
        if($sha -cne $held.Item.PreparedSHA256){throw 'Candidate bytes changed at publication'}
        $written+=,[string]$held.Item.Label
        if($TestFailDuringWrite -gt 0 -and $written.Count -eq $TestFailDuringWrite){
            $held.Stream.Position=0;$held.Stream.SetLength($bytes.Length);$held.Stream.Write($bytes,0,[math]::Max(1,[int]($bytes.Length/2)));$held.Stream.Flush($true)
            throw 'INJECTED_TEST_PARTIAL_WRITE_FAILURE'
        }
        $held.Stream.Position=0;$held.Stream.SetLength($bytes.Length);$held.Stream.Write($bytes,0,$bytes.Length);$held.Stream.Flush($true)
        if((Test-HashHeldStream65 $held.Stream) -cne $held.Item.PreparedSHA256){throw 'Published file differs from saved native candidate'}
        $books+=,[ordered]@{Label=$held.Item.Label;Path=$held.Item.Original;BeforeSHA256=$held.Item.OriginalSHA256;AfterSHA256=$held.Item.PreparedSHA256;Backup=$held.Backup;BackupSHA256=$held.Item.OriginalSHA256;ChangedParts=@('xl/vbaProject.bin');OtherZipPartsByteExact=[bool]$BaseProof.AllOtherZipPartsByteExact;UserDataPreserved=[bool]$BaseProof.UserDataAndResultsPreserved;SavedNDSResultsPreserved=[bool]$BaseProof.UserDataAndResultsPreserved;SourceBodies181Plus6Verified=[bool]$BaseProof.NativeSourceBodies181Plus6Verified;Recalculated=$false}
    }
    $proof.Status='BOTH_USER_BOOKS_CODE_ONLY_PUBLISHED_WITH_DATA_PRESERVED';$proof.Books=$books
}catch{
    $failure=$_.Exception.ToString();$restored=$true;$restoreErrors=@()
    foreach($held in $handles){
        if($written -ccontains [string]$held.Item.Label){
            try{
                if(-not $held.Backup -or (Get-FileHash -LiteralPath $held.Backup).Hash -cne $held.Item.OriginalSHA256){throw 'Verified backup unavailable'}
                $bytes=[IO.File]::ReadAllBytes($held.Backup);$held.Stream.Position=0;$held.Stream.SetLength($bytes.Length);$held.Stream.Write($bytes,0,$bytes.Length);$held.Stream.Flush($true)
                if((Test-HashHeldStream65 $held.Stream) -cne $held.Item.OriginalSHA256){throw 'Original restoration hash differs'}
            }catch{$restored=$false;$restoreErrors+=,$_.Exception.Message}
        }
    }
    $proof.Status=if($restored){'PUBLICATION_FAILED_CURRENT_ORIGINALS_PRESERVED_OR_RESTORED'}else{'PUBLICATION_RESTORE_REQUIRES_ATTENTION'}
    $proof.OriginalsRestoredOnFailure=$restored;$proof['Error']=$failure;$proof['RestorationErrors']=$restoreErrors
    throw
}finally{
    foreach($held in $handles){$held.Stream.Dispose()}
    $proof['FinishedUTC']=[DateTimeOffset]::UtcNow.ToString('o')
    $proof|ConvertTo-Json -Depth 9|Set-Content -LiteralPath $receiptPath -Encoding UTF8
}
foreach($item in $candidates){if((Get-FileHash -LiteralPath $item.Original).Hash -cne $item.PreparedSHA256){throw 'Published file changed after exclusive close'}}

    return $proof
}
