# ДЛЯ ТЕСТОВ: публикует принятую code-only книгу только при неизменном пользовательском оригинале и закрытых собственных процессах.
$ErrorActionPreference='Stop'
$p=$PSScriptRoot
$root='C:/Users/avidzba/Desktop/программы/В_РАБОТЕ!/RC_Section_NDM'
$target=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$codeTarget=Join-Path $root 'workbook/output/VBA_All_Code.txt'
$ready=Join-Path $p 'FinalCodeOnly03.xlsm'
$export=Join-Path $p 'Inspect_NoFinal03/VBA_All_Code.txt'
$proof=Get-Content -LiteralPath (Join-Path $p 'CodeOnlyPreservation03.json') -Raw|ConvertFrom-Json
$accept=Get-Content -LiteralPath (Join-Path $p 'Acceptance03.json') -Raw|ConvertFrom-Json
$own=Get-Content -LiteralPath (Join-Path $p 'AllOwnClosures03.json') -Raw|ConvertFrom-Json
if($accept.Status -ne 'PASS_DIRECTED' -or -not $own.Passed){throw 'Acceptance missing'}
foreach($item in $own.Identities){
    try{$v=Get-Process -Id $item.PID -ErrorAction Stop
        if($v.ProcessName -eq 'EXCEL' -and $v.StartTime.ToUniversalTime().Ticks -eq [long]$item.Ticks){throw 'Accepted own Excel is live'}
    }catch{if($_.FullyQualifiedErrorId -notlike 'NoProcessFoundForGivenId*'){throw}}
}
if((Get-FileHash -LiteralPath $target).Hash -cne $proof.OriginalSHA256){throw 'User workbook changed since backup; do not publish'}
if((Get-FileHash -LiteralPath $ready).Hash -cne $proof.FinalSHA256){throw 'Candidate changed'}
$handle=[IO.File]::Open($target,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::None)
$handle.Dispose()
$backup=Join-Path $p 'PublicationBackup03'
if(Test-Path -LiteralPath $backup){
    if((Get-FileHash -LiteralPath (Join-Path $backup 'RC_Section_NDM.xlsm')).Hash -cne $proof.OriginalSHA256){throw 'Previous backup differs'}
}else{
    [void][IO.Directory]::CreateDirectory($backup)
    Copy-Item -LiteralPath $target -Destination (Join-Path $backup 'RC_Section_NDM.xlsm')
    Copy-Item -LiteralPath $codeTarget -Destination (Join-Path $backup 'VBA_All_Code.txt')
}
$temporary=Join-Path (Split-Path -Parent $target) 'RC_Section_NDM_RebarMessages03.tmp'
if(Test-Path -LiteralPath $temporary){
    if((Get-FileHash -LiteralPath $temporary).Hash -cne $proof.FinalSHA256){throw 'Unexpected staging content'}
}else{Copy-Item -LiteralPath $ready -Destination $temporary}
if((Get-FileHash -LiteralPath $target).Hash -cne $proof.OriginalSHA256){throw 'User saved workbook during publication preparation'}
[IO.File]::Replace($temporary,$target,(Join-Path $backup 'AtomicPrevious.xlsm'))
Copy-Item -LiteralPath $export -Destination $codeTarget -Force
if((Get-FileHash -LiteralPath $target).Hash -cne $proof.FinalSHA256){throw 'Published book hash mismatch'}
if((Get-FileHash -LiteralPath $export).Hash -cne (Get-FileHash -LiteralPath $codeTarget).Hash){throw 'Published AllCode mismatch'}
$receipt=[ordered]@{Status='PUBLISHED_DIRECTED_PASS';PublishedUTC=[DateTimeOffset]::UtcNow.ToString('o');Workbook=$target;WorkbookSHA256=$proof.FinalSHA256;PreviousUserSHA256=$proof.OriginalSHA256;AllCodeSHA256=(Get-FileHash -LiteralPath $codeTarget).Hash;OnlyVBAChanged=$true;UnchangedZIPParts=110;Backup=$backup;FullSuiteRun=$false}
$evidence=Join-Path $root 'docs/regression/RebarLabelAndMessages_2026-10-30_Run01'
$receipt|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $evidence 'Publication03.json') -Encoding utf8
$receipt|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $p 'Publication03.json') -Encoding utf8
$receipt|ConvertTo-Json -Depth 5
