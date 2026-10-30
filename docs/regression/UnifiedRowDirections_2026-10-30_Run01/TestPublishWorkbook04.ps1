# ДЛЯ ТЕСТОВ: публикует принятую книгу с сохранёнными данными после проверки закрытий и неизменности оригинала.
$ErrorActionPreference='Stop'
$p=$PSScriptRoot;$root='C:/Users/avidzba/Desktop/программы/В_РАБОТЕ!/RC_Section_NDM'
$proof=Get-Content -LiteralPath (Join-Path $p 'PackagePreservation02.json') -Raw -Encoding utf8|ConvertFrom-Json
$accept=Get-Content -LiteralPath (Join-Path $p 'Acceptance04.json') -Raw -Encoding utf8|ConvertFrom-Json
$own=Get-Content -LiteralPath (Join-Path $p 'AllOwnClosures04.json') -Raw -Encoding utf8|ConvertFrom-Json
if($accept.Status -ne 'PASS_DIRECTED' -or $accept.AssertionsFailed -ne 0 -or -not $own.Passed){throw 'Missing directed acceptance'}
foreach($identity in $own.Identities){
    try{$v=Get-Process -Id $identity.PID -ErrorAction Stop
        if($v.ProcessName -eq 'EXCEL' -and $v.StartTime.ToUniversalTime().Ticks -eq [long]$identity.Ticks){throw 'Own Excel unexpectedly live'}
    }catch{if($_.FullyQualifiedErrorId -notlike 'NoProcessFoundForGivenId*'){throw}}
}
$target=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$codeTarget=Join-Path $root 'workbook/output/VBA_All_Code.txt'
$ready=Join-Path $p 'FinalPreserved02.xlsm';$export=Join-Path $p 'Inspect_No_Final03/VBA_All_Code.txt'
if((Get-FileHash -LiteralPath $target).Hash -cne $proof.OriginalSHA256){throw 'User workbook changed since backup; publication stopped'}
if((Get-FileHash -LiteralPath $ready).Hash -cne $proof.FinalSHA256){throw 'Final candidate changed'}
foreach($file in @($target,$codeTarget)){$handle=[IO.File]::Open($file,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::None);$handle.Dispose()}
$backup=Join-Path $p 'BeforePublication04'
if(Test-Path -LiteralPath $backup){throw 'Backup directory already exists; inspect previous publication'}
[void][IO.Directory]::CreateDirectory($backup)
Copy-Item -LiteralPath $target -Destination (Join-Path $backup 'RC_Section_NDM.xlsm')
Copy-Item -LiteralPath $codeTarget -Destination (Join-Path $backup 'VBA_All_Code.txt')
$bookTemp=Join-Path (Split-Path -Parent $target) 'RC_Section_NDM_Directions04.tmp'
$codeTemp=Join-Path (Split-Path -Parent $target) 'VBA_All_Code_Directions04.tmp'
foreach($file in @($bookTemp,$codeTemp)){if(Test-Path -LiteralPath $file){throw 'Unexpected staging file'}}
Copy-Item -LiteralPath $ready -Destination $bookTemp;Copy-Item -LiteralPath $export -Destination $codeTemp
if((Get-FileHash -LiteralPath $target).Hash -cne $proof.OriginalSHA256){throw 'User saved workbook during preparation'}
[IO.File]::Replace($bookTemp,$target,(Join-Path $backup 'AtomicPrevious.xlsm'))
[IO.File]::Replace($codeTemp,$codeTarget,(Join-Path $backup 'AtomicPreviousAllCode.txt'))
if((Get-FileHash -LiteralPath $target).Hash -cne $proof.FinalSHA256){throw 'Published book mismatch'}
if((Get-FileHash -LiteralPath $codeTarget).Hash -cne (Get-FileHash -LiteralPath $export).Hash){throw 'Published AllCode mismatch'}
$receipt=[ordered]@{Status='PUBLISHED_DIRECTED_PASS';PublishedUTC=[DateTimeOffset]::UtcNow.ToString('o');Workbook=$target;OriginalSHA256=$proof.OriginalSHA256;WorkbookSHA256=$proof.FinalSHA256;AllCodeSHA256=(Get-FileHash -LiteralPath $codeTarget).Hash;ChangedZIPParts=$proof.ChangedZIPParts;UnchangedZIPParts=109;UserDataResultsAndPlotPreserved=$true;Backup=$backup;FullSuiteRun=$false}
$e=Join-Path $root 'docs/regression/UnifiedRowDirections_2026-10-30_Run01'
$receipt|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $p 'Publication04.json') -Encoding utf8
$receipt|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $e 'Publication04.json') -Encoding utf8
$receipt|ConvertTo-Json -Depth 6
