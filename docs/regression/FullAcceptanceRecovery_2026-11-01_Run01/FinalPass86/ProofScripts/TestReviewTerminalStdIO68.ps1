# ==================== ДЛЯ ТЕСТОВ ====================
# После фактического выхода всех own читает оба terminal stdio целиком; mixed encoding и ошибки требуют разбора.
# Не заменяет primary/official acceptance и не запускает Excel/CAD.
param([switch]$PrintText)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
. (Join-Path $PSScriptRoot 'TestAssertCurrentRun02Pass60.ps1')
$root='C:/Users/avidzba/Desktop/программы/В_РАБОТЕ!/RC_Section_NDM'
$launch=Join-Path $root 'docs/regression/FullSuiteLaunch_2026-11-01_Current_Run02'
if(-not (Test-Path -LiteralPath (Join-Path $launch 'Completion.json'))){throw 'Full still live; terminal stdio review not allowed'}
$full=Get-Content -LiteralPath (Join-Path $launch 'FullLaunch.json') -Raw -Encoding UTF8|ConvertFrom-Json
$own=Get-Content -LiteralPath (Join-Path $launch 'AllOwnClosuresReviewed.json') -Raw -Encoding UTF8|ConvertFrom-Json
if($full.PID -ne 98996 -or [string]$full.StartTicks -cne '639291636492982491' -or $full.SourceSHA256 -cne '831D725C06439CA1B771645F18B0D546169C1E22A97BA7688D64CB3F8509B74A' -or -not $own.AllOwnIdentitiesTerminal -or -not $own.PrimaryLogsReviewed -or $own.SourceSHA256 -cne $full.SourceSHA256 -or $own.SourceCommit -cne $full.SourceCommit){throw 'Exact entire own terminal proof required'}
foreach($identity in $own.Identities){Test-AssertRun02OwnTerminal60 $identity}
$records=@();$errors=@();$fallback=$false
foreach($name in @('FullLaunch.stdout.log','FullLaunch.stderr.log')){
    $path=Join-Path $launch $name
    $stream=[IO.File]::Open($path,[IO.FileMode]::Open,[IO.FileAccess]::Read,([IO.FileShare]::ReadWrite-bor[IO.FileShare]::Delete));$buffer=New-Object IO.MemoryStream
    try{$stream.CopyTo($buffer);$bytes=$buffer.ToArray()}finally{$buffer.Dispose();$stream.Dispose()}
    $encoding='UTF8_STRICT'
    try{$text=([Text.UTF8Encoding]::new($false,$true)).GetString($bytes).TrimStart([char]0xFEFF)}catch{$text=[Text.Encoding]::GetEncoding(1251).GetString($bytes);$encoding='CP1251_FALLBACK_REQUIRES_MANUAL_REVIEW';$fallback=$true}
    $matches=@([regex]::Matches($text,'(?im)^\s*(?:FULL_DRIVER_ERROR:|FAIL:|ERROR:|RUNTIME ERROR:|CLEANUP ERROR:|WATCHDOG_FAILURE:|Traceback \(most recent call last\):|FAILED \(|\+\s+CategoryInfo\s*:)[^\r\n]*')|ForEach-Object{$_.Value})
    foreach($line in $matches){$errors+=,($name+': '+$line)}
    $sha=[Security.Cryptography.SHA256]::Create();try{$hash=[BitConverter]::ToString($sha.ComputeHash($bytes)).Replace('-','')}finally{$sha.Dispose()}
    $records+=,[ordered]@{Name=$name;Bytes=$bytes.Length;SHA256=$hash;AllBytesRead=$true;SharedReadWriteDelete=$true;Encoding=$encoding;Errors=$matches;Text=$text}
    if($PrintText){Write-Output ('FILE: '+$name);Write-Output $text}
}
$status=if($errors.Count){'TERMINAL_STDIO_RUNTIME_ERRORS_REQUIRE_REVIEW'}elseif($fallback){'TERMINAL_STDIO_MIXED_ENCODING_REQUIRES_MANUAL_REVIEW'}else{'ENTIRE_TERMINAL_STDIO_READ_NO_RUNTIME_ERRORS'}
[ordered]@{Status=$status;AtUTC=[DateTimeOffset]::UtcNow.ToString('o');SourceSHA256=$full.SourceSHA256;SourceCommit=$full.SourceCommit;AllOwnIdentitiesTerminal=$true;Files=$records;Errors=$errors;DecodedFallback=$fallback;FullAcceptanceClaim=$false}|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'TerminalStdIOReviewed68.json') -Encoding UTF8
[pscustomobject]@{Status=$status;Files=$records.Count;Errors=$errors.Count;Fallback=$fallback;FullAcceptanceClaim=$false}|ConvertTo-Json -Compress
if($errors.Count -or $fallback){exit 1}
