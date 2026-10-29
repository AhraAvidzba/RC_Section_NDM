# ==========================================================================
# ДЛЯ ТЕСТОВ: readonly-дополнение наблюдения для FAIL с датой в UI Progress.
# Не меняет замороженный observer, VBA, книгу или результаты текущего набора.
# ==========================================================================
param([switch]$SelfTestOnly)
$ErrorActionPreference='Stop'
if($PSVersionTable.PSVersion.Major -ne5){throw 'Windows PowerShell5.1 required'}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Принимает только начало реальной ошибки, с необязательной датой штатного logger.
function Test-FindUiProgressFailures([string]$Text){
 $pattern='(?m)^(?:\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} )?(?:FAIL[: ]|RUNTIME(?: ERROR|:)|SCRIPT ERROR|WATCHDOG_FAILURE|CLEANUP ERROR|RESTORE ERROR)[^\r\n]*'
 return @([regex]::Matches($Text,$pattern)|ForEach-Object Value)
}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Не мешает действующему writer; BOM сохраняет Unicode, ANSI Progress читает CP1251.
function Test-ReadSharedUiText([string]$Path,[Text.Encoding]$Encoding=[Text.Encoding]::Default){
 for($attempt=1;$attempt -le4;$attempt++){
  $stream=$null;$reader=$null
  try{
   $stream=[IO.FileStream]::new($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
   $reader=[IO.StreamReader]::new($stream,$Encoding,$true)
   $text=$reader.ReadToEnd()
   return [pscustomobject]@{Read=$true;Text=$text;CodePage=$reader.CurrentEncoding.CodePage;Attempts=$attempt}
  }catch [IO.IOException]{
   if($attempt -eq4){return [pscustomobject]@{Read=$false;Text='';Attempts=$attempt;Reason=$_.Exception.Message}}
  }finally{if($reader){$reader.Dispose()}elseif($stream){$stream.Dispose()}}
  Start-Sleep -Milliseconds 50
 }
}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Проверяет обе формы FAIL и не принимает ожидаемые отрицательные пробы за сбой.
function Test-UiFailureParserSelfChecks {
 $checks=@(
  @('FAIL: actual',1),@('FAIL actual',1),@('2026-10-28 01:00:01 FAIL: ошибка',1),
  @('2026-10-28 01:00:01 FAIL actual',1),@('RUNTIME ERROR: 14',1),@('RUNTIME: 9',1),
  @('2026-10-28 01:00:01 RUNTIME ERROR: 14',1),@('SCRIPT ERROR: bad',1),
  @('WATCHDOG_FAILURE: timeout',1),@('2026-10-28 01:00:01 RESTORE ERROR: bad',1),
  @('OK: expected FAIL: sample',0),@('2026-10-28 01:00:01 RUN: invalid example FAIL: sample',0),
  @('REQUIRED_GEOMETRY_ERROR: expected rejection',0),@('CIRCLE_REBAR_INPUT: sample; error=0;',0),
  @('GENERAL_NUMBER: expected=1; actual=1',0),@('',0),
  @("OK: one`r`n2026-10-28 01:00:01 FAIL: one`r`nFAIL: two`r`nOK: end",2)
 )
 $failed=@()
 for($i=0;$i -lt$checks.Count;$i++){
  $actual=@(Test-FindUiProgressFailures ([string]$checks[$i][0])).Count
  if($actual -ne[int]$checks[$i][1]){$failed+=,[ordered]@{Case=$i;Expected=$checks[$i][1];Actual=$actual}}
 }
 if($failed.Count){throw ('UI failure parser selftest failed: '+($failed|ConvertTo-Json -Compress))}
 return [ordered]@{Status='PASS_OFFLINE_UI_TIMESTAMP_FAILURE_PARSER_ONLY';Checks=$checks.Count;Failed=0;NativeProcessesStarted=0;FullAcceptanceClaim=$false}
}

$selfTest=Test-UiFailureParserSelfChecks
if($SelfTestOnly){$selfTest|ConvertTo-Json;exit}
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path
$request=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'LaunchRequest.json') -Raw -Encoding UTF8|ConvertFrom-Json
$monitor=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'MonitorLast.json') -Raw -Encoding UTF8|ConvertFrom-Json
$result=[ordered]@{CheckedUTC=[DateTimeOffset]::UtcNow.ToString('o');Status='NOT_APPLICABLE_OUTSIDE_CORE_FULL';CurrentJob=$monitor.CurrentJob;FrozenObserverChanged=$false;NativeProcessesStarted=0;FullAcceptanceClaim=$false;SelfTest=$selfTest;Files=@();Errors=@();InputUnitCasesStarted=0;AllProgressRead=$false}
if([string]$monitor.CurrentJob -cmatch '^Full_(Off|On)$'){
 $mode=if($Matches[1] -ceq'Off'){'No'}else{'Yes'}
 $evidence=[IO.Path]::GetFullPath((Join-Path $root $request.EvidenceDirectory))
 $primary=if([IO.Path]::IsPathRooted([string]$monitor.CurrentLog)){[IO.Path]::GetFullPath([string]$monitor.CurrentLog)}else{[IO.Path]::GetFullPath((Join-Path $root $monitor.CurrentLog))}
 if(-not$primary.StartsWith($evidence+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase) -or [IO.Path]::GetFileName($primary) -cne([string]$monitor.CurrentJob+'.txt')){throw 'Foreign or historical primary cannot select a fixture'}
 $read=Test-ReadSharedUiText $primary ([Text.Encoding]::UTF8)
 if(-not$read.Read){throw 'Transient primary read; repeat the same observation, do not restart'}
 $source=[regex]::Match($read.Text,'(?m)^SOURCE_SHA256: ([0-9A-F]{64})\r?$')
 $fixture=[regex]::Match($read.Text,'(?m)^FIXTURE: ([^\r\n]+)')
 $owner=[regex]::Match($read.Text,'(?m)^TEST_EXCEL_PROCESS: id=(\d+); startTicks=(\d+)\r?$')
 if(-not$source.Success -or $source.Groups[1].Value -cne$request.SourceSHA256 -or -not$read.Text.Contains('GLOBAL_MODE: General.DiagramExtension='+$mode+';') -or -not$fixture.Success -or -not$owner.Success){throw 'Exact current source/mode/fixture/owner header required'}
 $folder=[IO.Path]::GetFullPath((Split-Path -Parent $fixture.Groups[1].Value))
 $temp=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\','/')+[IO.Path]::DirectorySeparatorChar
 if(-not$folder.StartsWith($temp,[StringComparison]::OrdinalIgnoreCase) -or (Split-Path -Leaf $folder) -notlike'RC_NDM_ModeSuite_*'){throw 'Fixture is outside own mode-suite temp scope'}
 $result.Status='CURRENT_CORE_UI_PROGRESS_READ';$result['Primary']=$primary;$result['SourceSHA256']=$source.Groups[1].Value
 $result['OwnExcelPID']=[int]$owner.Groups[1].Value;$result['OwnExcelStartTicksUTC']=$owner.Groups[2].Value
 $cases=@();$allRead=$true
 foreach($file in Get-ChildItem -LiteralPath $folder -Filter '*Progress*' -File){
  $body=Test-ReadSharedUiText $file.FullName
  if(-not$body.Read){$allRead=$false;$result.Files+=,[ordered]@{Name=$file.Name;Read=$false;Reason=$body.Reason};continue}
  $errors=@(Test-FindUiProgressFailures $body.Text)
  foreach($line in $errors){$result.Errors+=,[ordered]@{File=$file.Name;Message=$line}}
  $cases+=@([regex]::Matches($body.Text,'RUN: (audit03\.inputUnits\.[A-Za-z]+\.l\d+\.s\d+\.k\d+)')|ForEach-Object {$_.Groups[1].Value})
  $result.Files+=,[ordered]@{Name=$file.Name;Read=$true;Attempts=$body.Attempts;DecodedCodePage=$body.CodePage;Bytes=$file.Length;LastWriteUTC=$file.LastWriteTimeUtc.ToString('o');ErrorCount=$errors.Count;AdvancedToShapeConfig=$body.Text.Contains('RUN: Audit03 Shape Config')}
 }
 $result.AllProgressRead=$allRead;$result.InputUnitCasesStarted=@($cases|Sort-Object -Unique).Count
 if($result.Errors.Count){$result.Status='CURRENT_PROGRESS_FAILURES_REQUIRE_REVIEW'}elseif(-not$allRead){$result.Status='READING_DEFERRED_REPEAT_SAME_OBSERVATION'}
}
$result|ConvertTo-Json -Depth 9|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'UiProgressSupplementLast.json') -Encoding UTF8
[pscustomobject]@{CheckedUTC=$result.CheckedUTC;Status=$result.Status;CurrentJob=$result.CurrentJob;AllProgressRead=$result.AllProgressRead;InputUnitCasesStarted=$result.InputUnitCasesStarted;Errors=$result.Errors;Files=$result.Files;FullAcceptanceClaim=$false}|ConvertTo-Json -Depth 6
