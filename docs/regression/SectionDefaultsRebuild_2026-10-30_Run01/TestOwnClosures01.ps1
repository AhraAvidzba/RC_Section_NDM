# -------------------- ДЛЯ ТЕСТОВ --------------------
# Проверяет собственные PID вместе со временем старта после обычного выхода; другой процесс не считается прежним Excel.
param([switch]$IncludeNative)
$ErrorActionPreference='Stop'
$p=$PSScriptRoot;$records=@();$files=@('Build/BuildOwnership.json')
if($IncludeNative){$files+='Native01/Ownership.json'}
foreach($file in $files){
 $owner=Get-Content -LiteralPath (Join-Path $p $file) -Raw -Encoding UTF8|ConvertFrom-Json
 $state='Unknown';$process=$null
 try{
  $process=Get-Process -Id ([int]$owner.PID) -ErrorAction Stop
  $actualTicks=[string]$process.StartTime.ToUniversalTime().Ticks
  if($process.ProcessName-ceq'EXCEL'-and$actualTicks-ceq[string]$owner.StartTicksUTC){$state='ExactOwnStillLive'}else{$state='DifferentIdentity'}
 }catch{if($_.FullyQualifiedErrorId-like'NoProcessFoundForGivenId*'){$state='ConfirmedNotFound'}else{throw}}
 $records+=,[ordered]@{Evidence=$file;PID=$owner.PID;ExpectedStartTicksUTC=$owner.StartTicksUTC;State=$state;ObservedName=$(if($process){$process.ProcessName}else{$null});ObservedStartTicksUTC=$(if($process){$actualTicks}else{$null})}
 if($state-notin@('ConfirmedNotFound','DifferentIdentity')){throw ('Own not terminal: '+$file+'/'+$state)}
}
$result=[ordered]@{Passed=$true;AtUTC=[DateTimeOffset]::UtcNow.ToString('o');OwnProcesses=$records;ForceKilled=$false}
$result|ConvertTo-Json -Depth 7|Set-Content -LiteralPath (Join-Path $p 'OwnClosures.json') -Encoding utf8
$result|ConvertTo-Json -Depth 7
