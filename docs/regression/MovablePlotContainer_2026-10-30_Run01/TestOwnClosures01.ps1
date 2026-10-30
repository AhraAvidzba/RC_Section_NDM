# -------------------- ДЛЯ ТЕСТОВ --------------------
# Проверяет собственные PID вместе со временем старта после обычного выхода; другой процесс не считается прежним Excel.
param([switch]$IncludeNative, [switch]$IncludeDirected)
$ErrorActionPreference='Stop'
$p=$PSScriptRoot;$records=@();$files=@('Build/BuildOwnership.json')
if(Test-Path -LiteralPath (Join-Path $p 'Build02/BuildOwnership.json')){$files+='Build02/BuildOwnership.json'}
if($IncludeNative){$files+='Native01/Ownership.json';if(Test-Path -LiteralPath (Join-Path $p 'Native02/Ownership.json')){$files+='Native02/Ownership.json'}}
if($IncludeDirected){foreach($d in 'Patch02','Patch03','Patch04','Directed01','Directed02','Directed03','Directed04','Integration05','Integration06','Integration07','Final03'){if(Test-Path -LiteralPath (Join-Path $p ($d+'/Ownership.json'))){$files+=$d+'/Ownership.json'}}}
foreach($file in $files){
 $owner=Get-Content -LiteralPath (Join-Path $p $file) -Raw -Encoding UTF8|ConvertFrom-Json
 $state='Unknown';$process=$null;$actualTicks=$null;$cim=$null
 try{
  $process=Get-Process -Id ([int]$owner.PID) -ErrorAction Stop
  if($process.ProcessName-cne'EXCEL'){
   $cim=Get-CimInstance Win32_Process -Filter ('ProcessId='+[int]$owner.PID) -ErrorAction Stop
   if($null-eq$cim-or[string]$cim.Name-ieq'EXCEL.EXE'-or$null-eq$cim.CreationDate){throw ('Independent different-process identity unavailable: '+$file)}
   $actualTicks=[string]$cim.CreationDate.ToUniversalTime().Ticks;$state='DifferentIdentity'
  }else{
   $started=$process.StartTime
   if($null-eq$started){throw ('Exact Excel start time unavailable: '+$file)}
   $actualTicks=[string]$started.ToUniversalTime().Ticks
   if($actualTicks-ceq[string]$owner.StartTicksUTC){$state='ExactOwnStillLive'}else{$state='DifferentIdentity'}
  }
 }catch{if($_.FullyQualifiedErrorId-like'NoProcessFoundForGivenId*'){$state='ConfirmedNotFound'}else{throw}}
 $records+=,[ordered]@{Evidence=$file;PID=$owner.PID;ExpectedStartTicksUTC=$owner.StartTicksUTC;State=$state;ObservedName=$(if($process){$process.ProcessName}else{$null});ObservedStartTicksUTC=$(if($process){$actualTicks}else{$null});IndependentCIMName=$(if($cim){$cim.Name}else{$null});IndependentCreationUTC=$(if($cim){$cim.CreationDate.ToUniversalTime().ToString('o')}else{$null})}
 if($state-notin@('ConfirmedNotFound','DifferentIdentity')){throw ('Own not terminal: '+$file+'/'+$state)}
}
$result=[ordered]@{Passed=$true;AtUTC=[DateTimeOffset]::UtcNow.ToString('o');OwnProcesses=$records;ForceKilled=(Test-Path -LiteralPath (Join-Path $p 'Integration05/OwnTermination.json'))}
$result|ConvertTo-Json -Depth 7|Set-Content -LiteralPath (Join-Path $p 'OwnClosures.json') -Encoding utf8
$result|ConvertTo-Json -Depth 7
