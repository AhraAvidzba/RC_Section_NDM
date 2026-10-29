# -------------------- ДЛЯ ТЕСТОВ --------------------
# Заново читает все известные собственные PID/UTC ticks перед полным запуском, не закрывая процессы.
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path
$private='C:/Users/avidzba/.codex/visualizations/2026/10/12/01a12ece-8192-7141-89b3-f83f1176184c/UnifiedRebarPaths_2026-10-29_Run01'
$known=[Collections.Generic.List[object]]::new()
foreach($name in @('docs/regression/FullSuiteLaunch_2026-10-29_PlotLabels_Run02/AllOwnClosuresReviewed.json','docs/regression/FullSuiteLaunch_2026-10-29_PlotLabels_Run02/ReleaseEligibility.json')){
 $j=Get-Content -LiteralPath (Join-Path $root $name) -Raw -Encoding UTF8|ConvertFrom-Json
 $rows=if($j.PSObject.Properties.Name-contains'Identities'){$j.Identities}else{$j.OwnIdentities}
 foreach($row in $rows){$known.Add([pscustomobject]@{Name=$row.Name;PID=$row.PID;StartTicksUTC=[string]$row.StartTicksUTC;Provenance=$name})}
}
foreach($name in @('Build01','Build02','Build03','Build04')){
 $j=Get-Content -LiteralPath (Join-Path $PSScriptRoot ($name+'/BuildOwnership.json')) -Raw -Encoding UTF8|ConvertFrom-Json
 $known.Add([pscustomobject]@{Name='EXCEL';PID=$j.PID;StartTicksUTC=[string]$j.StartTicksUTC;Provenance=$name})
}
foreach($file in Get-ChildItem -LiteralPath $PSScriptRoot -Filter 'Directed*.ProcessClosure.json' -File){
 $j=Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8|ConvertFrom-Json
 if(-not$j.Passed -or -not$j.Closure.Identity.Valid){throw ('Own directed closure missing: '+$file.Name)}
 $known.Add([pscustomobject]@{Name='EXCEL';PID=$j.Closure.Identity.PID;StartTicksUTC=[string]$j.Closure.Identity.Ticks;Provenance=$file.Name})
}
foreach($name in @('Native02Ownership.json','Native04Ownership.json')){
 $j=Get-Content -LiteralPath (Join-Path $private $name) -Raw -Encoding UTF8|ConvertFrom-Json
 $known.Add([pscustomobject]@{Name='EXCEL';PID=$j.PID;StartTicksUTC=[string]$j.StartTicksUTC;Provenance=$name})
}
$seen=@{};$observed=@()
foreach($row in $known){
 $key=$row.Name+':'+$row.PID+':'+$row.StartTicksUTC
 if($seen.ContainsKey($key)){continue};$seen[$key]=$true
 $state=[ordered]@{Name=$row.Name;PID=$row.PID;StartTicksUTC=$row.StartTicksUTC;Provenance=$row.Provenance;State='Unknown';OwnIdentityTerminal=$false;ForceKilled=$false;ObservedUTC=[DateTimeOffset]::UtcNow.ToString('o')}
 $p=$null
 try{$p=Get-Process -Id ([int]$row.PID) -ErrorAction Stop}catch{if($_.FullyQualifiedErrorId-like'NoProcessFoundForGivenId*'){$state.State='Absent';$state.OwnIdentityTerminal=$true}else{throw}}
 if($p){
  $cim=Get-CimInstance Win32_Process -Filter ('ProcessId='+[int]$row.PID) -ErrorAction Stop
  if($null-eq$cim-or$null-eq$cim.CreationDate){throw 'Current identity unreadable'}
  $state['ObservedName']=[IO.Path]::GetFileNameWithoutExtension([string]$cim.Name);$state['ObservedStartTicksUTC']=[string]$cim.CreationDate.ToUniversalTime().Ticks
  if($state.ObservedName-ieq$row.Name -and $state.ObservedStartTicksUTC-ceq$row.StartTicksUTC){$state.State='ExactOwnStillLive'}else{$state.State='DifferentIdentityPresent';$state.OwnIdentityTerminal=$true}
 }
 $observed+=,$state
}
$live=@($observed|Where-Object {-not$_.OwnIdentityTerminal})
$result=[ordered]@{AtUTC=[DateTimeOffset]::UtcNow.ToString('o');AllOwnClosed=($live.Count-eq0);Identities=$observed;AnyProcessClosed=$false;FullAcceptanceClaim=$false;PublicationClaim=$false}
$result|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'OwnClosuresBeforeFull01.json') -Encoding UTF8
[ordered]@{AllOwnClosed=$result.AllOwnClosed;Identities=$observed.Count;Pending=$live.Count;AnyProcessClosed=$false}|ConvertTo-Json
if($live.Count){throw 'Current own identity still live; no full launch allowed'}