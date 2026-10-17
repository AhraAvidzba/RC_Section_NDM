# Удаляет только собственные завершённые временные копии по журналам этих запусков.
$ErrorActionPreference='Stop'
$root=(Get-Location).Path
$directories=@('docs/regression/FocusedAcceptance_2026-10-17_Run03','docs/regression/FocusedAcceptance_2026-10-17_Run04','docs/regression/FocusedAcceptance_2026-10-17_Final')
$tempRoot=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
$targets=@{}
foreach($directory in $directories){foreach($file in Get-ChildItem -LiteralPath (Join-Path $root $directory) -Filter '*.txt' -File){
 $body=Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8
 $fixture=[regex]::Match($body,'(?m)^FIXTURE: (.+)\r?$')
 if(-not $fixture.Success){continue}
 $path=[IO.Path]::GetFullPath((Split-Path -Parent $fixture.Groups[1].Value.Trim()))
 $identity=[regex]::Match($body,'TEST_EXCEL_PROCESS: id=(\d+); startTicks=(\d+)')
 if(-not $identity.Success){throw 'Owned Excel identity missing.'}
 $targets[$path]=@{Path=$path;PID=[int]$identity.Groups[1].Value;StartTicks=[long]$identity.Groups[2].Value;Process='EXCEL';Pattern='^RC_NDM_ModeSuite_[a-f0-9]{32}$'}
}}
$cadIdentity=Join-Path $root 'docs/regression/FocusedAcceptance_2026-10-17_Run04/NativeCAD_Diagnostic/RunRealAutoCADFarOriginTests/NativeCADOwnership.json'
if(Test-Path -LiteralPath $cadIdentity){$identity=Get-Content -LiteralPath $cadIdentity -Raw -Encoding UTF8|ConvertFrom-Json;if($identity.BootstrapDirectory){$path=[IO.Path]::GetFullPath($identity.BootstrapDirectory);$targets[$path]=@{Path=$path;PID=[int]$identity.OwnedPID;StartTicks=[long]$identity.StartTicks;Process='acad';Pattern='^NDM_NativeBootstrap_[a-f0-9]{32}$'}}}
$plan=@()
foreach($item in $targets.Values){
 if(-not $item.Path.StartsWith($tempRoot,[StringComparison]::OrdinalIgnoreCase) -or (Split-Path -Leaf $item.Path) -notmatch $item.Pattern){throw ('Unsafe cleanup target: '+$item.Path)}
 if(-not (Test-Path -LiteralPath $item.Path)){continue}
 $folder=Get-Item -LiteralPath $item.Path -Force
 if($folder.Attributes -band [IO.FileAttributes]::ReparsePoint){throw 'Reparse target refused.'}
 $children=@(Get-ChildItem -LiteralPath $item.Path -Force -Recurse)
 if(@($children|Where-Object {$_.Attributes -band [IO.FileAttributes]::ReparsePoint}).Count){throw 'Reparse child refused.'}
 $bytes=($children|Where-Object {-not $_.PSIsContainer}|Measure-Object Length -Sum).Sum
 $plan+=@{Path=$item.Path;Bytes=[long]$bytes;PID=$item.PID;StartTicks=$item.StartTicks;Process=$item.Process}
}
$plan|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'CleanupPlan.json') -Encoding UTF8
$removed=@();$stopped=@()
foreach($item in $plan){
 $process=Get-Process -Id $item.PID -ErrorAction SilentlyContinue
 if($process -and $process.ProcessName -eq $item.Process -and $process.StartTime.ToUniversalTime().Ticks -eq $item.StartTicks){Stop-Process -Id $item.PID -Force;$stopped+=$item.PID}
 $resolved=(Resolve-Path -LiteralPath $item.Path).Path
 if(-not $resolved.StartsWith($tempRoot,[StringComparison]::OrdinalIgnoreCase) -or $resolved -cne $item.Path){throw 'Resolved path changed.'}
 Remove-Item -LiteralPath $resolved -Recurse -Force
 $removed+=$item
}
[long]$freed=0
foreach($item in $removed){$freed+=[long]$item.Bytes}
@{Removed=$removed.Count;BytesFreed=$freed;StoppedOwnedPIDs=$stopped;NoGlobalProcessTermination=$true;OriginalWorkbooksPreserved=$true}|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'Cleanup.json') -Encoding UTF8
Get-Content -LiteralPath (Join-Path $PSScriptRoot 'Cleanup.json')
