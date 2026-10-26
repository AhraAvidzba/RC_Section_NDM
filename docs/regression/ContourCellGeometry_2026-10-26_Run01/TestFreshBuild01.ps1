# ДЛЯ ТЕСТОВ: полная сборка нового шаблона в изоляции, без перезаписи пользовательской книги.
param([string]$Label='FreshBuild01')
$ErrorActionPreference='Stop'
$evidenceOuter=$PSScriptRoot
$repoOuter=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path
Set-Location -LiteralPath $repoOuter
$memoryOuter=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
$ramOuter=[double]$memoryOuter.AvailableMBytes/1024
$commitOuter=([double]$memoryOuter.CommitLimit-[double]$memoryOuter.CommittedBytes)/1GB
if($ramOuter -lt 8 -or $commitOuter -lt 14){throw 'Insufficient combined reserve; no Excel started.'}
if($Label -notmatch '^FreshBuild[0-9]+$'){throw 'Invalid fresh build label.'}
$outputOuter='docs/regression/ContourCellGeometry_2026-10-26_Run01/'+$Label+'/RC_Section_NDM.xlsm'
$imagesOuter='docs/regression/ContourCellGeometry_2026-10-26_Run01/'+$Label+'/FormulaImages'
if(Test-Path -LiteralPath (Split-Path -Parent $outputOuter)){throw 'Fresh build directory required.'}
$builderScript=Join-Path $repoOuter 'tools/build_workbook/Build-Workbook.ps1'
$startedOuter=[DateTimeOffset]::Now.ToString('o')
. $builderScript -OutputPath $outputOuter -FormulaImageDirectory $imagesOuter
[ordered]@{Status='FRESH_BUILD_COMPLETED';Started=$startedOuter;Finished=[DateTimeOffset]::Now.ToString('o');PID=$ownedExcelProcessId;StartTicksUTC=[string]$ownedExcelStartTicks;FreePhysicalBeforeGiB=$ramOuter;FreeCommitBeforeGiB=$commitOuter;SourceBuildScriptSHA256=(Get-FileHash -LiteralPath $builderScript).Hash;BookSHA256=(Get-FileHash -LiteralPath $outputOuter).Hash}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $evidenceOuter ($Label+'.json')) -Encoding UTF8
