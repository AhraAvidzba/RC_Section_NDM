# -------------------- ДЛЯ ТЕСТОВ --------------------
# Собирает чистый шаблон из актуальных исходников; расчёт и автоматические поведенческие тесты не запускает.
$ErrorActionPreference='Stop'
$repoOuter=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path
Set-Location -LiteralPath $repoOuter
# Перенос разрешён только после чтения всех журналов и выхода всех собственных процессов первого набора.
$previousLaunch=Join-Path $repoOuter 'docs/regression/FullSuiteLaunch_2026-10-28_PlotLabels_Run01'
$closure=Get-Content -LiteralPath (Join-Path $previousLaunch 'AllOwnClosuresReviewed.json') -Raw -Encoding UTF8|ConvertFrom-Json
if(-not $closure.AllOwnIdentitiesTerminal -or -not $closure.PrimaryLogsReviewed -or $closure.CoreMacrosReviewed -ne182 -or $closure.FocusedSingleCasesReviewed -ne122 -or $closure.FocusedRecordCount -ne26){throw 'Previous full inventory/own closures incomplete; no build permitted'}
$transfer=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'ParallelFix01/SourceTransfer01.json') -Raw -Encoding UTF8|ConvertFrom-Json
if($transfer.Status -cne 'VERIFIED_TEST_ONLY_PATCH_TRANSFERRED_AFTER_ALL_FULL_OWN_EXITED' -or (Get-FileHash -LiteralPath (Join-Path $repoOuter 'tests/modTestWorkbookInterface.bas')).Hash -cne $transfer.CandidateSHA256){throw 'Exact verified test-only source transfer required'}
$outOuter=Join-Path $PSScriptRoot 'FreshBuild03'
$timerTransfer=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'ParallelFix02/SourceTransfer01.json') -Raw -Encoding UTF8|ConvertFrom-Json
if($timerTransfer.Status -cne 'TEST_TIMER_TRANSFERRED_AFTER_ALL_FULL_OWN_EXITED' -or (Get-FileHash -LiteralPath (Join-Path $repoOuter $timerTransfer.Source)).Hash -cne $timerTransfer.CandidateSHA256){throw 'Exact test timer correction required'}
$captionTransfer=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'ParallelFix03/SourceTransfer01.json') -Raw -Encoding UTF8|ConvertFrom-Json
if($captionTransfer.Status -cne 'TEST_CAPTION_TRANSFERRED_AFTER_ALL_FULL_OWN_EXITED' -or (Get-FileHash -LiteralPath (Join-Path $repoOuter $captionTransfer.Source)).Hash -cne $captionTransfer.CandidateSHA256){throw 'Exact current-caption test correction required'}
if(Test-Path -LiteralPath $outOuter){throw 'Fresh build directory required'}
[void][IO.Directory]::CreateDirectory($outOuter)
$memoryOuter=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
if($memoryOuter.AvailableMBytes-lt8192-or($memoryOuter.CommitLimit-$memoryOuter.CommittedBytes)-lt14GB){throw 'RAM/commit reserve insufficient; no Excel started'}
[ordered]@{AtUTC=[DateTimeOffset]::UtcNow.ToString('o');RAMGiB=$memoryOuter.AvailableMBytes/1024;CommitFreeGiB=($memoryOuter.CommitLimit-$memoryOuter.CommittedBytes)/1GB}|ConvertTo-Json|Set-Content (Join-Path $outOuter 'MemoryBefore.json') -Encoding UTF8
$startedOuter=[DateTimeOffset]::UtcNow.ToString('o')
. (Join-Path $repoOuter 'tools/build_workbook/Build-Workbook.ps1') -OutputPath 'docs/regression/PlotLabelsPerformance_2026-10-28_Run01/FreshBuild03/RC_Section_NDM.xlsm' -FormulaImageDirectory 'docs/regression/PlotLabelsPerformance_2026-10-28_Run01/FreshBuild03/FormulaImages'
[ordered]@{Status='CLEAN_BUILD_COMPILED';StartedUTC=$startedOuter;FinishedUTC=[DateTimeOffset]::UtcNow.ToString('o');PID=$ownedExcelProcessId;StartTicksUTC=[string]$ownedExcelStartTicks;BookSHA256=(Get-FileHash -LiteralPath (Join-Path $outOuter 'RC_Section_NDM.xlsm')).Hash;CompileCommand=578;UserDataCopied=$false;CalculationsOrBehaviorTestsRun=$false;PlotterSHA256=(Get-FileHash -LiteralPath (Join-Path $repoOuter 'src/Excel/CSectionPlotter.cls')).Hash}|ConvertTo-Json|Set-Content (Join-Path $outOuter 'Build.json') -Encoding UTF8
