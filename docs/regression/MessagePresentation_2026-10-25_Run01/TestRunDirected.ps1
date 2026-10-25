# Directed message checks, one owned test Excel at a time; no full UI suite.
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path
$macro='modTestAudit02Api.RunProgramMessagePresentationTests,modTestAudit02Api.RunProgramMessageTests,modTestGeometryReview.RunCollapsedInputMessageTests,modTestGeometryReview.RunCollapsedInputAddressTests,modTestWorkbookInterface.RunAudit03SettingErrorMessageTests,modTestWorkbookInterface.RunSectionWorkflowCalculationTests'
$observations=@()
foreach($mode in @('No','Yes')){
    $memory=Get-CimInstance Win32_OperatingSystem
    $observation=[ordered]@{Mode=$mode;At=(Get-Date).ToString('o');FreePhysicalGiB=[math]::Round($memory.FreePhysicalMemory*1KB/1GB,3);FreeCommitGiB=[math]::Round($memory.FreeVirtualMemory*1KB/1GB,3)}
    $observations+=$observation
    $observations|ConvertTo-Json -Depth 3|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'DirectedMemory.json') -Encoding UTF8
    if($memory.FreePhysicalMemory*1KB -lt 2GB -or $memory.FreeVirtualMemory*1KB -lt 2GB){throw 'Insufficient reserve; no next Excel started.'}
    & (Join-Path $PSHOME 'powershell.exe') -NoProfile -File (Join-Path $root 'tools/build_workbook/Run-Audit03Watchdog.ps1') -SourceWorkbook (Join-Path $PSScriptRoot 'Candidate03/RC_Section_NDM.xlsm') -ReportPath ('docs/regression/MessagePresentation_2026-10-25_Run01/Directed_Final_'+$mode+'.txt') -Macro $macro -Mode $mode -TimeoutSeconds 300
    if($LASTEXITCODE -ne 0){throw ('Directed checks failed: '+$mode)}
}
