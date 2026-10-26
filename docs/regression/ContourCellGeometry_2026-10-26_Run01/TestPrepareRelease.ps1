# ДЛЯ ТЕСТОВ: обновляет справку и статические заголовки только в отдельной копии кандидата.
param([string]$SourceRelative='Compile08/RC_Section_NDM.xlsm',[string]$Label='ReleaseCandidate01')
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSHOME 'Modules/Microsoft.PowerShell.Utility/Microsoft.PowerShell.Utility.psd1') -Global
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path
$source=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot $SourceRelative)).Path
$directory=Join-Path $PSScriptRoot $Label
if(Test-Path -LiteralPath $directory){throw 'Use a fresh label.'}
[void][IO.Directory]::CreateDirectory($directory)
$candidate=Join-Path $directory 'RC_Section_NDM.xlsm'
$sourceHash=(Get-FileHash -LiteralPath $source).Hash
Copy-Item -LiteralPath $source -Destination $candidate
$memory=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
$freeRAM=[double]$memory.AvailableMBytes/1024
$freeCommit=([double]$memory.CommitLimit-[double]$memory.CommittedBytes)/1GB
if($freeRAM -lt 4 -or $freeCommit -lt 8){throw 'Insufficient memory reserve; no Excel started.'}
. (Join-Path $root 'tools/build_workbook/SettingsCatalog.ps1')
. (Join-Path $root 'tools/build_workbook/SectionRecipeGuide.ps1')
$script:FormulaImageDir=Join-Path $directory 'FormulaImages'
Add-Type 'using System;using System.Runtime.InteropServices;public static class NdmContourReleaseOwner{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
$prior=@(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$app=$null;$book=$null;$owned=0;$ticks=0L;$failure=$null
try{
    $app=New-Object -ComObject Excel.Application
    [uint32]$native=0;[void][NdmContourReleaseOwner]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$native)
    $owned=[int]$native;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
    if($owned -in $prior){throw 'New owned Excel required.'}
    [ordered]@{PID=$owned;StartTicksUTC=[string]$ticks;SourceSHA256=$sourceHash;FreeRAMGiB=$freeRAM;FreeCommitGiB=$freeCommit}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $directory 'Ownership.json') -Encoding UTF8
    $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=3
    $book=$app.Workbooks.Open($candidate,0,$false)
    if($book.ReadOnly){throw 'Private output must be writable.'}
    $config=$book.Worksheets.Item('Config');$help=$book.Worksheets.Item('Справка')
    $settings=$book.Names.Item('rngSystemSettings').RefersToRange
    $catalog=@{}
    foreach($section in (Get-SystemSettingsCatalog)){foreach($row in $section.Rows){$catalog[[string]$row[0]]=$row}}
    foreach($key in @('Mesh.CellGeometry','Mesh.StepMode','Mesh.BoundarySubdivisions')){
        $found=0
        for($i=1;$i -le $settings.Rows.Count;$i++){
            if([string]$settings.Cells.Item($i,1).Value2 -ceq $key){
                $settings.Cells.Item($i,4).Value2=[string]$catalog[$key][3];$found++
            }
        }
        if($found -ne 1){throw ('Expected one setting '+$key)}
    }
    Add-SettingsInstructions $book $config $help
    # Справка только что создана заново: старое имя её последнего раздела
    # больше не относится к прежнему тексту. Штатный сборщик добавляет раздел с нуля.
    try{$book.Names.Item('rngSectionRecipeGuide').Delete()}catch{}
    Update-SectionRecipeGuide $book ([string]$root) | Out-Null
    Apply-ConfigNamedRangeBorders $book
    Fit-SettingsInstructionRows $help
    Set-HelpSectionOrderAndOutline $book -Expanded | Out-Null
    $results=$book.Worksheets.Item('Results')
    $results.Range('CR155:DI155').Merge()
    $results.Cells.Item(155,96).Value2='Фактические границы бетонных элементов'
    $results.Range('CR155:DI156').Font.Bold=$true
    foreach($edge in 7,8,9,10,11,12){$border=$results.Range('CR156:DI156').Borders.Item($edge);$border.LineStyle=1;$border.Weight=2;$border.Color=12632256}
    $book.Save()
}catch{$failure=$_.Exception.ToString();throw}finally{
    $process=Get-Process -Id $owned -ErrorAction SilentlyContinue
    if($process -and $process.ProcessName -eq 'EXCEL' -and $process.StartTime.ToUniversalTime().Ticks -eq $ticks -and $owned -notin $prior){
        if($book){$book.Close($false)}
        if($app){$app.Quit()}
    }
    $book=$null
    if($app){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
$sourceUnchanged=(Get-FileHash -LiteralPath $source).Hash -ceq $sourceHash
[ordered]@{Status='PREPARED_REQUIRES_ACCEPTANCE';SourceSHA256=$sourceHash;SourceUnchanged=$sourceUnchanged;
    CandidateSHA256=(Get-FileHash -LiteralPath $candidate).Hash;SourceVBAReimported=$false;Failure=$failure;FullRegressionPass=$false}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $directory 'Preparation.json') -Encoding UTF8
if(-not $sourceUnchanged){throw 'Source candidate changed.'}
