# Адресно проверяет начальную служебную таблицу нового шаблона и точечные тексты; только собственная копия.
$ErrorActionPreference='Stop'
$root=(Get-Location).Path
$directory=Join-Path $root 'docs/regression/ExcelGuardGuide_2026-10-13/NoMigrations_2026-10-15'
. (Join-Path $root 'tools/build_workbook/SettingsCatalog.ps1')
$copy=Join-Path $directory 'BootstrapProbe.xlsm'
if(Test-Path -LiteralPath $copy){throw 'Do not replace an existing probe.'}
Copy-Item -LiteralPath (Join-Path $directory 'Candidate.xlsm') -Destination $copy
$protected=@('workbook/output/RC_Section_NDM.xlsm','workbook/output/VBA_All_Code.txt','workbook/output/RC_Section_NDM_execution_report.txt')
$before=@{};foreach($path in $protected){$before[$path]=(Get-FileHash -LiteralPath $path).Hash}
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class BootstrapExcelIdentity{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hWnd,out uint id);}'
$excel=$null;$book=$null;$ownedPid=0;$ownedTicks=0;$success=$false
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=1
    [uint32]$id=0;[void][BootstrapExcelIdentity]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $ownedPid=[int]$id;$ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    Write-Output ('OWNED_EXCEL: '+$ownedPid)
    $book=$excel.Workbooks.Open($copy,0,$false);$excel.Calculation=-4135
    $texts=Get-Content -LiteralPath (Join-Path $directory 'PresentationText.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $checked=0
    foreach($property in $texts.PSObject.Properties){
        $parts=$property.Name.Split('!')
        if([string]$book.Worksheets.Item($parts[0]).Range($parts[1]).Value2 -cne [string]$property.Value){throw ('Saved text mismatch: '+$property.Name)}
        $checked++
    }
    $help=[string]$book.Worksheets.Item('Справка').Range('A576').Value2
    if($help -notmatch 'UpdateSectionRecipeTypes' -or $help -match 'UpdateSectionRecipeConfig'){throw 'Saved short help is stale.'}
    $config=$book.Worksheets.Item('Config');$range=$config.UsedRange
    $values=ConvertTo-Json -InputObject $range.Formula -Compress -Depth 4
    $names=@{};foreach($name in $book.Names){$names[[string]$name.Name]=[string]$name.RefersTo}
    # Сборщик нового шаблона создаёт именованный источник; рабочий updater его только обновляет.
    $ids=@($excel.Run("'$($book.Name)'!modSectionRecipeSetup.RegisteredSectionRecipeTypes"))
    $book.Names.Item('rngSectionRecipeTypes').Delete()
    Reset-ConfigValidationSources $config
    $script:ConfigValidationNextColumn=$range.Column+$range.Columns.Count
    Get-ConfigValidationListAddress $config $ids 'rngSectionRecipeTypes' | Out-Null
    $excel.Run("'$($book.Name)'!modSectionRecipeSetup.UpdateSectionRecipeTypes")
    $list=$book.Names.Item('rngSectionRecipeTypes').RefersToRange
    if($list.Rows.Count -ne $ids.Count){throw 'Initial registry list count differs.'}
    for($index=0;$index -lt $ids.Count;$index++){if([string]$list.Cells.Item($index+1,1).Value2 -cne [string]$ids[$index]){throw 'Initial registry list differs.'}}
    if((ConvertTo-Json -InputObject $range.Formula -Compress -Depth 4) -cne $values){throw 'Unrelated existing data changed.'}
    foreach($name in $book.Names){if($name.Name -ne 'rngSectionRecipeTypes' -and $names[[string]$name.Name] -cne [string]$name.RefersTo){throw ('Unrelated name changed: '+$name.Name)}}
    $success=$true
    [ordered]@{Passed=$true;TextsChecked=$checked+1;RegistryTypes=$ids;BootstrapMatchesBuilder=$true;OtherConfigDataAndNamesUnchanged=$true;ProductionUnchanged=$true} |
        ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $directory 'BootstrapAndText.json') -Encoding UTF8
    Write-Output 'BOOTSTRAP_AND_TEXT_PASS'
} finally {
    if($book){try{$book.Close($false)}finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}}
    if($excel){try{$excel.Quit()}finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($ownedPid){$owned=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue;if($owned -and $owned.ProcessName -eq 'EXCEL' -and $owned.StartTime.ToUniversalTime().Ticks -eq $ownedTicks){Stop-Process -Id $ownedPid -Force}}
    foreach($path in $before.Keys){if((Get-FileHash -LiteralPath $path).Hash -cne $before[$path]){throw ('Protected file changed: '+$path)}}
}
