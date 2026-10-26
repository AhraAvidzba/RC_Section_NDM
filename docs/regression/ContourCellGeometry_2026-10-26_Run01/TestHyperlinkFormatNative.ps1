# ДЛЯ ТЕСТОВ: повторно создаёт существующие ссылки в собственной копии без VBA и расчёта.
param([string]$SourceRelative='ReleaseCandidate04/RC_Section_NDM.xlsm',[string]$Label='HyperlinkNative01',[switch]$SaveOnly)
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSHOME 'Modules/Microsoft.PowerShell.Utility/Microsoft.PowerShell.Utility.psd1') -Global
$repo=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path
$source=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot $SourceRelative)).Path
$directory=Join-Path $PSScriptRoot $Label
if(Test-Path -LiteralPath $directory){throw 'Fresh label required.'}
[void][IO.Directory]::CreateDirectory($directory)
$target=Join-Path $directory 'RC_Section_NDM.xlsm'
$hash=(Get-FileHash -LiteralPath $source).Hash
Copy-Item -LiteralPath $source -Destination $target
$memory=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
if($memory.AvailableMBytes -lt 4096 -or ([double]$memory.CommitLimit-[double]$memory.CommittedBytes) -lt 8GB){throw 'Insufficient memory; no Excel started.'}
. (Join-Path $repo 'tools/build_workbook/SettingsCatalog.ps1')
Add-Type 'using System;using System.Runtime.InteropServices;public static class NdmLinkSmokeOwner{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
$prior=@(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$app=$null;$book=$null;$owned=0;$ticks=0L;$links=@();$failure=$null
try{
    $app=New-Object -ComObject Excel.Application
    [uint32]$native=0;[void][NdmLinkSmokeOwner]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$native)
    $owned=[int]$native;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
    if($owned -in $prior){throw 'New own Excel required.'}
    [ordered]@{PID=$owned;StartTicksUTC=[string]$ticks;SourceSHA256=$hash;MacrosEnabled=$false}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $directory 'Ownership.json') -Encoding UTF8
    $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=3
    $book=$app.Workbooks.Open($target,0,$false)
    if($book.ReadOnly){throw 'Own copy must be writable.'}
    foreach($sheetName in @('Config','Сечения')){
        $sheet=$book.Worksheets.Item($sheetName)
        foreach($link in $sheet.Hyperlinks){
            $sub=[string]$link.SubAddress
            if($sub -notmatch "^'?Справка'?!A([0-9]+)$"){continue}
            $row=[int]$Matches[1];$cell=$link.Range.Cells.Item(1,1);$text=[string]$cell.Value2
            if($text -ne 'Подробнее' -and -not $text.EndsWith(' (Подробнее)')){throw ('Unexpected linked caption '+$text)}
            $links+=,[pscustomobject]@{Sheet=$sheetName;Address=$cell.Address();TargetRow=$row;Text=$text}
        }
    }
    foreach($record in $(if($SaveOnly){@()}else{$links})){
        $cell=$book.Worksheets.Item($record.Sheet).Range($record.Address)
        if($record.Text -ceq 'Подробнее'){Add-InstructionHyperlink $cell 'Справка' $record.TargetRow}
        else{Add-HeaderInstructionHyperlink $cell $record.Text.Substring(0,$record.Text.Length-' (Подробнее)'.Length) 'Справка' $record.TargetRow}
    }
    if($links.Count -lt 100){throw 'Incomplete hyperlink inventory.'}
    $book.Save()
}catch{
    $failure=$_.Exception.ToString()
    [ordered]@{Status='FAIL';Reason=$failure;ScriptStack=$_.ScriptStackTrace}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $directory 'Failure.json') -Encoding UTF8
    throw
}finally{
    $process=Get-Process -Id $owned -ErrorAction SilentlyContinue
    if($process -and $process.StartTime.ToUniversalTime().Ticks -eq $ticks -and $owned -notin $prior){if($book){$book.Close($false)};if($app){$app.Quit()}}
    $book=$null;if($app){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
[ordered]@{Status='NATIVE_ACTIONS_COMPLETED_XML_REVIEW_REQUIRED';SaveOnly=[bool]$SaveOnly;Links=$links.Count;Records=$links;SourceUnchanged=((Get-FileHash -LiteralPath $source).Hash -ceq $hash);
    OutputSHA256=(Get-FileHash -LiteralPath $target).Hash;SourceSHA256=$hash;VBAExecuted=$false;Failure=$failure}|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $directory 'NativeActions.json') -Encoding UTF8
