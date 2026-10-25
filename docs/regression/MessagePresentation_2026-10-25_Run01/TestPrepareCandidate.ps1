# ДЛЯ ТЕСТОВ: компилирует только изменённые модули на собственной копии книги.
# Исходный пакет и пользовательские процессы не меняются; сохраняет код до/после.
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path
$source=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$sourceHash=(Get-FileHash -LiteralPath $source).Hash
$memory=Get-CimInstance Win32_OperatingSystem
if($memory.FreeVirtualMemory*1KB -lt 2GB -or $memory.FreePhysicalMemory*1KB -lt 2GB){throw 'Memory reserve insufficient; no Excel started.'}
$directory=Join-Path $PSScriptRoot 'Candidate03'
if(Test-Path -LiteralPath $directory){throw 'Fresh candidate directory required.'}
[void][IO.Directory]::CreateDirectory($directory)
$path=Join-Path $directory 'RC_Section_NDM.xlsm'
Copy-Item -LiteralPath $source -Destination $path
$changed=@{
    modWorkbookMessages='src/Excel/modWorkbookMessages.bas'
    modWorkbookCalculation='src/Excel/modWorkbookCalculation.bas'
    modTestAudit02Api='tests/modTestAudit02Api.bas'
    'ЭтаКнига'='tools/build_workbook/WorkbookHelpEvents.vba'
}
$beforePids=@(Get-Process EXCEL -ErrorAction SilentlyContinue|ForEach-Object Id)
Add-Type 'using System;using System.Runtime.InteropServices;public static class NdmMessageTestOwner{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
$app=$null;$book=$null;$owned=0;$ticks=0
try{
    $app=New-Object -ComObject Excel.Application
    [uint32]$nativeId=0;[void][NdmMessageTestOwner]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$nativeId)
    $owned=[int]$nativeId;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
    if($owned -in $beforePids){throw 'Own Excel identity required.'}
    [ordered]@{PID=$owned;StartTicks=[string]$ticks;SourceSHA256=$sourceHash;FreeCommitBeforeGiB=[math]::Round($memory.FreeVirtualMemory*1KB/1GB,2)}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'CandidateOwnership.json') -Encoding UTF8
    $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=3
    $book=$app.Workbooks.Open($path,0,$false)
    $before=[ordered]@{}
    foreach($component in $book.VBProject.VBComponents){$count=$component.CodeModule.CountOfLines;$before[$component.Name]=if($count){$component.CodeModule.Lines(1,$count)}else{''}}
    $before|ConvertTo-Json -Depth 3|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'BeforeNativeCode.json') -Encoding UTF8
    foreach($name in $changed.Keys){
        $body=[regex]::Match([IO.File]::ReadAllText((Join-Path $root $changed[$name])),'(?ms)^Option Explicit.*').Value
        if(-not $body){throw ('Source body missing: '+$name)}
        $module=$book.VBProject.VBComponents.Item($name).CodeModule
        $module.DeleteLines(1,$module.CountOfLines);$module.AddFromString(($body -split '\r?\n') -join "`r`n")
    }
    $app.VBE.ActiveVBProject=$book.VBProject
    $compile=$app.VBE.CommandBars.FindControl(1,578)
    if(-not $compile){throw 'Compile command unavailable.'}
    if($compile.Enabled){$compile.Execute()}
    if($compile.Enabled){throw 'Compilation failed.'}
    $after=[ordered]@{}
    foreach($component in $book.VBProject.VBComponents){$count=$component.CodeModule.CountOfLines;$after[$component.Name]=if($count){$component.CodeModule.Lines(1,$count)}else{''}}
    $after|ConvertTo-Json -Depth 3|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'CandidateNativeCode.json') -Encoding UTF8
    $differences=@($before.Keys|Where-Object {$before[$_] -cne $after[$_]})
    # VBE канонизирует регистр идентификаторов во всём проекте при компиляции.
    # Строковые литералы и комментарии не нормализуются: они должны совпасть точно.
    $caseOnly=@()
    foreach($name in $differences){
        if($changed.ContainsKey($name)){continue}
        $pattern='"(?:[^"\r\n]|"")*"|''[^\r\n]*|[A-Za-z_][A-Za-z_0-9]*'
        $normalizer=[Text.RegularExpressions.MatchEvaluator]{param($m) if($m.Value.StartsWith('"') -or $m.Value.StartsWith("'")){$m.Value}else{$m.Value.ToLowerInvariant()}}
        if([regex]::Replace($before[$name],$pattern,$normalizer) -cne [regex]::Replace($after[$name],$pattern,$normalizer)){throw ('Unexpected native code change: '+$name)}
        $caseOnly+=$name
    }
    if($before.Count -ne 178 -or $after.Count -ne 178 -or @($changed.Keys|Where-Object {$_ -notin $differences}).Count){throw 'Unexpected native component changes.'}
    $book.Save();$book.Close($false);$book=$null
    [ordered]@{Status='COMPILED_ISOLATED';SourceSHA256=$sourceHash;CandidateSHA256=(Get-FileHash -LiteralPath $path).Hash;ChangedComponents=@($changed.Keys);CompilerIdentifierCaseOnly=$caseOnly;OtherwiseUnchangedComponents=174;FullPass=$false}|ConvertTo-Json -Depth 4|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'CandidatePreparation.json') -Encoding UTF8
}finally{
    $process=Get-Process -Id $owned -ErrorAction SilentlyContinue
    if($process -and $process.ProcessName -eq 'EXCEL' -and $process.StartTime.ToUniversalTime().Ticks -eq $ticks -and $owned -notin $beforePids){if($book){$book.Close($false)};if($app){$app.Quit()}}
    if($app){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
if((Get-FileHash -LiteralPath $source).Hash -cne $sourceHash){throw 'Published source workbook changed.'}
Get-Content -LiteralPath (Join-Path $PSScriptRoot 'CandidatePreparation.json') -Raw
