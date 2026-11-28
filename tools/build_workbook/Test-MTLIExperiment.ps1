# Адресные исследования MTLI и точности в частной нативной книге Excel.
# Исходная книга не открывается и не заменяется. Все собственные процессы
# идентифицируются по PID и времени старта; первичные журналы не переписываются.
param(
    [string]$EvidenceDirectory='docs/regression/MTLI_2026-11-28/Kernel_Run01',
    [string]$Macro='modTestThickLayerIntegration.RunThickLayerIntegrationTests',
    [switch]$Baseline,
    [string]$ReuseGeometryFrom='',
    [string]$SourceWorkbookPath='workbook/output/RC_Section_NDM.xlsm',
    [switch]$UseCompiledSource
)
$ErrorActionPreference='Stop'
$root=(Get-Location).Path
$dir=[IO.Path]::GetFullPath((Join-Path $root $EvidenceDirectory))
if(Test-Path -LiteralPath (Join-Path $dir 'Started.json')){throw 'Use a new evidence directory; primary run evidence already exists'}
New-Item -ItemType Directory -Path $dir -Force|Out-Null
$memory=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
if($memory.AvailableMBytes -lt 4096 -or ($memory.CommitLimit-$memory.CommittedBytes) -lt 7GB){throw 'Insufficient RAM/commit before private Excel'}
$memory|Select-Object AvailableMBytes,CommitLimit,CommittedBytes|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $dir 'Memory.json') -Encoding UTF8
$source=Join-Path $root $SourceWorkbookPath
$path=Join-Path $dir 'MTLI_Experimental.xlsm'
$sourceHash=(Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
if($Baseline){$Macro='modTestMTLIWorkflowBaseline.RunThickLayerWorkflowBaselineBenchmarkTests'}
Copy-Item -LiteralPath $source -Destination $path
if($ReuseGeometryFrom){
    $reuse=@()
    foreach($name in @('MTLI_Workflow_1.ndm','MTLI_Workflow_2.ndm')){
        $original=Join-Path ([IO.Path]::GetFullPath((Join-Path $root $ReuseGeometryFrom))) $name
        $target=Join-Path $dir $name
        Copy-Item -LiteralPath $original -Destination $target
        $reuse+=[ordered]@{Source=$original;Target=$target;SHA256=(Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash}
    }
    $reuse|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $dir 'GeometryReuse.json') -Encoding UTF8
}
Add-Type 'using System;using System.Runtime.InteropServices;public static class MTLITestWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint id);}'
$existing=@(Get-Process EXCEL -ErrorAction SilentlyContinue|ForEach-Object Id)
$excel=$null;$book=$null;$owned=0;$ticks=0
try{
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=1
    [uint32]$id=0;[void][MTLITestWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $owned=[int]$id;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
    if($owned -in $existing){throw 'Private Excel must not reuse an existing process'}
    [ordered]@{PID=$owned;StartTicksUTC=[string]$ticks;SourceSHA256=$sourceHash;Macro=$Macro;StartedUTC=[DateTimeOffset]::UtcNow.ToString('o')}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $dir 'Started.json') -Encoding UTF8
    Write-Output "OWNED_EXCEL: $owned; ticks=$ticks"
    $book=$excel.Workbooks.Open($path,0,$false)
    if($book.ReadOnly){throw 'Private workbook is read-only'}
    $importedSources=@()
    $sourcePaths=@('src/Solver/CThickLayerIntegrator.cls','src/Solver/CSectionSolver.cls','src/Solver/CCapacitySolver.cls','src/Solver/CStateSolutionRunner.cls','src/Solver/CSectionStrainLimits.cls','src/Solver/CCapacityCalculator.cls','src/Batch/CStateProvider.cls','src/Batch/CBatchSectionCalculator.cls','src/Crack/CCrackFormationCalculator.cls','src/Excel/CSystemSettingsReader.cls','tests/modTestThickLayerIntegration.bas')
    if($Baseline){$sourcePaths=@('tests/modTestThickLayerIntegration.bas')}
    if($UseCompiledSource){$sourcePaths=@()}
    foreach($relative in $sourcePaths){
        $sourceText=[IO.File]::ReadAllText((Join-Path $root $relative),[Text.Encoding]::UTF8)
        $name=[regex]::Match($sourceText,'Attribute VB_Name = "([^"]+)"').Groups[1].Value
        if($Baseline){
            $header=[regex]::Match($sourceText,'(?ms)^#If VBA7 Then.*?^#End If').Value
            $bodyParts=@('Option Explicit',"' ДЛЯ ТЕСТОВ: единая проба пользовательской команды на релизном VBA.",$header,'Private mPassed As Long, mFailed As Long, mReport As String')
            foreach($method in @('RunThickLayerWorkflowBaselineBenchmarkTests','TestWorkflowSeries','TestWorkflowCrackedLoads','TestCaptureWorkflowSnapshot','TestWorkbookSetting','TestWorkflowNow','TestNumber','TestTrue')){
                $matches=[regex]::Matches($sourceText,'(?ms)^(?:Public|Private) (?:Function|Sub) '+$method+'\b.*?^End (?:Function|Sub)\r?$')
                if($matches.Count -ne 1){throw ('Missing unique workflow method '+$method)}
                $bodyParts+=$matches[0].Value
            }
            $name='modTestMTLIWorkflowBaseline'
            $sourceText='Attribute VB_Name = "'+$name+'"'+"`r`n"+($bodyParts -join "`r`n`r`n")
            [IO.File]::WriteAllText((Join-Path $dir 'modTestMTLIWorkflowBaseline.bas'),$sourceText,[Text.UTF8Encoding]::new($true))
            $Macro='modTestMTLIWorkflowBaseline.RunThickLayerWorkflowBaselineBenchmarkTests'
        }
        $type=1;if($relative.EndsWith('.cls')){$type=2}
        $component=$null
        try{$component=$book.VBProject.VBComponents.Item($name)}catch{}
        if(-not $component){$component=$book.VBProject.VBComponents.Add($type);$component.Name=$name}
        if($component.CodeModule.CountOfLines -gt 0){$component.CodeModule.DeleteLines(1,$component.CodeModule.CountOfLines)}
        $body=$sourceText -replace '(?ms)\AVERSION.*?^END\r?\n','' -replace '(?m)^Attribute VB_[^\r\n]*\r?\n',''
        $component.CodeModule.AddFromString($body)
        $importedSources += [ordered]@{Path=$relative;SHA256=(Get-FileHash -LiteralPath (Join-Path $root $relative) -Algorithm SHA256).Hash}
    }
    $importedSources|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $dir 'SourceHashes.json') -Encoding UTF8
    . (Join-Path $root 'tools/build_workbook/SettingsCatalog.ps1')
    $layout=Get-Content -LiteralPath (Join-Path $root 'tools/build_workbook/WorkbookLayout.json') -Raw -Encoding UTF8|ConvertFrom-Json
    if(-not $UseCompiledSource){Add-ExperimentalNumericsSettings $book $layout}
    $excel.VBE.ActiveVBProject=$book.VBProject
    $book.Save()
    $compile=$excel.VBE.CommandBars.FindControl(1,578)
    if($compile.Enabled){$compile.Execute()}
    if($compile.Enabled){
        $pane=$excel.VBE.ActiveCodePane
        [int]$startLine=0;[int]$startColumn=0;[int]$endLine=0;[int]$endColumn=0
        $pane.GetSelection([ref]$startLine,[ref]$startColumn,[ref]$endLine,[ref]$endColumn)
        [ordered]@{Module=$pane.CodeModule.Name;Line=$startLine;Column=$startColumn;Code=$pane.CodeModule.Lines([Math]::Max(1,$startLine-2),5)}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $dir 'CompileFailure.json') -Encoding UTF8
        throw ('VBA compilation failed: '+$pane.CodeModule.Name+':'+$startLine)
    }
    Write-Output ('COMPILED: '+$book.VBProject.VBComponents.Count)
    $timer=[Diagnostics.Stopwatch]::StartNew()
    $report=[string]$excel.Run("'"+$book.Name+"'!"+$Macro)
    $timer.Stop()
    [IO.File]::WriteAllText((Join-Path $dir 'Native.txt'),$report,[Text.UTF8Encoding]::new($true))
    $failed=$report -match '(?m)^FAIL:|RUNTIME ERROR|failed=[1-9]'
    $hasTotal=$report -match 'TOTAL_MTLI_'
    $book.Save()
    $text=[Text.StringBuilder]::new()
    foreach($component in $book.VBProject.VBComponents){
        [void]$text.AppendLine('='*100);[void]$text.AppendLine('COMPONENT: '+$component.Name)
        $lines=$component.CodeModule.CountOfLines
        [void]$text.AppendLine('TYPE: '+$component.Type);[void]$text.AppendLine('LINES: '+$lines);[void]$text.AppendLine('='*100)
        if($lines -gt 0){[void]$text.AppendLine($component.CodeModule.Lines(1,$lines))}
    }
    [IO.File]::WriteAllText((Join-Path $dir 'NativeExport.txt'),$text.ToString(),[Text.UTF8Encoding]::new($true))
    $componentCount=$book.VBProject.VBComponents.Count
    $book.Close($false);$book=$null
    [ordered]@{Status=$(if($failed -or -not $hasTotal){'FAIL'}else{'PASS'});Compiled=$true;Components=$componentCount;BaselineRelease=[bool]$Baseline;Macro=$Macro;ElapsedSec=$timer.Elapsed.TotalSeconds;SourceSHA256=$sourceHash;CandidateSHA256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash;UserWorkbookOpened=$false;Report='Native.txt'}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $dir 'Summary.json') -Encoding UTF8
    $report -split '\r?\n'|Where-Object{$_ -match '^(TOTAL_|FAIL:|VALUES:|RUNTIME ERROR)'}|Write-Output
    if($failed -or -not $hasTotal){throw 'Directed MTLI checks failed; primary evidence preserved'}
}catch{
    [ordered]@{Status='HARNESS_OR_TEST_FAILURE';Error=$_.Exception.Message;Stack=$_.ScriptStackTrace;UTC=[DateTimeOffset]::UtcNow.ToString('o')}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $dir 'Failure.json') -Encoding UTF8
    throw
}finally{
    if($book){$book.Close($false)}
    if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($owned -and $owned -notin $existing){$p=Get-Process -Id $owned -ErrorAction SilentlyContinue;if($p -and $p.StartTime.ToUniversalTime().Ticks -eq $ticks){Stop-Process -Id $owned -Force}}
}
if((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -ne $sourceHash){throw 'Source workbook changed during experiment'}
