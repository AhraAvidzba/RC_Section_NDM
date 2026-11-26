# ДЛЯ ТЕСТОВ: сравнивает исходную и измененную книги на сохраненной геометрии; все открытия ReadOnly.
param([switch]$PublishReport,[ValidateSet('Baseline','Candidate')][string[]]$Labels=@('Baseline','Candidate'),
    [ValidateSet('UserExample','TOnlyBlankProfile','SP63Stability','SP35Width','SP63Width','CompressedNMT')]
    [string[]]$TestCases=@('UserExample','TOnlyBlankProfile','SP63Stability','SP35Width','SP63Width','CompressedNMT'))
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../../..')).Path
Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class ReportExampleTestProcess { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint id); }'
function Assert-ReportTestMemory {
    $os=Get-CimInstance Win32_OperatingSystem;$memory=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
    if($os.FreePhysicalMemory/1MB -lt 7 -or ($memory.CommitLimit-$memory.CommittedBytes)/1GB -lt 7){throw 'Insufficient RAM/commit reserve'}
}
function Set-ReportTestSetting($book,[string]$key,$value){
    $table=$book.Names.Item('rngSystemSettings').RefersToRange;$data=$table.Value2
    for($r=1;$r -le $data.GetLength(0);$r++){if([string]$data[$r,1] -ceq $key){$table.Cells.Item($r,2).Value2=$value;return}}
    throw ('Setting not found: '+$key)
}
function Get-ReportTestRows($book,[string]$name,[int]$columns,[int]$count=0,[int]$skipColumns=0){
    $anchor=$book.Names.Item($name).RefersToRange
    if(-not $count){$count=$anchor.Worksheet.Cells.Item($anchor.Worksheet.Rows.Count,$anchor.Column).End(-4162).Row-$anchor.Row+1}
    $values=$anchor.Resize($count,$columns).Value2;$result=New-Object Collections.ArrayList
    for($r=1;$r -le $count;$r++){
        if($null -eq $values[$r,1]){continue}
        $item=New-Object Collections.ArrayList
        for($c=$skipColumns+1;$c -le $columns;$c++){[void]$item.Add($values.GetValue($r,$c))}
        [void]$result.Add($item.ToArray())
    }
    return ,$result.ToArray()
}
$sources=if($PublishReport){@('Published')}else{$Labels}
$results=@()
if(-not $PublishReport -and (Test-Path -LiteralPath (Join-Path $PSScriptRoot 'SavedExamples.json'))){
    $previous=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'SavedExamples.json') -Raw -Encoding UTF8|ConvertFrom-Json
    foreach($item in $previous){if($item.Label -notin $Labels -or $item.Case -notin $TestCases){$results+=,$item}}
}
foreach($label in $sources){
    Assert-ReportTestMemory
    $path=if($PublishReport){Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'}else{Join-Path $PSScriptRoot ($label+'.xlsm')}
    $sourceHash=(Get-FileHash -LiteralPath $path).Hash
    $app=$null;$book=$null;$ownedPid=0;$ownedTicks=0
    try{
        $app=New-Object -ComObject Excel.Application
        $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=1
        [uint32]$processId=0;[void][ReportExampleTestProcess]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$processId)
        $ownedPid=[int]$processId;$ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
        Write-Output ('OWNED_EXCEL '+$label+': '+$ownedPid+'; '+$ownedTicks)
        $book=$app.Workbooks.Open($path,0,$true)
        if(-not $book.ReadOnly){throw 'Private report test must be ReadOnly'}
        $loadTable=$book.Names.Item('rngLoadCombinations').RefersToRange
        $torqueTable=$book.Names.Item('rngTorsionCombinations').RefersToRange
        $systemTable=$book.Names.Item('rngSystemSettings').RefersToRange
        $originalLoads=$loadTable.Formula;$originalTorque=$torqueTable.Formula;$originalSettings=$systemTable.Formula
        $cases=if($PublishReport){@('UserExample')}else{$TestCases}
        foreach($case in $cases){
            $loadTable.Formula=$originalLoads;$torqueTable.Formula=$originalTorque;$systemTable.Formula=$originalSettings
            if($case -eq 'TOnlyBlankProfile'){
                $loadTable.Cells.Item(2,5).Value2=''
                $torqueTable.Cells.Item(2,2).Value2=10.0
            }elseif($case -eq 'CompressedNMT'){
                $loadTable.Cells.Item(3,2).Value2=-500.0
                $torqueTable.Cells.Item(3,2).Value2=0.1
            }elseif($case -eq 'SP63Stability'){
                Set-ReportTestSetting $book 'Stability.Code' 'SP63'
            }elseif($case -in @('SP35Width','SP63Width')){
                Set-ReportTestSetting $book 'SLS.Crack.Code' $(if($case -eq 'SP35Width'){'SP35'}else{'SP63'})
                $loadTable.Cells.Item(17,3).Value2=80.0;$loadTable.Cells.Item(17,4).Value2=60.0
            }
            $message=[string]$app.Run("'$($book.Name)'!modWorkbookCalculation.RunSectionCalculationForWorkbook",$book,$false)
            $reportPath=Join-Path $book.Path 'RC_Section_NDM_execution_report.txt'
            if(-not (Test-Path -LiteralPath $reportPath)){throw 'Execution report missing'}
            Copy-Item -LiteralPath $reportPath -Destination (Join-Path $PSScriptRoot ($label+'_'+$case+'.txt'))
            $elements=Get-ReportTestRows $book 'rngNDMElementResults' 9 0 1
            if($elements.Count -lt 465){throw ('Calculated states incomplete: '+$case)}
            $summary=[ordered]@{}
            foreach($pair in @(@('rngStrengthSummaryAnchor',49),@('rngCrackSummaryAnchor',72),@('rngStabilitySummaryAnchor',95),@('rngTorsionSummaryAnchor',48))){
                $summary[$pair[0]]=Get-ReportTestRows $book $pair[0] $pair[1] 30
            }
            $results+=,[ordered]@{Label=$label;Case=$case;Message=$message;SourceSHA256=$sourceHash;Elements=$elements;Summaries=$summary;
                Counters=@($book.Worksheets.Item('Results').Range('G1').Value2,$book.Worksheets.Item('Results').Range('G2').Value2,$book.Worksheets.Item('Results').Range('G3').Value2)}
            Write-Output ('REPORT_CASE_COMPLETE '+$label+' '+$case+'; physical states='+($elements.Count-1))
        }
        $book.Close($false);$book=$null
        $app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app);$app=$null
    }finally{
        if($book){$book.Close($false)}
        if($app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
        [GC]::Collect();[GC]::WaitForPendingFinalizers()
        if($ownedPid){$p=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue;if($p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.ToUniversalTime().Ticks -eq $ownedTicks){Stop-Process -Id $ownedPid -Force}}
    }
    if((Get-FileHash -LiteralPath $path).Hash -ne $sourceHash){throw 'Source workbook changed'}
}
$results|ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $PSScriptRoot $(if($PublishReport){'PublishedRun.json'}else{'SavedExamples.json'})) -Encoding UTF8
Write-Output 'SAVED_REPORT_EXAMPLES_COMPLETE'
